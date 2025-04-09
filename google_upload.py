#!/usr/bin/env python3
import os
import pandas as pd
import csv
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
import time

# Define scopes
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.file']
CONTACTS_SCOPES = ['https://www.googleapis.com/auth/contacts']

# Global sleep duration to slow processing between API calls (in seconds)
API_SLEEP = 1.0

def get_credentials(scopes, token_file='token.json', credentials_file='credentials.json', auth_message=None):
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, scopes)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        if auth_message:
            print(auth_message)
        creds = flow.run_local_server(port=0)
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    return creds

def upload_to_drive(file_path, folder_id):
    creds = get_credentials(
        DRIVE_SCOPES,
        token_file='shepherd_drive_token.json',
        credentials_file='credentials_account1.json',
        auth_message="Authorize for the Shepherd Account"
    )
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': os.path.basename(file_path), 'parents': [folder_id]}
    media = MediaFileUpload(file_path,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Uploaded file to Drive with file ID: {file.get('id')}")
    return file.get('id')

def get_all_contacts(service):
    contacts = {}
    page_token = None
    while True:
        response = service.people().connections().list(
            resourceName='people/me',
            personFields='names,emailAddresses,memberships',
            pageToken=page_token,
            pageSize=200
        ).execute()
        connections = response.get('connections', [])
        for person in connections:
            for email_obj in person.get('emailAddresses', []):
                email_val = email_obj.get('value', '').strip().lower()
                if email_val:
                    contacts[email_val] = person
        page_token = response.get('nextPageToken')
        if not page_token:
            break
    return contacts

def modify_membership(service, resource, body, max_retries=5):
    """Wrapper to modify contact group membership with exponential backoff."""
    delay = 1
    for i in range(max_retries):
        try:
            result = service.contactGroups().members().modify(
                resourceName=resource,
                body=body
            ).execute()
            return result
        except HttpError as e:
            if e.resp.status == 429:
                print(f"Quota exceeded for membership modification. Retrying in {delay} seconds... (Attempt {i+1}/{max_retries})")
                time.sleep(delay)
                delay *= 2
            else:
                raise e
    print("Max retries exceeded for membership modification for resource", resource, "with body", body)
    return None

def import_to_google_contacts_for_service(csv_path, service):
    """
    Deletes the target groups (labels) in Google Contacts and then processes the CSV.
    Each CSV row is processed to add or update a contact and add it to the appropriate group.
    Since the target labels are deleted at the start, the groups are re-created solely based on the CSV.
    """
    # Define the target groups (labels) that will be managed.
    target_groups = ["2025_Rider", "2025_Volunteer"]
    
    # Retrieve existing contact groups.
    groups_response = service.contactGroups().list().execute()
    existing_groups = groups_response.get('contactGroups', [])
    group_name_to_resource = {}
    
    # Delete any existing target groups.
    for group in existing_groups:
        name = group.get('name')
        resource = group.get('resourceName')
        if name in target_groups:
            try:
                print(f"Deleting existing group '{name}' with resource {resource}.")
                service.contactGroups().delete(resourceName=resource).execute()
                time.sleep(API_SLEEP)
            except HttpError as e:
                print(f"Failed to delete group '{name}': {e}")
        else:
            group_name_to_resource[name] = resource

    print("Waiting for groups to be fully deleted...")
    time.sleep(5)
    
    # Read the CSV rows.
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        csv_rows = list(reader)
    
    # Retrieve contacts once.
    contacts_cache = get_all_contacts(service)
    
    # Process each CSV row.
    for row in csv_rows:
        contact_email = row["E-mail 1 - Value"].strip().lower()
        group_label = row["Labels"].strip()
        
        # For target groups, create them fresh if needed.
        if group_label in target_groups:
            if group_label not in group_name_to_resource:
                group_body = {"contactGroup": {"name": group_label}}
                try:
                    group_result = service.contactGroups().create(body=group_body).execute()
                    group_resource = group_result.get('resourceName')
                    group_name_to_resource[group_label] = group_resource
                    print(f"Created contact group: {group_label} with resource {group_resource}")
                    time.sleep(API_SLEEP)
                except HttpError as e:
                    if e.resp.status == 409:
                        # The group already exists; retrieve it.
                        groups_response = service.contactGroups().list().execute()
                        for group in groups_response.get('contactGroups', []):
                            if group.get('name') == group_label:
                                group_resource = group.get('resourceName')
                                group_name_to_resource[group_label] = group_resource
                                print(f"Found existing contact group: {group_label} with resource {group_resource} (after conflict)")
                                break
                        else:
                            raise e
                    else:
                        raise e
            else:
                group_resource = group_name_to_resource[group_label]
        else:
            # For non-target groups, create as needed.
            if group_label not in group_name_to_resource:
                group_body = {"contactGroup": {"name": group_label}}
                group_result = service.contactGroups().create(body=group_body).execute()
                group_resource = group_result.get('resourceName')
                group_name_to_resource[group_label] = group_resource
                print(f"Created contact group: {group_label} with resource {group_resource}")
                time.sleep(API_SLEEP)
            else:
                group_resource = group_name_to_resource[group_label]

        # Check if the contact already exists.
        found_contact = contacts_cache.get(contact_email)
        if found_contact:
            contact_resource = found_contact.get("resourceName")
            # Check if the contact is already a member of the target group.
            membership_found = any(
                membership.get("contactGroupMembership", {}).get("contactGroupResourceName") == group_resource
                for membership in found_contact.get("memberships", [])
            )
            if membership_found:
                print(f"Contact {contact_email} already exists in group '{group_label}'. Skipping.")
            else:
                print(f"Contact {contact_email} exists but is not in group '{group_label}'. Adding to group.")
                modify_membership(service, group_resource, {"resourceNamesToAdd": [contact_resource]})
                print(f"Added contact {contact_email} to group '{group_label}'.")
        else:
            # Create a new contact.
            contact_body = {
                "names": [{
                    "givenName": row["First Name"].strip(),
                    "familyName": row["Last Name"].strip()
                }],
                "emailAddresses": [{
                    "value": contact_email
                }],
                "phoneNumbers": [{
                    "value": row["Phone 1 - Value"].strip()
                }]
            }
            result = service.people().createContact(body=contact_body).execute()
            contact_resource = result.get('resourceName')
            print(f"Created contact: {contact_resource}")
            modify_membership(service, group_resource, {"resourceNamesToAdd": [contact_resource]})
            print(f"Added contact {contact_resource} to group '{group_label}'.")
            # Update the local contacts cache.
            contacts_cache[contact_email] = result
        
        time.sleep(API_SLEEP)

def import_to_google_contacts(csv_path):
    creds_account1 = get_credentials(
        CONTACTS_SCOPES,
        token_file='contacts_token_account1.json',
        credentials_file='credentials_account2.json',
        auth_message="Authorize for the Shepherd Account"
    )
    service_account1 = build('people', 'v1', credentials=creds_account1)
    creds_account2 = get_credentials(
        CONTACTS_SCOPES,
        token_file='contacts_token_account2.json',
        credentials_file='credentials_account3.json',
        auth_message="Authorize for the Admin account"
    )
    service_account2 = build('people', 'v1', credentials=creds_account2)
    
    print("Importing contacts to Account 1...")
    import_to_google_contacts_for_service(csv_path, service_account1)
    print("Importing contacts to Account 2...")
    import_to_google_contacts_for_service(csv_path, service_account2)

def process_mv_sheets(excel_path):
    try:
        # Read the mapping file with columns: Ticket Number, Incorrect, Correct.
        mapping_df = pd.read_csv("data_map.txt", header=None, names=["Ticket Number", "Incorrect", "Correct"])
        # Build a mapping dictionary keyed by the ticket number.
        mapping_dict = {
            str(row["Ticket Number"]).strip(): {
                "Incorrect": str(row["Incorrect"]).strip(),
                "Correct": str(row["Correct"]).strip()
            }
            for index, row in mapping_df.iterrows()
        }
    except Exception as e:
        print("Error reading data_map.txt:", e)
        mapping_dict = {}

    google_columns = [
        'Name Prefix', 'First Name', 'Middle Name', 'Last Name', 'Name Suffix',
        'Phonetic First Name', 'Phonetic Middle Name', 'Phonetic Last Name',
        'Nickname', 'File As', 'E-mail 1 - Label', 'E-mail 1 - Value',
        'Phone 1 - Label', 'Phone 1 - Value', 'Address 1 - Label', 'Address 1 - Country',
        'Address 1 - Street', 'Address 1 - Extended Address', 'Address 1 - City',
        'Address 1 - Region', 'Address 1 - Postal Code', 'Address 1 - PO Box',
        'Organization Name', 'Organization Title', 'Organization Department', 'Birthday',
        'Event 1 - Label', 'Event 1 - Value', 'Relation 1 - Label', 'Relation 1 - Value',
        'Website 1 - Label', 'Website 1 - Value', 'Custom Field 1 - Label',
        'Custom Field 1 - Value', 'Notes', 'Labels'
    ]
    
    xl = pd.ExcelFile(excel_path)
    # Expect sheets whose names begin with 'MV'
    mv_sheets = [sheet for sheet in xl.sheet_names if sheet.startswith('MV')]
    all_rows = []
    
    for sheet in mv_sheets:
        # Expect the Excel file to have a "Ticket Number" column as the first column
        df = xl.parse(sheet)
        # Limit the columns: Ticket Number, First Name, Last Name, Email.
        data = df[['Ticket Number', 'First Name', 'Last Name', 'Email']].copy()
        # Use "Phone" column if it exists.
        if 'Phone' in df.columns:
            data['Phone'] = df['Phone']
        else:
            data['Phone'] = ''
        # Set the tag based on sheet name.
        data['Tag'] = '2025_Volunteer' if sheet == 'MV Volunteer' else '2025_Rider'
        
        for _, row in data.iterrows():
            ticket_number = str(row["Ticket Number"]).strip()
            first_name = str(row["First Name"]).strip().title()
            last_name  = str(row["Last Name"]).strip().title()
            orig_email = str(row["Email"]).strip()
            # Replace email if ticket number is in mapping and the email matches the "Incorrect" value.
            if ticket_number in mapping_dict and orig_email == mapping_dict[ticket_number]["Incorrect"]:
                email = mapping_dict[ticket_number]["Correct"]
            else:
                email = orig_email
            phone = str(row["Phone"]).strip()
            tag = str(row["Tag"]).strip()
            
            contact = {col: "" for col in google_columns}
            contact["First Name"] = first_name
            contact["Last Name"] = last_name
            contact["E-mail 1 - Value"] = email
            contact["Phone 1 - Value"] = phone
            contact["Labels"] = tag
            contact["E-mail 1 - Label"] = "Email"
            contact["Phone 1 - Label"] = "Phone"
            all_rows.append(contact)
    
    final_df = pd.DataFrame(all_rows, columns=google_columns)
    final_csv = 'output.csv'
    final_df.to_csv(final_csv, index=False, header=True)
    print(f"Google Contacts CSV saved to {final_csv}")
    return final_csv

if __name__ == '__main__':
    print("This module provides Google upload functions.")

