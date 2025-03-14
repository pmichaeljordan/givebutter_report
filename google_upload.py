#!/usr/bin/env python3
import os
import pandas as pd
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

def get_credentials(scopes, token_file='token.json', credentials_file='credentials.json'):
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, scopes)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        creds = flow.run_local_server(port=0)
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    return creds

def upload_to_drive(file_path, folder_id):
    # Use credentials_account1.json for drive upload.
    creds = get_credentials(DRIVE_SCOPES, token_file='shepherd_drive_token.json', credentials_file='credentials_account1.json')
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Uploaded file to Drive with file ID: {file.get('id')}")
    return file.get('id')

def get_all_contacts(service):
    """Retrieve all contacts and return a dict mapping email to contact details."""
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
            emails = person.get('emailAddresses', [])
            for email in emails:
                email_val = email.get('value', '').strip().lower()
                if email_val:
                    contacts[email_val] = person
        page_token = response.get('nextPageToken')
        if not page_token:
            break
    return contacts

def import_to_google_contacts_for_service(csv_path, service):
    """Imports contacts from CSV to a single Google account, ensuring that the contact is
       in the specified group (tag) before creating a new contact.
       This version caches all contacts to reduce API calls.
    """
    import csv

    # Retrieve existing contact groups.
    groups_response = service.contactGroups().list().execute()
    existing_groups = groups_response.get('contactGroups', [])
    group_name_to_resource = {group['name']: group['resourceName'] for group in existing_groups}

    # Retrieve and cache all contacts.
    contacts_cache = get_all_contacts(service)
    
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            contact_email = row["E-mail 1 - Value"].strip().lower()
            group_label = row["Labels"].strip()

            # Ensure the contact group exists.
            if group_label not in group_name_to_resource:
                group_body = {"contactGroup": {"name": group_label}}
                group_resource = None
                try:
                    group_result = service.contactGroups().create(body=group_body).execute()
                    group_resource = group_result.get('resourceName')
                    group_name_to_resource[group_label] = group_resource
                    print(f"Created contact group: {group_label} with resource {group_resource}")
                except HttpError as e:
                    if e.resp.status == 409:
                        # Group already exists; retrieve its resource name.
                        groups_response = service.contactGroups().list().execute()
                        for group in groups_response.get('contactGroups', []):
                            if group.get('name') == group_label:
                                group_resource = group.get('resourceName')
                                group_name_to_resource[group_label] = group_resource
                                print(f"Found existing contact group: {group_label} with resource {group_resource}")
                                break
                        if group_resource is None:
                            raise e
                    else:
                        raise e
            else:
                group_resource = group_name_to_resource[group_label]

            found_contact = contacts_cache.get(contact_email)
            membership_found = False

            if found_contact:
                memberships = found_contact.get("memberships", [])
                for membership in memberships:
                    if membership.get("contactGroupMembership", {}).get("contactGroupResourceName") == group_resource:
                        membership_found = True
                        break

            if found_contact:
                contact_resource = found_contact.get("resourceName")
                if membership_found:
                    print(f"Contact {contact_email} already exists in group '{group_label}'. Skipping.")
                    continue
                else:
                    print(f"Contact {contact_email} exists but is not in group '{group_label}'. Adding to group.")
                    modify_body = {"resourceNamesToAdd": [contact_resource]}
                    service.contactGroups().members().modify(
                        resourceName=group_resource,
                        body=modify_body
                    ).execute()
                    print(f"Added contact {contact_resource} to group '{group_label}'.")
            else:
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
                contacts_cache[contact_email] = result  # update cache
                modify_body = {"resourceNamesToAdd": [contact_resource]}
                service.contactGroups().members().modify(
                    resourceName=group_resource,
                    body=modify_body
                ).execute()
                print(f"Added contact {contact_resource} to group '{group_label}'.")
            
            # Small delay to help avoid rate limits.
            time.sleep(0.2)

def import_to_google_contacts(csv_path):
    """Imports contacts from CSV into two different Google accounts."""
    # Contacts Account 1 uses credentials_account2.json.
    creds_account1 = get_credentials(
        CONTACTS_SCOPES,
        token_file='contacts_token_account1.json',
        credentials_file='credentials_account2.json'
    )
    service_account1 = build('people', 'v1', credentials=creds_account1)

    # Contacts Account 2 uses credentials_account3.json.
    creds_account2 = get_credentials(
        CONTACTS_SCOPES,
        token_file='contacts_token_account2.json',
        credentials_file='credentials_account3.json'
    )
    service_account2 = build('people', 'v1', credentials=creds_account2)

    print("Importing contacts to Account 1...")
    import_to_google_contacts_for_service(csv_path, service_account1)
    print("Importing contacts to Account 2...")
    import_to_google_contacts_for_service(csv_path, service_account2)

def process_mv_sheets(excel_path):
    # Read mapping file if available.
    mapping_file = "data_map.txt"
    try:
        mapping_df = pd.read_csv(mapping_file, header=None, names=["Incorrect", "Correct"])
        mapping_dict = mapping_df.set_index("Incorrect")["Correct"].to_dict()
    except Exception as e:
        mapping_dict = {}

    # Define header template for Google Contacts CSV.
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
    mv_sheets = [sheet for sheet in xl.sheet_names if sheet.startswith('MV')]
    all_rows = []

    for sheet in mv_sheets:
        df = xl.parse(sheet)
        data = df[['First Name', 'Last Name', 'Email']].copy()
        data['Phone'] = df['Phone'] if 'Phone' in df.columns else ''
        data['Tag'] = '2025_Volunteer' if sheet == 'MV Volunteer' else '2025_Rider'

        for _, row in data.iterrows():
            first_name = mapping_dict.get(row["First Name"], row["First Name"]).title()
            last_name  = mapping_dict.get(row["Last Name"], row["Last Name"]).title()
            email      = mapping_dict.get(row["Email"], row["Email"])
            phone      = mapping_dict.get(row["Phone"], row["Phone"])
            tag        = mapping_dict.get(row["Tag"], row["Tag"])

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

