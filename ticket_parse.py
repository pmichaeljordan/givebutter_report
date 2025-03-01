#!/usr/bin/env python3
import os
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# If modifying these scopes, delete your previously saved token files.
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.file']
CONTACTS_SCOPES = ['https://www.googleapis.com/auth/contacts']

# Define output_path at the global level
date_str = datetime.now().strftime("%m-%d-%Y")
output_path = f"./parsed_tickets_{date_str}.xlsx"

def get_credentials(scopes, token_file='token.json', credentials_file='credentials.json'):
    """
    Obtains credentials for a given scope.
    You must create credentials.json from the Google Cloud Console.
    """
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, scopes)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    return creds

def upload_to_drive(file_path, folder_id):
    """Uploads a file to Google Drive inside the specified folder."""
    creds = get_credentials(DRIVE_SCOPES, token_file='drive_token.json')
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    # Set the MIME type based on your file; here assuming an Excel file.
    media = MediaFileUpload(file_path,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Uploaded file to Drive with file ID: {file.get('id')}")
    return file.get('id')

def import_to_google_contacts(csv_path):
    import csv
    from googleapiclient.discovery import build
    # Obtain credentials for the People API
    creds = get_credentials(CONTACTS_SCOPES, token_file='contacts_token.json')
    service = build('people', 'v1', credentials=creds)
    
    # Retrieve existing contact groups once and cache them in a dictionary:
    groups_response = service.contactGroups().list().execute()
    existing_groups = groups_response.get('contactGroups', [])
    group_name_to_resource = {group['name']: group['resourceName'] for group in existing_groups}
    
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            # Build the contact body using fields from your CSV.
            contact_body = {
                "names": [{
                    "givenName": row["First Name"],
                    "familyName": row["Last Name"]
                }],
                "emailAddresses": [{
                    "value": row["E-mail 1 - Value"]
                }],
                "phoneNumbers": [{
                    "value": row["Phone 1 - Value"]
                }]
            }
            # Create the contact
            result = service.people().createContact(body=contact_body).execute()
            contact_resource = result.get('resourceName')
            print(f"Created contact: {contact_resource}")
            
            # Get the label from the CSV (which you want to use as a contact group)
            label = row["Labels"].strip()
            # If the group doesn't exist, create it.
            if label not in group_name_to_resource:
                group_body = {"contactGroup": {"name": label}}
                group_result = service.contactGroups().create(body=group_body).execute()
                group_resource = group_result.get('resourceName')
                group_name_to_resource[label] = group_resource
                print(f"Created contact group: {label} with resource {group_resource}")
            else:
                group_resource = group_name_to_resource[label]
            
            # Add the contact to the contact group.
            # This uses the People API endpoint to modify group members.
            modify_body = {
                "resourceNamesToAdd": [contact_resource]
            }
            service.contactGroups().members().modify(
                resourceName=group_resource,
                body=modify_body
            ).execute()
            print(f"Added contact {contact_resource} to group {label}")

def process_mv_sheets(output_path):
    # --- Read the mapping file ---
    mapping_file = "data_map.txt"
    try:
        mapping_df = pd.read_csv(mapping_file, header=None, names=["Incorrect", "Correct"])
        mapping_dict = mapping_df.set_index("Incorrect")["Correct"].to_dict()
    except Exception as e:
        mapping_dict = {}

    # Define Google Contacts header template
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

    xl = pd.ExcelFile(output_path)
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
    # Your script already generates an Excel file; assume output_path is defined.
    # For example, after processing tickets:
    print("Workbook saved to", output_path)

    # Upload the Excel file to a Google Drive folder.
    # Replace 'your_folder_id_here' with the actual folder ID on your Google Drive.
    drive_folder_id = '11Jsj1pPf7NWYdVCzmuTrjTWj80JObSa4'
    upload_to_drive(output_path, drive_folder_id)

    # Process the MV sheets and create output.csv for Google Contacts.
    csv_file = process_mv_sheets(output_path)

    # Import contacts to Google Contacts using the People API.
    import_to_google_contacts(csv_file)

