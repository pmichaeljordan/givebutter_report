#!/usr/bin/env python3
import os
import requests
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

# ----------------------------
# Configuration from exported environment variables
# ----------------------------
WIX_APP_ID = os.environ.get("WIX_APP_ID")
WIX_APP_SECRET = os.environ.get("WIX_APP_SECRET")

if not WIX_APP_ID:
    raise ValueError("WIX_APP_ID is not set in your environment.")
if not WIX_APP_SECRET:
    raise ValueError("WIX_APP_SECRET is not set in your environment.")

# Wix API endpoints
WIX_APP_INSTANCE_URL = "https://www.wixapis.com/apps/v1/instance"   # For retrieving app instance info
WIX_OAUTH_TOKEN_URL = "https://www.wixapis.com/oauth2/token"         # For obtaining an OAuth token
WIX_CONTACTS_URL = "https://www.wixapis.com/contacts/v4/contacts"      # Contacts endpoint, v4

# Google People API configuration
SCOPES = ['https://www.googleapis.com/auth/contacts']
TOKEN_FILE = "token.json"  # Contains your Google OAuth token

# ----------------------------
# Functions for Wix API: Instance & OAuth Token
# ----------------------------
def get_app_instance():
    """
    Retrieves your app instance information from Wix.
    
    This endpoint returns data about the instance of your app that’s installed 
    on a Wix site (including the instance ID needed for the OAuth token request).
    
    Note: For this call to succeed, your app must be installed on a Wix site. 
    You must authenticate this call as a Wix app. Here we assume API key authentication 
    with a Bearer token works—but check your docs as needed.
    """
    headers = {
        "Content-Type": "application/json",
        # Depending on your setup, you might need a Bearer prefix here.
        "Authorization": f"Bearer {WIX_APP_ID}"  
    }
    try:
        response = requests.get(WIX_APP_INSTANCE_URL, headers=headers)
        if response.status_code == 200:
            instance_info = response.json()
            # Inspect the returned JSON to locate the instance ID.
            # It might be under a field named "instanceId", "id", or similar.
            instance_id = instance_info.get("instanceId") or instance_info.get("id")
            if instance_id:
                print(f"Retrieved app instance ID: {instance_id}")
                return instance_id
            else:
                raise Exception("Instance ID not found in the response.")
        else:
            raise Exception(f"Failed to retrieve instance info: {response.status_code} {response.text}")
    except Exception as e:
        raise Exception(f"Error calling /apps/v1/instance: {e}")

def get_wix_oauth_token():
    """
    Obtains an OAuth access token from Wix using the client_credentials flow.
    
    The payload includes:
      - grant_type: "client_credentials"
      - client_id: Your Wix App ID
      - client_secret: Your Wix App Secret
      - instance_id: The valid instance ID obtained from get_app_instance()
    """
    try:
        instance_id = get_app_instance()
    except Exception as e:
        raise Exception(f"Unable to get app instance ID: {e}")
    
    payload = {
        "grant_type": "client_credentials",
        "client_id": WIX_APP_ID,
        "client_secret": WIX_APP_SECRET,
        "instance_id": instance_id
    }
    headers = {"Content-Type": "application/json"}
    response = requests.post(WIX_OAUTH_TOKEN_URL, json=payload, headers=headers)
    if response.status_code == 200:
        token_data = response.json()
        access_token = token_data.get("access_token")
        if not access_token:
            raise Exception("No access token returned in the OAuth response.")
        print(f"Obtained OAuth token: {access_token[:10]}...")  # Print a snippet for debugging
        return access_token
    else:
        raise Exception(f"Failed to obtain Wix OAuth token: {response.status_code} {response.text}")

def get_wix_contacts():
    """
    Fetches contacts from the Wix Contacts API using the OAuth token.
    
    According to the docs, the access token is placed directly in the Authorization header.
    """
    try:
        access_token = get_wix_oauth_token()
    except Exception as e:
        print("Error obtaining Wix OAuth token:", e)
        return []
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": access_token  # Per docs, the token is sent without a "Bearer" prefix
    }
    
    try:
        response = requests.get(WIX_CONTACTS_URL, headers=headers)
        if response.status_code == 200:
            data = response.json()
            contacts = data.get("contacts", [])
            print(f"Fetched {len(contacts)} contact(s) from Wix.")
            return contacts
        else:
            print(f"Error fetching Wix contacts: {response.status_code} {response.text}")
            return []
    except Exception as e:
        print("Exception while fetching Wix contacts:", e)
        return []

# ----------------------------
# Functions for Google People API integration
# ----------------------------
def get_google_service():
    """
    Builds and returns a Google People API service object.
    
    Assumes a valid OAuth token is stored in TOKEN_FILE.
    """
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    service = build('people', 'v1', credentials=creds)
    return service

def find_google_contact(service, email):
    """
    Searches for an existing Google contact by email.
    """
    try:
        results = service.people().searchContacts(
            query=email,
            readMask="names,emailAddresses",
            pageSize=1
        ).execute()
        contacts = results.get("results", [])
        if contacts:
            return contacts[0].get("person")
        return None
    except Exception as e:
        print(f"Error searching for Google contact with email {email}: {e}")
        return None

def create_google_contact(service, wix_contact):
    """
    Creates a new Google contact using data from a Wix contact.
    """
    contact_body = {
        "names": [{
            "givenName": wix_contact.get("firstName", ""),
            "familyName": wix_contact.get("lastName", "")
        }],
        "emailAddresses": [{
            "value": wix_contact.get("email", "")
        }]
    }
    try:
        new_contact = service.people().createContact(body=contact_body).execute()
        print(f"Created contact for {wix_contact.get('email')}")
        return new_contact
    except Exception as e:
        print(f"Error creating Google contact for {wix_contact.get('email')}: {e}")
        return None

def update_google_contact(service, google_contact, wix_contact):
    """
    Updates an existing Google contact with data from the Wix contact.
    """
    resource_name = google_contact.get("resourceName")
    contact_body = {
        "names": [{
            "givenName": wix_contact.get("firstName", ""),
            "familyName": wix_contact.get("lastName", "")
        }],
        "emailAddresses": [{
            "value": wix_contact.get("email", "")
        }]
    }
    try:
        updated_contact = service.people().updateContact(
            resourceName=resource_name,
            updatePersonFields="names,emailAddresses",
            body=contact_body
        ).execute()
        print(f"Updated contact for {wix_contact.get('email')}")
        return updated_contact
    except Exception as e:
        print(f"Error updating Google contact for {wix_contact.get('email')}: {e}")
        return None

def sync_contacts():
    """
    Synchronizes contacts from Wix to Google.
    
    For each Wix contact:
      - If the contact exists in Google, updates it.
      - Otherwise, creates a new contact.
    """
    wix_contacts = get_wix_contacts()
    if not wix_contacts:
        print("No contacts found in Wix to sync.")
        return

    # google_service = get_google_service()
    
    # for wix_contact in wix_contacts:
    #     email = wix_contact.get("email")
    #     if not email:
    #         print("Wix contact does not have an email; skipping record.")
    #         continue
    #     google_contact = find_google_contact(google_service, email)
    #     if google_contact:
    #         update_google_contact(google_service, google_contact, wix_contact)
    #     else:
    #         create_google_contact(google_service, wix_contact)

def main():
    sync_contacts()

if __name__ == "__main__":
    main()
