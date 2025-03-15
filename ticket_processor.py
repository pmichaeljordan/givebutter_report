#!/usr/bin/env python3

import ticket_parse
import google_upload

def main():
    # Step 1: Run the ticket parsing (from ticket_parse.py)
    output_excel = ticket_parse.parse_tickets()
    if not output_excel:
        print("Ticket parsing failed. Exiting.")
        return

    # Step 2: Use google_upload.py functions to handle Google Drive and Contacts.
    # Upload the Excel file to Google Drive.
    drive_folder_id = '11Jsj1pPf7NWYdVCzmuTrjTWj80JObSa4'  # Update this as needed.
    google_upload.upload_to_drive(output_excel, drive_folder_id)

    # Process the MV sheets from the Excel file to generate the contacts CSV.
    csv_file = google_upload.process_mv_sheets(output_excel)

    # Import contacts into Google Contacts.
    google_upload.import_to_google_contacts(csv_file)

if __name__ == '__main__':
    main()

