#!/usr/bin/env python3
import pandas as pd
import glob
import os
import re
from datetime import datetime

def find_latest_csv(prefix='tickets-', directory='.'):
    files = glob.glob(os.path.join(directory, f"{prefix}*.csv"))
    if files:
        return max(files, key=os.path.getmtime)
    else:
        raise FileNotFoundError(f"No CSV files starting with '{prefix}' found in {directory}")

try:
    file_path = find_latest_csv(prefix='tickets-', directory='.')
    print(f"Processing file: {file_path}")
except FileNotFoundError as e:
    print(str(e))
    exit()

# -- Read CSV with extra parameters + debug prints
data = pd.read_csv(
    file_path,
    sep=",",
    quotechar='"',
    encoding="utf-8-sig",  # BOM-safe
    # engine="python",     # Uncomment if you have parsing errors with the default engine
)
cols_to_drop = [
    'Ticket Suffix',
    'Campaign Code',
    'Campaign Title',
    'Price',
    'Check In',
    'Promo Code',
    'Checked in by',
    'Checkin type',
    'Checkin source',
    'Check-in Date (UTC)',
    'Bundled',
    'Bundle Type'
]
data.drop(columns=cols_to_drop, errors='ignore', inplace=True)
print("Data shape (rows, cols):", data.shape)
print("Columns found:")
for c in data.columns:
    print(f"  '{c}'")

# Optional: strip quotes/spaces from column names
data.columns = data.columns.str.strip().str.strip('"')

# Re-check columns
print("Columns after cleaning:")
print(data.columns.tolist())

# Now confirm you see 'Ticket Type' in data.columns
if 'Ticket Type' not in data.columns:
    raise ValueError("Could not find 'Ticket Type' in the CSV columns.")

# Ensure all rows are processed
ticket_types = data['Ticket Type'].unique()
print("Unique ticket types:", ticket_types)

date_str = datetime.now().strftime("%m-%d-%Y")
output_path = f"./parsed_tickets_{date_str}.xlsx"
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

def clean_sheet_name(ticket_type):
    # 1) Strip off everything before " - "
    parts = ticket_type.split(' - ', 1)
    if len(parts) == 2:
        new_name = parts[1].strip()
    else:
        new_name = ticket_type.strip()
    
    # 2) Remove (or replace) invalid Excel characters
    # Excel disallows: []:*?/\
    new_name = re.sub(r' / ', ' ', new_name)
    invalid_chars = r'[\[\]\:\*\?\\\/]'
    new_name = re.sub(invalid_chars, '', new_name)
    
    # 3) Truncate to max 31 characters for Excel
    return new_name[:31]

for ticket_type in ticket_types:
    sheet_name = clean_sheet_name(ticket_type)
    ticket_data = data[data['Ticket Type'] == ticket_type]
    # Drop columns if they are empty in this subset
    ticket_data = ticket_data.dropna(axis=1, how='all')
    ticket_data.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for i, col in enumerate(ticket_data.columns):
        max_len = max(ticket_data[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

# T-shirt sizes
tshirt_data = data[['T-shirt sizing (Unisex)', 'First Name', 'Last Name', 'Email']]

tshirt_summary = tshirt_data['T-shirt sizing (Unisex)'].value_counts().reset_index()
tshirt_summary.columns = ['T-Shirt Size', 'Count']

# Write T-shirt data and summary to a sheet
tshirt_data.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=0)
tshirt_summary.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=len(tshirt_data) + 2)

writer.close()

print(f"Workbook saved to {output_path}")

def process_mv_sheets():

    # Define the Google Contacts template columns in order
    google_columns = [
        'Name Prefix',
        'First Name',
        'Middle Name',
        'Last Name',
        'Name Suffix',
        'Phonetic First Name',
        'Phonetic Middle Name',
        'Phonetic Last Name',
        'Nickname',
        'File As',
        'E-mail 1 - Label',
        'E-mail 1 - Value',
        'Phone 1 - Label',
        'Phone 1 - Value',
        'Address 1 - Label',
        'Address 1 - Country',
        'Address 1 - Street',
        'Address 1 - Extended Address',
        'Address 1 - City',
        'Address 1 - Region',
        'Address 1 - Postal Code',
        'Address 1 - PO Box',
        'Organization Name',
        'Organization Title',
        'Organization Department',
        'Birthday',
        'Event 1 - Label',
        'Event 1 - Value',
        'Relation 1 - Label',
        'Relation 1 - Value',
        'Website 1 - Label',
        'Website 1 - Value',
        'Custom Field 1 - Label',
        'Custom Field 1 - Value',
        'Notes',
        'Labels'
    ]

    # Read the Excel file
    file_path = output_path  # ensure output_path is defined elsewhere in your script
    xl = pd.ExcelFile(file_path)

    # Extract sheet names that start with "MV"
    mv_sheets = [sheet for sheet in xl.sheet_names if sheet.startswith('MV')]
    all_rows = []

    for sheet in mv_sheets:
        df = xl.parse(sheet)

        # Create a copy of the DataFrame with the base columns
        data = df[['First Name', 'Last Name', 'Email']].copy()

        # Add Phone column if it exists; otherwise, use empty strings
        if 'Phone' in df.columns:
            data['Phone'] = df['Phone']
        else:
            data['Phone'] = ''

        # Add Tag column based on the sheet name
        if sheet == 'MV Volunteer':
            data['Tag'] = '2025_Volunteer'
        else:
            data['Tag'] = '2025_Rider'

        # For each row, create a new dictionary matching the Google Contacts columns
        for _, row in data.iterrows():
            contact = {col: "" for col in google_columns}
            contact["First Name"] = row["First Name"]
            contact["Last Name"] = row["Last Name"]
            contact["E-mail 1 - Value"] = row["Email"]
            contact["Phone 1 - Value"] = row["Phone"]
            contact["Labels"] = row["Tag"]

            # Optionally, set default labels for email and phone
            contact["E-mail 1 - Label"] = "Email"
            contact["Phone 1 - Label"] = "Phone"

            all_rows.append(contact)

    # Create a final DataFrame with the Google Contacts header order
    final_df = pd.DataFrame(all_rows, columns=google_columns)
    final_df.to_csv('output.csv', index=False, header=True)
process_mv_sheets()
