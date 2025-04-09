#!/usr/bin/env python3
import pandas as pd
import glob
import os
import re
from datetime import datetime

def parse_tickets():
    # Find the latest CSV file with prefix 'tickets-'
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
        return None

    # Read the CSV with specific parameters
    data = pd.read_csv(
        file_path,
        sep=",",
        quotechar='"',
        encoding="utf-8-sig"
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

    # Clean column names
    data.columns = data.columns.str.strip().str.strip('"')
    print("Columns after cleaning:")
    print(data.columns.tolist())
    
    # Filter out revoked tickets if the column exists.
    if 'Ticket Revoked' in data.columns:
        initial_shape = data.shape
        # Remove rows where Ticket Revoked is TRUE (case insensitive)
        data = data[~(data["Ticket Revoked"].astype(str).str.upper() == "TRUE")]
        print(f"Filtered revoked tickets: {initial_shape} -> {data.shape}")

    if 'Ticket Type' not in data.columns:
        raise ValueError("Could not find 'Ticket Type' in the CSV columns.")

    ticket_types = data['Ticket Type'].unique()
    print("Unique ticket types:", ticket_types)

    date_str = datetime.now().strftime("%m-%d-%Y")
    output_path = f"./parsed_tickets_{date_str}.xlsx"
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    def clean_sheet_name(ticket_type):
        parts = ticket_type.split(' - ', 1)
        new_name = parts[1].strip() if len(parts) == 2 else ticket_type.strip()
        new_name = re.sub(r' / ', ' ', new_name)
        invalid_chars = r'[\[\]\:\*\?\\\/]'
        new_name = re.sub(invalid_chars, '', new_name)
        return new_name[:31]

    for ticket_type in ticket_types:
        sheet_name = clean_sheet_name(ticket_type)
        ticket_data = data[data['Ticket Type'] == ticket_type]
        ticket_data = ticket_data.dropna(axis=1, how='all')
        ticket_data.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for i, col in enumerate(ticket_data.columns):
            max_len = max(ticket_data[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

    # Create a T-shirt Sizes sheet
    tshirt_data = data[['T-shirt sizing (Unisex)', 'First Name', 'Last Name', 'Email']]
    tshirt_summary = tshirt_data['T-shirt sizing (Unisex)'].value_counts().reset_index()
    tshirt_summary.columns = ['T-Shirt Size', 'Count']
    tshirt_data.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=0)
    tshirt_summary.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=len(tshirt_data) + 2)

    writer.close()
    print(f"Workbook saved to {output_path}")

    return output_path

if __name__ == '__main__':
    parse_tickets()

