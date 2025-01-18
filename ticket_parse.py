#!/usr/bin/env python3
import pandas as pd
import glob
import os
import re

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

output_path = './parsed_tickets.xlsx'
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
    invalid_chars = r'[\[\]\:\*\?\\\/]'
    new_name = re.sub(invalid_chars, '', new_name)
    
    # 3) Truncate to max 31 characters for Excel
    return new_name[:31]

for ticket_type in ticket_types:
    sheet_name = clean_sheet_name(ticket_type)
    ticket_data = data[data['Ticket Type'] == ticket_type]
    ticket_data.to_excel(writer, sheet_name=sheet_name, index=False)

# T-shirt sizes
tshirt_data = data[['T-shirt sizing (Unisex)', 'First Name', 'Last Name', 'Email']]

tshirt_summary = tshirt_data['T-shirt sizing (Unisex)'].value_counts().reset_index()
tshirt_summary.columns = ['T-Shirt Size', 'Count']

# Write T-shirt data and summary to a sheet
tshirt_data.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=0)
tshirt_summary.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=len(tshirt_data) + 2)

writer.close()

print(f"Workbook saved to {output_path}")

