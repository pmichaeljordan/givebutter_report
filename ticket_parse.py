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

        # Check file age and warn if old
        file_mod_time = os.path.getmtime(file_path)
        file_age_days = (datetime.now().timestamp() - file_mod_time) / (24 * 3600)
        file_mod_date = datetime.fromtimestamp(file_mod_time).strftime('%Y-%m-%d %H:%M:%S')

        print(f"Processing file: {file_path}")
        print(f"Last modified: {file_mod_date}")

        if file_age_days > 7:
            print(f"WARNING: This file is {file_age_days:.1f} days old.")
            response = input("Do you want to proceed with this file? (yes/no): ").strip().lower()
            if response not in ['yes', 'y']:
                print("Please export a fresh tickets CSV from Givebutter and try again.")
                return None

    except FileNotFoundError as e:
        print(str(e))
        print("Please export the tickets CSV from Givebutter and place it in this directory.")
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

    # Build registration summary counts for the pie chart
    # There are duplicate "Are you a new or returning rider?" columns (pandas appends .1, .2...)
    # Regular riders use the first occurrence; High School riders use .1 — coalesce both.
    nr_cols = [c for c in data.columns if c.startswith('Are you a new or returning rider?')]
    if nr_cols:
        new_ret_combined = data[nr_cols].bfill(axis=1).iloc[:, 0]
    else:
        new_ret_combined = pd.Series('', index=data.index)

    volunteer_mask   = data['Ticket Type'].str.contains('Volunteer',   case=False, na=False)
    reciprocal_mask  = data['Ticket Type'].str.contains('Reciprocal',  case=False, na=False)
    high_school_mask = data['Ticket Type'].str.contains('High School', case=False, na=False)
    regular_mask     = ~volunteer_mask & ~reciprocal_mask & ~high_school_mask

    new_mask      = new_ret_combined.str.contains('New',       case=False, na=False)
    returning_mask = new_ret_combined.str.contains('Returning', case=False, na=False)

    counts = {
        'New Riders':            int((regular_mask & new_mask).sum()),
        'Returning Riders':      int((regular_mask & returning_mask).sum()),
        'Reciprocal':            int(reciprocal_mask.sum()),
        'High School New':       int((high_school_mask & new_mask).sum()),
        'High School Returning': int((high_school_mask & returning_mask).sum()),
        'Volunteer':             int(volunteer_mask.sum()),
    }
    total_riders = sum(v for k, v in counts.items() if k != 'Volunteer')
    summary_counts = list(counts.items())
    summary_df = pd.DataFrame(summary_counts, columns=['Category', 'Count'])

    date_str = datetime.now().strftime("%m-%d-%Y")
    output_path = f"./parsed_tickets_{date_str}.xlsx"
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    # Write Summary sheet first so it appears as the first tab
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    workbook = writer.book
    summary_ws = writer.sheets['Summary']
    summary_ws.set_column(0, 0, 22)
    summary_ws.set_column(1, 1, 8)

    chart = workbook.add_chart({'type': 'pie'})
    chart.add_series({
        'name':       'Registrations',
        'categories': ['Summary', 1, 0, len(summary_df), 0],
        'values':     ['Summary', 1, 1, len(summary_df), 1],
        'data_labels': {'percentage': True, 'category': True, 'separator': '\n'},
    })
    summary_ws.write(len(summary_df) + 2, 0, 'Total Riders (excl. Volunteers)')
    summary_ws.write(len(summary_df) + 2, 1, total_riders)

    chart.set_title({'name': f'Registration Breakdown by Type ({total_riders} total riders)'})
    chart.set_style(10)
    summary_ws.insert_chart('D2', chart, {'x_scale': 1.8, 'y_scale': 1.8})
    print("Summary sheet with pie chart created")

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

    # Create a T-shirt Sizes sheet if the column exists (match prefix to handle event-suffixed column names)
    tshirt_col = next((c for c in data.columns if c.startswith('T-shirt sizing (Unisex)')), None)
    if tshirt_col:
        tshirt_columns = [tshirt_col, 'First Name', 'Last Name', 'Email']
        available_tshirt_cols = [col for col in tshirt_columns if col in data.columns]

        if available_tshirt_cols:
            tshirt_data = data[available_tshirt_cols]
            tshirt_summary = tshirt_data[tshirt_col].value_counts().reset_index()
            tshirt_summary.columns = ['T-Shirt Size', 'Count']
            tshirt_data.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=0)
            tshirt_summary.to_excel(writer, sheet_name='T-Shirt Sizes', index=False, startrow=len(tshirt_data) + 2)
            print("T-Shirt Sizes sheet created")
    else:
        print("T-shirt sizing column not found, skipping T-Shirt Sizes sheet")

    writer.close()
    print(f"Workbook saved to {output_path}")

    return output_path

if __name__ == '__main__':
    parse_tickets()

