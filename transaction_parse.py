#!/usr/bin/env python3
"""
Export transactions with detail from Givegbutter.

This script processes the transactions CSV by:
  - Standardizing column names.
  - Correcting a campaign description (if needed) in the 'Item Description' column.
  - Filtering for ticket transactions (where 'Item Subtype' is "ticket").
  - Dropping unnecessary columns.
  - Optionally correcting team member email addresses using a mapping CSV, if available.
  - Splitting the "Team Member" column into first and last names for comparison.
  - Saving CSV subsets (one per campaign) and a master Excel file with separate sheets.
"""

import pandas as pd
import os
import glob
import time

def auto_adjust_columns_width(df, writer, sheet_name):
    """Auto-adjust column widths for an Excel sheet."""
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)

def save_and_compare_df(df, description, output_dir):
    """
    Save a DataFrame to CSV with a filename based on the description and current time.
    If a previous file exists, compare and output any new rows.
    """
    filename = f"{description.replace(' ', '_')}_{int(time.time())}.csv"
    filepath = os.path.join(output_dir, filename)
    df.to_csv(filepath, index=False)
    
    # Find previous files for this description (exclude the current file)
    previous_files = [f for f in glob.glob(os.path.join(output_dir, f"{description.replace(' ', '_')}*.csv"))
                      if f != filepath]
    
    if previous_files:
        previous_file = max(previous_files, key=os.path.getctime)
        previous_df = pd.read_csv(previous_file)
        # This simplistic comparison looks for new rows
        comparison_df = pd.concat([df, previous_df, previous_df]).drop_duplicates(keep=False)
        if not comparison_df.empty:
            changes_filename = f"changes_{filename}"
            changes_filepath = os.path.join(output_dir, changes_filename)
            comparison_df.to_csv(changes_filepath, index=False)
            print(f"Changes saved to {changes_filepath}")
        else:
            print("No changes detected.")
    else:
        print(f"No previous files found for {description}. Current file saved as {filename}.")

def main():
    # Load the transactions CSV and strip whitespace from column names
    file_path = 'transactions.csv'
    df = pd.read_csv(file_path)
    df.columns = df.columns.str.strip()

    # Correct campaign description if needed in the 'Item Description' column.
    df['Item Description'] = df['Item Description'].replace(
        '2025 Ride for Missing Children - MV New and Returning Riders',
        '2025 Ride for Missing Children - MV New / Returning Riders'
    )
    
    # Filter rows where 'Item Subtype' is 'ticket'
    tickets_df = df[df['Item Subtype'] == 'ticket'].copy()
    
    # Drop columns that are not needed for processing
    initial_drop_columns = ["Campaign Title", "Campaign Slug", "Team", "Reference Number"]
    tickets_df.drop(columns=initial_drop_columns, errors='ignore', inplace=True)
    
    # Try to load the name mapping CSV if it exists and is not empty
    mapping_file = 'name_mapping.csv'
    df_mapping = None
    if os.path.exists(mapping_file) and os.path.getsize(mapping_file) > 0:
        try:
            temp_mapping = pd.read_csv(mapping_file)
            if not temp_mapping.empty:
                df_mapping = temp_mapping.copy()
                df_mapping.columns = df_mapping.columns.str.strip()
        except pd.errors.EmptyDataError:
            print("Mapping file is empty. Skipping email correction.")
            df_mapping = None
    else:
        print("Mapping file not found or is empty. Skipping email correction.")
    
    # Split the 'Team Member' column into first and last names for comparison
    if 'Team Member' in tickets_df.columns and 'First Name' in tickets_df.columns:
        tickets_df[['First name from Team Member', 'Last name from Team Member']] = \
            tickets_df['Team Member'].str.split(' ', n=1, expand=True)
        
        # Convert names to lowercase for accurate comparison
        tickets_df['First name from Team Member'] = tickets_df['First name from Team Member'].str.lower()
        tickets_df['First Name'] = tickets_df['First Name'].str.lower()
        
        # Identify mismatches between the first name extracted from "Team Member" and the "First Name" column
        mismatches = tickets_df['First name from Team Member'] != tickets_df['First Name']
        
        # If mapping data is available, update the 'Email' using the mapping CSV (match on "Team member")
        if df_mapping is not None:
            for i in tickets_df[mismatches].index:
                team_member = tickets_df.loc[i, 'Team Member']
                mapping_entry = df_mapping[df_mapping['Team member'] == team_member].reset_index(drop=True)
                if not mapping_entry.empty:
                    correct_email = mapping_entry.loc[0, 'Email']
                    tickets_df.loc[i, 'Email'] = correct_email
    
    # Drop unnecessary columns for final output.
    drop_columns = [
        "First Name", "Last Name", "Country", "Status Friendly", "Fund ID", "Fund Code", "Fund Name", 
        "Dedication Type", "Dedication Name", "Company", "Dedication Recipient Name", "Dedication Recipient Email", 
        "Method", "CC Last Four", "CC Expiration Date", "Discount Code", "Method Subtype", "Amount", "Fee", "Fee Covered", 
        "Donated", "Payout", "Currency", "Plan ID", "Frequency", "Check Number", "Check Deposited (UTC)", 
        "Payment Captured (UTC)", "Refund Date (UTC)", "Dispute Status", "Acknowledged", "Hide Name", "Hide Amount", 
        "Public Name", "Public Message", "Donor's Local Timezone", "Payment Captured (Donor's Local Timezone)", 
        "UTM Source", "UTM Medium", "UTM Campaign", "UTM Term", "UTM Content", "Referrer", "Widget Id", "Match Name", 
        "Match Amount", "External ID", "Household ID", "Household Name", "Item Subtype", "Item Quantity", "Item Price", 
        "Item Discount", "Item Total", "First name from Team Member", "Last name from Team Member"
    ]
    tickets_df.drop(columns=drop_columns, errors='ignore', inplace=True)
    
    # Mapping of campaign descriptions (from 'Item Description') to sheet names
    description_to_sheet_map = {
        '2025 Ride for Missing Children - MV New / Returning Riders': 'NewAndReturning',
        '2025 Ride for Missing Children - MV Reciprocal Riders': 'Reciprocal',
        '2025 Ride for Missing Children - MV Volunteer': 'Volunteer'
    }
    
    # Ensure output directory exists
    output_dir = 'Rider_Volunteer_CSVs'
    os.makedirs(output_dir, exist_ok=True)
    
    # Save individual CSV files for each campaign description
    for original_desc, sheet_name in description_to_sheet_map.items():
        subset_df = tickets_df[tickets_df['Item Description'] == original_desc]
        save_and_compare_df(subset_df, sheet_name, output_dir)
    
    # Create a master Excel file with separate sheets per campaign
    current_time = int(time.time())
    mapped_excel_path = f'Rider_Volunteer_MasterList_{current_time}.xlsx'
    
    with pd.ExcelWriter(mapped_excel_path, engine='xlsxwriter') as writer:
        for original_desc, sheet_name in description_to_sheet_map.items():
            subset_df = tickets_df[tickets_df['Item Description'] == original_desc]
            subset_df.to_excel(writer, sheet_name=sheet_name, index=False)
            auto_adjust_columns_width(subset_df, writer, sheet_name)
    print(f"Master Excel file saved as {mapped_excel_path}")

if __name__ == '__main__':
    main()

