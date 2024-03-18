import pandas as pd
import os
import glob
import time

#!/usr/bin/env python3

''' Export transactions with detail from Givegbutter. Include all fields. Parser will drop any uncessary fields.
    Create a mapper file for any tickets that were paid for someone other than the rider as csv with the following columns:
    1. Name Team member (must match)
    2. email address from Team member (Will replace the incorrect email address)
    '''

# Load the CSV file into a DataFrame
file_path = 'transactions.csv'
df = pd.read_csv(file_path)

# Change any column data that matches '2024 Ride for Missing Children - MV New and Returning Riders' to '2024 Ride for Missing Children - MV New / Returning Riders'
# This was due to testing tickets before going live. Don't do this next year. Also use shorter ticket names.
df['Description'] = df['Description'].replace('2024 Ride for Missing Children - MV New and Returning Riders', '2024 Ride for Missing Children - MV New / Returning Riders')

# Filter out the rows where Subtype is 'ticket'
tickets_df = df[df['Subtype'] == 'ticket']

# Columns to be dropped
columns_to_drop = ["Campaign", "Campaign slug", "Team", "Reference #"]

# Drop the specified columns
tickets_df = tickets_df.drop(columns=columns_to_drop, errors='ignore')

# Load the name_mapping CSV file
df_mapping = pd.read_csv('name_mapping.csv')

# Split the 'Team member' column into first and last names
tickets_df[['First name from Team member', 'Last name']] = tickets_df['Team member'].str.split(' ', n=1, expand=True)

# Convert to lowercase
tickets_df['First name from Team member'] = tickets_df['First name from Team member'].str.lower()
tickets_df['First name'] = tickets_df['First name'].str.lower()

# Find mismatches between 'First name from Team member' and 'First name'
mismatches = tickets_df['First name from Team member'] != tickets_df['First name']

# Strip leading/trailing spaces from column names
df_mapping.columns = df_mapping.columns.str.strip()

# For each mismatch, find the correct email in the mapping DataFrame
for i in tickets_df[mismatches].index:
    team_member = tickets_df.loc[i, 'Team member']
    mapping_entry = df_mapping[df_mapping['Team member'] == team_member].reset_index(drop=True)

    # If a match is found in the mapping DataFrame, update the 'Email' column in tickets_df
    if not mapping_entry.empty:
        correct_email = mapping_entry.loc[0, 'Email']
        tickets_df.loc[i, 'Email'] = correct_email

# Columns to be dropped
columns_to_drop = ["First name", "Last name", "Country", "Status", "Fund ID", "Fund code", "Fund name", "Dedication type", "Dedication name", "Company", "Dedication recipient name", "Dedication recipient email", 
                   "Method", "Credit card last four", "Credit card expiration date", "Discount code", "Method subtype", "Amount", "Fee", "Fee covered", "Donated", "Payout", "Currency",
                    "Recurring plan ID", "Frequency", "Check number", "Check deposited (UTC)", "Payment captured (UTC)", "Refund date (UTC)", "Dispute status", "Acknowledged", "Anonymous hide name",
                    "Anonymous hide amount", "Public name", "Public message", "Donor local timezone", "UTM source", "UTM medium", "UTM campaign", "UTM term", "UTM content", "Referrer", "Widget ID", "Match name", "Match amount", "External ID",
                     "Household ID", "Household name", "Subtype", "Quantity", "Price", "Discount", "Total", "First name from Team member" ]

# Drop the specified columns
tickets_df = tickets_df.drop(columns=columns_to_drop, errors='ignore')

# Mapping of original descriptions to new sheet names provided by the user
description_to_sheet_map = {
    '2024 Ride for Missing Children - MV New / Returning Riders': 'NewAndReturning',
    '2024 Ride for Missing Children - MV Reciprocal Riders': 'Reciprocal',
    '2024 Ride for Missing Children - MV Volunteer': 'Volunteer'
}

# Function to auto-adjust columns' width
def auto_adjust_columns_width(df, writer, sheet_name):
    for column in df:
        # Find the length of the longest entry in the column
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        # Find the column index (0-indexed)
        col_idx = df.columns.get_loc(column)
        # Adjust the column width at the column index
        writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)

# Assuming the directory structure and file naming convention
output_dir = 'Rider_Volunteer_CSVs'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Function to save DataFrame to CSV and compare with the previous version
def save_and_compare_df(df, description):
    # Filename convention: Description_CurrentTime.csv
    filename = f"{description.replace(' ', '_')}_{int(time.time())}.csv"
    filepath = os.path.join(output_dir, filename)
    
    # Save current DataFrame to CSV
    df.to_csv(filepath, index=False)
    
    # Find previous file for the same description
    previous_files = glob.glob(os.path.join(output_dir, f"{description.replace(' ', '_')}*.csv"))
    previous_files = [f for f in previous_files if f != filepath]  # Exclude current file
    
    if previous_files:
        # Assuming there's only one previous file for simplicity
        previous_file = max(previous_files, key=os.path.getctime)  # Get the most recent file
        previous_df = pd.read_csv(previous_file)
        
        # Compare DataFrames to find new rows in the current DataFrame
        # This simplistic comparison assumes you're only looking for new rows added
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

for original_desc, new_sheet_name in description_to_sheet_map.items():
    # Filter the DataFrame based on the Description
    subset_df = tickets_df[tickets_df['Description'] == original_desc]
    
    # Save and compare the subset DataFrame
    save_and_compare_df(subset_df, new_sheet_name)

# Get the current epoch time
current_time = int(time.time())

# Create a new Excel writer object with the current epoch time in the filename
mapped_excel_path = f'Rider_Volunteer_MasterList_{current_time}.xlsx'

with pd.ExcelWriter(mapped_excel_path, engine='xlsxwriter') as mapped_writer:
    # Write each subset of data to a separate sheet based on the new mapping
    for original_desc, new_sheet_name in description_to_sheet_map.items():
        # Find the subset of dataframe for each original description
        subset_df = tickets_df[tickets_df['Description'] == original_desc]
        # Write to the corresponding new sheet name
        subset_df.to_excel(mapped_writer, sheet_name=new_sheet_name, index=False)
        # Auto-adjust columns' width
        auto_adjust_columns_width(subset_df, mapped_writer, new_sheet_name)
