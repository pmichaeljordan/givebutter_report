import pandas as pd
import time

#!/usr/bin/env python3

# Load the CSV file into a DataFrame
file_path = 'transactions.csv'
df = pd.read_csv(file_path)

# Change any column data that matches '2024 Ride for Missing Children - MV New and Returning Riders' to '2024 Ride for Missing Children - MV New / Returning Riders'
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
