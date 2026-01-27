# Givebutter Campaign Report Scripts

This repository contains scripts to process Givebutter campaign data for the Ride for Missing Children event. The scripts handle ticket parsing, transaction processing, contact management, and fundraising reporting.

## Overview

These scripts process exported data from Givebutter to generate:
- Excel workbooks with riders and volunteers separated by ticket type
- Google Contacts imports with appropriate labels
- Fundraising progress reports uploaded to Google Drive

## Prerequisites

- Python 3.x
- Required packages: `pandas`, `openpyxl`, `xlsxwriter`, `google-api-python-client`, `google-auth-oauthlib`
- Google OAuth credentials files (for Drive and Contacts API access)

## Workflow

### Step 1: Export Data from Givebutter

Before running any scripts, export the following files from your Givebutter campaign:

1. **Tickets Export**
   - Navigate to your campaign in Givebutter
   - Export tickets data
   - File will be saved as: `tickets-YYYY-MM-DD-{id}.csv`
   - Place in this directory

2. **Transactions Export**
   - Navigate to your campaign in Givebutter
   - Export transactions data
   - Rename file to: `transactions.csv`
   - Place in this directory

### Step 2: Run the Scripts

Run the scripts in the following order:

#### 1. Process Tickets & Sync Contacts (`ticket_processor.py`)

```bash
python3 ticket_processor.py
```

**What it does:**
- Automatically runs `ticket_parse.py` internally to:
  - Find the most recent `tickets-*.csv` file
  - Warn if the file is older than 7 days
  - Filter out revoked tickets
  - Create separate Excel sheets for each ticket type
  - Generate a T-shirt sizes summary
- Uploads the parsed tickets Excel file to Google Drive
- Processes MV (rider/volunteer) sheets to create a Google Contacts CSV
- Imports contacts to Google Contacts with year-based labels:
  - All volunteers → `2026_Volunteer`
  - All riders → `2026_Rider`
  - New riders specifically → `2026_Rider` AND `2026_New_Riders` (for targeted emails)

**Requirements:**
- Exported `tickets-*.csv` file from Givebutter
- Google OAuth credentials files:
  - `credentials_account1.json` (for Drive)
  - `credentials_account2.json` (for Contacts - Account 1)
  - `credentials_account3.json` (for Contacts - Account 2)

**Output:**
- `parsed_tickets_MM-DD-YYYY.xlsx` - Workbook with sheets for each ticket type
- Excel file uploaded to Google Drive
- `output.csv` - Google Contacts import file
- Contacts synced to two Google accounts with appropriate labels

---

#### 2. Process Transactions (`transaction_parse.py`)

```bash
python3 transaction_parse.py
```

**What it does:**
- Reads `transactions.csv` file
- Warns if the file is older than 7 days
- Filters for ticket transactions only
- Corrects team member email addresses using `name_mapping.csv` (if available)
- Separates riders and volunteers into different sheets

**Output:**
- `Rider_Volunteer_MasterList_{timestamp}.xlsx` - Master workbook with sheets:
  - **NewAndReturning** - New and returning riders
  - **Reciprocal** - Reciprocal riders
  - **Volunteer** - Volunteers
- Individual CSV files in `Rider_Volunteer_CSVs/` directory

---

#### 3. Fundraising Report (`fundraising_parse.py`)

```bash
python3 fundraising_parse.py
```

**What it does:**
- Connects to Givebutter API to fetch live campaign data
- Prompts you to select the active campaign
- Retrieves campaign members and their fundraising progress
- Merges ticket type information with fundraising data
- Creates a fundraising progress report
- Uploads the report to Google Drive (shared with committee chair)

**Requirements:**
- Environment variable: `GIVEBUTTER_API_TOKEN`
- Google OAuth credentials: `credentials_account3.json`

**Output:**
- `GiveButterReport.xlsx` - All tickets by type
- `Fundraising_Progress_{timestamp}.xlsx` - Campaign member fundraising data
- Both files uploaded to Google Drive folder

---

## Important Files

### Input Files
- `tickets-*.csv` - Exported from Givebutter
- `transactions.csv` - Exported from Givebutter
- `name_mapping.csv` - (Optional) Email corrections for team members

### Configuration Files
- `credentials_account1.json` - Google Drive OAuth
- `credentials_account2.json` - Google Contacts OAuth (Account 1)
- `credentials_account3.json` - Google Contacts OAuth (Account 2)
- `data_map.txt` - (Optional) Ticket number to email mapping

### Output Files
- `parsed_tickets_*.xlsx` - Parsed tickets by type
- `Rider_Volunteer_MasterList_*.xlsx` - Riders and volunteers by category
- `GiveButterReport.xlsx` - All tickets report
- `Fundraising_Progress_*.xlsx` - Fundraising progress report
- `output.csv` - Google Contacts import file

## Notes

- All scripts are **year-agnostic** and will automatically work for future years
- Scripts will warn you if input files are older than 7 days
- The scripts filter data for the current year automatically
- Contact labels are generated dynamically (e.g., `2026_Rider`, `2026_Volunteer`, `2026_New_Riders`)
- **New Rider Detection**: The script automatically detects new riders based on the "Are you a new or returning rider?" field in the tickets export. New riders will be added to both the general `{YEAR}_Rider` group and the specific `{YEAR}_New_Riders` group, allowing you to send targeted emails to just new riders.

## Google Contact Groups

The `ticket_processor.py` script creates and manages the following contact groups automatically:

- **`{YEAR}_Rider`** - All riders (new, returning, reciprocal, corporate, etc.)
- **`{YEAR}_New_Riders`** - Only new riders (subset of `{YEAR}_Rider`)
- **`{YEAR}_Volunteer`** - All volunteers

**Note:** These groups are recreated fresh each time you run the script. Old year groups (e.g., `2025_Rider`) are preserved and not deleted.

### Emailing Specific Groups

- To email all riders: Use the `{YEAR}_Rider` group
- To email only new riders: Use the `{YEAR}_New_Riders` group
- To email volunteers: Use the `{YEAR}_Volunteer` group

## Troubleshooting

**"No CSV files found"**
- Make sure you've exported the files from Givebutter
- Check that filenames match expected patterns

**"File is X days old"**
- Export fresh data from Givebutter for the most accurate results
- You can proceed with old files if needed

**Google API errors**
- Ensure credentials files are in the working directory
- Check that OAuth tokens are valid (delete `*_token.json` files to re-authenticate)

## Script Dependencies

```
ticket_processor.py (RUN THIS FIRST)
    ├── ticket_parse.py (called automatically)
    └── google_upload.py

transaction_parse.py (standalone)
fundraising_parse.py (standalone)
```

**Note:** You do NOT need to run `ticket_parse.py` separately - it is automatically called by `ticket_processor.py`.
