# Givebutter Reporting System

This project provides automated reporting and data management tools for Givebutter fundraising campaigns, specifically for the "Ride for Missing Children" events.

## Project Overview

This system fetches data from the Givebutter API, processes it, and generates Excel reports that are automatically uploaded to Google Drive. It also manages Google Contacts for riders and volunteers.

## Core Components

### Main Scripts

- **fundraising_parse.py** - Primary script for generating fundraising reports
  - Fetches campaign data and member information from Givebutter API
  - Generates two types of reports:
    1. Ticket reports (GiveButterReport.xlsx) - organized by ticket type
    2. Fundraising progress reports - includes member fundraising stats with ticket titles
  - Automatically uploads reports to Google Drive
  - Filters tickets by current year in title field

- **google_upload.py** - Google integration utilities
  - Uploads files to Google Drive
  - Manages Google Contacts import/export
  - Creates and manages contact groups: "2025_Rider" and "2025_Volunteer"
  - Processes Excel sheets starting with "MV" prefix
  - Handles email corrections via data_map.txt mapping file

- **transaction_parse.py** - Transaction data processing
  - Processes Givebutter transaction CSV exports
  - Filters for ticket transactions
  - Splits data by campaign
  - Creates Excel files with separate sheets per campaign

- **ticket_parse.py** - Ticket data parsing utilities

## Environment & Authentication

### Required Environment Variables

```bash
GIVEBUTTER_API_TOKEN  # Bearer token for Givebutter API authentication
```

### Google OAuth Credentials

The project uses multiple Google service accounts with different credential files:
- `credentials_account1.json` - Shepherd account (Drive)
- `credentials_account2.json` - Shepherd account (Contacts)
- `credentials_account3.json` - Admin account (Contacts & Drive)

Token files are auto-generated on first run:
- `drive_token.json` - Main Drive access
- `shepherd_drive_token.json` - Shepherd Drive access
- `contacts_token_account1.json` - Shepherd Contacts
- `contacts_token_account2.json` - Admin Contacts

### Google Drive Configuration

- Target folder ID: `1aFkyJD_Qtto8ZOVZ8Q6FZLS5rBF2ERcw`

## Data Processing

### Ticket Title Mapping

The system uses a hardcoded mapping to shorten long ticket titles:

```python
"2025 Ride for Missing Children - MV New / Returning High School Student Riders" → "High School Riders"
"2025 Ride for Missing Children - MV New / Returning Riders" → "New/Returning"
"2025 Ride for Missing Children - MV Volunteer" → "Volunteer"
"2025 Ride for Missing Children - MV Corporate Riders" → "Corporate Riders"
"2025 Ride for Missing Children - MV Reciprocal Riders" → "Reciprocal Riders"
```

### Email Standardization

- All email addresses are converted to lowercase for consistent matching
- Email corrections can be provided via `data_map.txt` file with format: `Ticket Number, Incorrect Email, Correct Email`

### Year Filtering

- Tickets are filtered to only include those with titles starting with the current year (e.g., "2025")
- This filter is applied in both ticket reports and fundraising reports

## Generated Reports

### GiveButterReport.xlsx
- Contains tickets filtered by current year
- One sheet per ticket type
- Columns: Name, First, Email, Phone, Title, Price, Signup Date
- Sheet names derived from ticket title (last part after final dash)

### Fundraising_Progress_YYYYMMDD_HHMMSS.xlsx
- Single "Fundraising" sheet
- Combines campaign member data with ticket information
- Columns: First Name, Last Name, Display Name, Email, Phone, Raised, Goal, Donors, Title
- Includes total raised summary at bottom
- Automatically uploaded to Google Drive

### output.csv (Google Contacts format)
- Generated from Excel sheets starting with "MV"
- Contains all columns required for Google Contacts import
- Tagged as "2025_Rider" or "2025_Volunteer"

## API Integration

### Givebutter API Endpoints

```
Base URL: https://api.givebutter.com/v1/

GET /campaigns - List all campaigns
GET /campaigns/{id}/members?page={n} - Get campaign members (paginated)
GET /tickets?page={n} - Get all tickets (paginated)
```

### API Headers
```python
{
    "accept": "application/json",
    "Authorization": f"Bearer {GIVEBUTTER_API_TOKEN}"
}
```

## Important Notes

### Security
- Never commit credential files (*.json) or token files
- Keep GIVEBUTTER_API_TOKEN secure
- Credential files contain OAuth client secrets

### Data Privacy
- Email addresses are processed and stored in reports
- Phone numbers are included in reports
- Reports are uploaded to shared Google Drive folder

### Rate Limiting
- Google Contacts API has rate limits - script includes 1-second delays (`API_SLEEP`)
- Membership modifications use exponential backoff on 429 errors
- Max retries: 5 attempts

### Contact Group Management
- Target groups ("2025_Rider", "2025_Volunteer") are deleted and recreated on each import
- This ensures clean group membership based solely on current CSV data
- Non-target groups are preserved if they already exist

## Common Tasks

### Running the Main Report Generation
```bash
./fundraising_parse.py
```
This will:
1. Prompt for campaign selection
2. Generate both ticket and fundraising reports
3. Upload fundraising report to Google Drive

### Processing Excel Sheets for Google Contacts
```python
from google_upload import process_mv_sheets, import_to_google_contacts

csv_file = process_mv_sheets('input_file.xlsx')
import_to_google_contacts(csv_file)
```

### Manual Google Drive Upload
```python
from google_upload import upload_to_drive

upload_to_drive('file.xlsx', 'folder_id')
```

## Dependencies

- pandas - Data manipulation
- openpyxl - Excel file generation
- requests - HTTP API calls
- google-auth-oauthlib - Google OAuth flow
- google-api-python-client - Google API clients
- google-auth - Google authentication

## Workflow

1. Script authenticates with Givebutter API using bearer token
2. User selects a campaign from the list
3. System fetches all campaign members and tickets (handles pagination)
4. Data is processed, filtered by year, and formatted
5. Excel reports are generated with proper column widths
6. Fundraising report is uploaded to designated Google Drive folder
7. Contact data can be imported to Google Contacts with proper group assignments

## Code Style & Conventions

- Scripts use `#!/usr/bin/env python3` shebang
- Email addresses always converted to `.lower()` for consistency
- DataFrame column names standardized to lowercase before processing
- Excel column widths auto-adjusted based on content
- Timestamps in filename format: `YYYYMMDD_HHMMSS`
- Sheet names truncated to 31 characters (Excel limit)
- Forward slashes replaced with underscores in sheet names
