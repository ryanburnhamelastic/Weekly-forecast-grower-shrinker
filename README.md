# Weekly Forecast Grower/Shrinker Script

Automated Google Apps Script for weekly customer growth/shrinkage tracking and Customer Architect notifications.

## Overview

This script automates the weekly process of:
1. Reading Top 10 Growers and Shrinkers from a Global forecast sheet
2. Looking up assigned Customer Architects (CAs) from a lookup table
3. Creating a dated snapshot tab with formatted data
4. Sending email notifications to CAs with their assigned accounts
5. Logging all executions for audit tracking

## Features

- ✅ **Selective Column Copy**: Copies only columns B-M from source data
- ✅ **Format Preservation**: Maintains all colors, number formats, and font styles
- ✅ **Account ID Matching**: Matches accounts by unique account_id (not name)
- ✅ **AMER Region Filter**: Only processes accounts in AMER regions
- ✅ **Smart Email Notifications**: Groups accounts by CA and sends consolidated emails
- ✅ **CC Management**: Auto-CCs Sr Director and relevant Area Managers
- ✅ **Execution Logging**: Tracks every run with timestamps, status, and metrics
- ✅ **Test Mode**: Safe testing without sending emails to actual CAs
- ✅ **CA Notes Interface**: Interactive sidebar for CAs to add and edit notes

## CA Notes Interface

The **CA Notes Interface** provides an easy-to-use sidebar for Customer Architects to view and edit their assigned accounts without navigating through the entire spreadsheet.

### Features

- **User Authorization**: Only authorized CAs (listed in CA-Lookup) can access the interface
- **Filtered View**: Shows only accounts assigned to the current user
- **Week Selection**: Navigate between weekly tabs (defaults to latest)
- **Growers/Shrinkers Sections**: Clear organization with account metrics (WoW %, MoM %, CA Forecast Var %)
- **Inline Editing**: Edit notes with a simple "Edit" button and save directly
- **Real-time Updates**: Changes are saved immediately to the spreadsheet
- **Security**: Server-side validation ensures users can only edit their own accounts

### How to Use

1. Open the target Google Sheet
2. Go to **Weekly Forecast** menu > **CA Notes Interface**
3. The sidebar will open on the right (300px width)
4. Select a week from the dropdown (latest week is pre-selected)
5. View your assigned accounts in Growers and Shrinkers sections
6. Click **Edit** on any account to add or modify notes
7. Click **Save** to save changes or **Cancel** to discard

### Technical Details

The interface uses:
- **Google Apps Script HTML Service** for the sidebar
- **Server-side authorization** via CA-Lookup email matching
- **Row-level security** - users can only edit accounts assigned to them
- **Optimistic UI updates** for instant feedback
- **Toast notifications** for save confirmation

### Development Workflow (Clasp)

This project uses Clasp for local development and version control:

```bash
# Clone the project (first time)
clasp login
clasp clone 1iwilUsB4oZSH0P0UUJsBTxmgNNee9Hv4D2FPtSyzHmHlnaelTQDr-b2C

# Development cycle
# 1. Edit files locally (FrontEnd.gs, Sidebar.html, etc.)
# 2. Push to Apps Script
clasp push

# 3. Test in spreadsheet (use custom menu)
# 4. Pull any changes made in web editor
clasp pull

# Deploy new version
clasp version "v1.0: CA Notes Interface"
clasp deploy --description "Production release"
```

**File Structure:**
- `Code.js` - Main backend script (weekly update logic)
- `FrontEnd.gs` - Server-side API for the sidebar interface
- `Sidebar.html` - Main HTML structure
- `Stylesheet.html` - CSS styles (included in Sidebar.html)
- `JavaScript.html` - Client-side logic (included in Sidebar.html)
- `.clasp.json` - Clasp configuration
- `.claspignore` - Files to exclude from push

## Setup Instructions

### 1. Create CA-Lookup Tab

In your target Google Sheet, create a tab named **CA-Lookup** with this structure:

| A: account_id | B: account_name | C: account_ca_name |
|---------------|-----------------|-------------------|
| 001b000000... | Booking.com B.V. | Stijn Holzhauer |
| 001b000000... | USAA | Ryan Burnham |

**Automated Formula** (recommended):
```
=UNIQUE(FILTER({'Renewals Extract'!A:A, 'Renewals Extract'!B:B, 'Renewals Extract'!O:O},
  'Renewals Extract'!A:A<>""))
```

### 2. Install the Script

1. Open your **target sheet** (not the Global sheet)
2. Go to **Extensions** > **Apps Script**
3. Delete any existing code
4. Copy the entire contents of `WeeklyForecastScript.js`
5. Paste into the Apps Script editor
6. Click **Save** and name it "Weekly Forecast Script"

### 3. Configure Settings

Update these values in the `CONFIG` section (lines 16-34):

```javascript
SR_DIRECTOR_EMAIL: 'your.director@elastic.co',
AREA_MANAGERS: {
  'AMER_STRATEGICS': 'manager1@elastic.co',
  'AMER_MIDMKT_GENBUS': 'manager2@elastic.co',
  'AMER_ENT_WEST': 'manager3@elastic.co',
  'AMER_ENT_EAST_CAN': 'manager4@elastic.co'
}
```

### 4. Grant Permissions

1. Click **Run** > **Run function** > **weeklyForecastUpdate**
2. Click **Review permissions**
3. Choose your Google account
4. Click **Advanced** > **Go to [project name] (unsafe)**
5. Click **Allow**

The script needs permission to:
- Read from the Global sheet
- Write to the target sheet
- Create new tabs
- Send emails

### 5. Test the Script

**IMPORTANT**: The script starts in TEST_MODE by default.

1. Run **Weekly Forecast** > **Run Weekly Update** from the custom menu
2. All emails will go to `ryan.burnham@elastic.co` (or configured TEST_EMAIL)
3. Email subject will show **[TEST]**
4. Email body will show who would have received it
5. Check the **ExecutionLog** tab for run details

### 6. Go Live

When ready for production:

1. Open the script editor
2. Change line 24: `TEST_MODE: false,`
3. Save the script
4. Run manually first to verify
5. Set up weekly trigger (see below)

## Setting Up Weekly Trigger

1. In Apps Script editor, click **Triggers** (clock icon)
2. Click **+ Add Trigger**
3. Configure:
   - **Function**: `weeklyForecastUpdate`
   - **Deployment**: Head
   - **Event source**: Time-driven
   - **Type**: Week timer
   - **Day**: Monday (or your preferred day)
   - **Time**: 8am-9am (or your preferred time)
4. Click **Save**

## Usage

### Manual Execution

Use the **Weekly Forecast** menu in Google Sheets:
- **Run Weekly Update**: Executes the full workflow
- **Test CA Lookup**: Verifies CA mappings are correct

### What Happens Each Week

1. Script runs at scheduled time
2. Creates new tab named `YYYY-MM-DD` (e.g., "2025-10-20")
3. If tab exists, appends version: `YYYY-MM-DD_v2`
4. Copies data from Global sheet (columns B-M only)
5. Preserves all formatting (colors, percentages, etc.)
6. Adds Customer Architect and Notes columns
7. Filters to AMER region accounts only
8. Sends emails to each CA with their account list
9. CCs Sr Director and relevant Area Managers
10. Logs execution to ExecutionLog tab

## Email Format

CAs receive emails like this:

```
Subject: Weekly Forecast - Action Required (6 accounts)

Hi Rajesh,

You have 6 accounts assigned in this week's forecast that need notes added.

Please review and add notes in the sheet:
https://docs.google.com/spreadsheets/d/...

📈 GROWERS (4):
  • United HealthCare Services, Inc. (+1025%)
  • Planet Labs PBC (+53%)
  • The Home Depot (+5%)
  • Dialpad (+31%)

📉 SHRINKERS (2):
  • Social Finance (-28%)
  • USAA (-5%)

Tab: 2025-10-20

Thank you!
```

## ExecutionLog

Every run is automatically logged to the **ExecutionLog** tab with:

| Column | Description |
|--------|-------------|
| Timestamp | When the script ran |
| Tab Name | Name of created tab |
| Status | Success (green) or Failed (red) |
| Total Accounts | Number of AMER accounts processed |
| Emails Sent | Number of CAs notified |
| Duration (sec) | Script execution time |
| Test Mode | Yes/No |
| Error Message | Details if failed |

## Configuration Reference

### Sheet IDs
- **GLOBAL_SHEET_ID**: Source sheet with forecast data
- **TARGET_SHEET_ID**: Destination sheet for weekly tabs

### Table Ranges
Adjust if your data moves:
```javascript
GROWERS_START_ROW: 25,
GROWERS_END_ROW: 36,
SHRINKERS_START_ROW: 39,
SHRINKERS_END_ROW: 50,
```

### Columns to Copy
Currently copies columns B-M (12 columns):
```javascript
START_COLUMN: 2,  // Column B
END_COLUMN: 13,   // Column M
```

### Column Mappings
After reading B-M, columns are 0-indexed:
```javascript
ACCOUNT_ID_COL: 1,      // Column C → index 1
ACCOUNT_NAME_COL: 2,    // Column D → index 2
AREA_COL: 3,            // Column E → index 3
WOW_PERCENT_COL: 8      // Column J → index 8
```

## Troubleshooting

### Common Issues

**"CA-Lookup tab not found"**
- Ensure tab name is exactly "CA-Lookup" (case-sensitive)
- Tab must be in the target sheet

**No CAs matched**
- Verify account_ids in CA-Lookup match Global sheet
- Check that Renewals Extract formula is working
- Run "Test CA Lookup" from menu

**Email notifications not received**
- Verify TEST_MODE is set to false
- Check email addresses are correct
- Ensure CAs have access to the sheet
- Check spam folders

**Script times out**
- Reduce number of rows being processed
- Run during off-peak hours
- Check for slow formulas in source sheets

**Wrong data in ExecutionLog**
- Delete the ExecutionLog tab and run again
- Script will recreate it with correct headers

### Viewing Logs

To debug issues:
1. Go to **Extensions** > **Apps Script**
2. Click **Executions** (left sidebar) to see run history
3. Click any execution to see detailed logs
4. Or press **Ctrl+Enter** for current execution logs

## Maintenance

### Weekly Tasks
- Monitor ExecutionLog for failures
- Verify emails are being received
- Check that CA assignments are current

### Monthly Tasks
- Review CA-Lookup for accuracy
- Update Renewals Extract if schema changes
- Archive old weekly tabs if needed

### As Needed
- Update Area Manager emails when team changes
- Adjust row ranges if data moves
- Add/remove columns from copy range

## File Structure

```
Weekly-forecast-grower-shrinker/
├── WeeklyForecastScript.js       # Main script file
├── WeeklyForecastScript_SETUP.md # Original setup guide
└── README.md                      # This file
```

## Support

For issues or questions:
1. Check this README
2. Review ExecutionLog for error messages
3. Check Apps Script execution logs
4. Verify all configuration values

## Version History

- **v2.0** - Added execution logging, email notifications, test mode
- **v1.5** - Added formatting preservation, account_id matching
- **v1.0** - Initial release with basic data copy

## License

Internal Elastic tool - For authorized use only.
