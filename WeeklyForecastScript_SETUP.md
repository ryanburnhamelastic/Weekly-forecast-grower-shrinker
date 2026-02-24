# Weekly Forecast Script - Setup Instructions

## Overview
This script automates the weekly process of:
1. Reading Growers and Shrinkers data from the Global sheet
2. Looking up Customer Architects from CA-Lookup
3. Creating a dated tab in the target sheet
4. Writing combined data with CA assignments
5. Adding @mention comments to notify CAs

---

## Setup Instructions

### Step 1: Create CA-Lookup Tab

1. Open your target sheet: `1rsQUBs1nD8Ywve88910jGcObJRiZnxSL0gHEK9RMrBA`
2. Create a new tab named **CA-Lookup** (exact name, case-sensitive)
3. Set up the structure:
   - **Column A**: Customer Name (must match names from Global sheet, without asterisk)
   - **Column B**: CA Name (e.g., "John Smith", "Jane Doe")

Example:
```
| Customer Name              | CA Name        |
|----------------------------|----------------|
| United HealthCare Services | John Smith     |
| Planet Labs PBC            | Jane Doe       |
| Barclays Bank PLC          | Bob Johnson    |
| THD                        | Alice Williams |
```

**Important**: Customer names should match exactly as they appear in the Global sheet (without the asterisk prefix).

---

### Step 2: Install the Script

1. Open the **target sheet** (not the Global sheet)
2. Go to **Extensions** > **Apps Script**
3. Delete any existing code in the editor
4. Copy the entire contents of `WeeklyForecastScript.js`
5. Paste into the Apps Script editor
6. Click **Save** (disk icon) and give it a name like "Weekly Forecast Script"

---

### Step 3: Grant Permissions

1. Click **Run** > **Run function** > **weeklyForecastUpdate**
2. You'll see "Authorization required" dialog
3. Click **Review permissions**
4. Choose your Google account
5. Click **Advanced** > **Go to [project name] (unsafe)**
6. Click **Allow**

The script needs permission to:
- Read from the Global sheet
- Write to the target sheet
- Create new tabs
- Add comments

---

### Step 4: Test the Script

#### Test CA Lookup First
1. Go to **Extensions** > **Apps Script**
2. Click **Run** > **Run function** > **testCALookup**
3. Check the alert dialog for number of mappings found
4. Go to **View** > **Logs** (or press Ctrl+Enter) to see details
5. Verify all customer names are mapped correctly

#### Test Full Script
1. Click **Run** > **Run function** > **weeklyForecastUpdate**
2. Wait for completion (should take 10-30 seconds)
3. Check for success alert
4. Review the new tab created (should be named with today's date like "2025-10-20")
5. Verify:
   - Both Growers and Shrinkers tables are present
   - Customer Architect column is populated
   - Notes column exists
   - Comments are added to Notes cells (hover over cells to see)

---

### Step 5: Set Up Weekly Trigger

1. In Apps Script editor, click **Triggers** (clock icon on left sidebar)
2. Click **+ Add Trigger** (bottom right)
3. Configure:
   - **Choose which function to run**: `weeklyForecastUpdate`
   - **Choose which deployment should run**: Head
   - **Select event source**: Time-driven
   - **Select type of time based trigger**: Week timer
   - **Select day of week**: Choose your preferred day (e.g., Monday)
   - **Select time of day**: Choose your preferred time (e.g., 8am-9am)
4. Click **Save**

The script will now run automatically every week at the scheduled time.

---

## Usage

### Manual Execution (Using Custom Menu)

After the script is installed, you'll see a new menu in your Google Sheet:

1. Open the target sheet
2. Click **Weekly Forecast** menu (next to Help)
3. Choose **Run Weekly Update**
4. Wait for completion alert

### What Happens Each Week

1. Script runs at scheduled time
2. Creates new tab with date name (e.g., "2025-10-20")
3. If tab exists, appends version number (e.g., "2025-10-20_v2")
4. Copies all data from both tables
5. Adds Customer Architect names based on CA-Lookup
6. Adds comments with @mentions to notify CAs
7. CAs receive email notifications (if they have comment notifications enabled)

---

## Troubleshooting

### Common Issues

**"CA-Lookup tab not found"**
- Ensure tab name is exactly "CA-Lookup" (case-sensitive)
- Tab must be in the target sheet, not Global sheet

**CA Names Not Matching**
- Check that customer names in CA-Lookup match exactly (without asterisk)
- Remove any extra spaces or special characters
- Test with `testCALookup` function to verify mappings

**Email Notifications Not Working**
- @mentions in comments only work if:
  - The CA has access to the sheet
  - The CA has comment notifications enabled in Google Sheets
  - Email format is correct (firstname.lastname@elastic.co)

**Script Times Out**
- If you have many rows, the script might timeout
- Try running during off-peak hours
- Contact support if issue persists

**No Data Copied**
- Verify Global sheet ID is correct: `1YkOxHI5EZjVoCKJuifn2tZGZVj7gp_spxEv0YIAFVKw`
- Verify tab name is exactly "Global"
- Check that rows 25-36 and 39-50 contain the expected data

---

## Viewing Logs

To debug issues or see what the script did:

1. Go to **Extensions** > **Apps Script**
2. Click **View** > **Logs** (or press Ctrl+Enter)
3. Or click **Executions** (left sidebar) to see history

Logs show:
- Number of CA mappings found
- Number of growers/shrinkers read
- Tab name created
- Number of rows written
- Comments added

---

## Customization Options

### Change Date Format

In the script, find this line (around line 172):
```javascript
const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
```

Change format:
- `'yyyy-MM-dd'` → "2025-10-20"
- `'MMM dd yyyy'` → "Oct 20 2025"
- `'MM-dd-yyyy'` → "10-20-2025"

### Change Email Domain

If your organization uses a different domain, update this line (around line 13):
```javascript
EMAIL_DOMAIN: '@elastic.co',
```

### Change Email Format

To use different email format (e.g., first initial + last name), modify the `convertNameToEmail` function (around line 328).

---

## Support

For issues:
1. Check the troubleshooting section above
2. Review execution logs for error messages
3. Test individual components (use testCALookup)
4. Verify all configuration values are correct

---

## Important Notes

- **Backup**: Script creates new tabs, doesn't modify existing data
- **Permissions**: Each user running the script needs appropriate sheet access
- **Notifications**: CAs must have sheet access to receive @mention notifications
- **Version Control**: Tabs append version numbers if run multiple times same day
- **Time Zone**: Date uses script's timezone setting (can be changed in Apps Script settings)
