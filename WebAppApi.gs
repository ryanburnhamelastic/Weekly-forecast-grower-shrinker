/**
 * WebAppApi.gs
 * Web App API for Weekly Forecast CA Notes Interface
 * Provides endpoints for web app to interact with spreadsheet data
 */

/**
 * Serves the web app HTML
 * @param {Object} e - Event object
 * @returns {HtmlOutput} Web app HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('Weekly Forecast - CA Notes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets current user information
 * @returns {Object} User info with { email, displayName, caName, hasAccounts }
 */
function getCurrentUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const displayName = userEmail.split('@')[0].replace('.', ' ');

    Logger.log(`getCurrentUserInfo: Getting info for ${userEmail}`);

    // Read CA-Lookup to find user's CA name
    const caMapping = readCALookup();
    let caName = null;
    let hasAccounts = false;

    Logger.log(`getCurrentUserInfo: Found ${Object.keys(caMapping).length} entries in CA-Lookup`);

    // Search for user email in CA-Lookup
    for (const [accountId, caInfo] of Object.entries(caMapping)) {
      if (caInfo.email && caInfo.email.toLowerCase() === userEmail.toLowerCase()) {
        caName = caInfo.name;
        hasAccounts = true;
        Logger.log(`getCurrentUserInfo: Found CA Name: ${caName}`);
        break;
      }
    }

    if (!hasAccounts) {
      Logger.log(`getCurrentUserInfo: User not found in CA-Lookup (will show empty state)`);
    }

    return {
      email: userEmail,
      displayName: displayName,
      caName: caName || displayName,
      hasAccounts: hasAccounts
    };
  } catch (error) {
    Logger.log(`Error in getCurrentUserInfo: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    return {
      email: 'unknown@elastic.co',
      displayName: 'Unknown User',
      caName: 'Unknown User',
      hasAccounts: false,
      error: error.message
    };
  }
}

/**
 * Gets list of all weekly tabs (YYYY-MM-DD format)
 * @returns {Array} Array of tab objects with { tabName, date, version, isLatest }
 */
function getWeeklyTabList() {
  try {
    const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
    const allSheets = targetSheet.getSheets();
    const weeklyTabs = [];

    // Date pattern: YYYY-MM-DD or YYYY-MM-DD_vN
    const datePattern = /^(\d{4}-\d{2}-\d{2})(_v(\d+))?$/;

    for (const sheet of allSheets) {
      const name = sheet.getName();
      const match = name.match(datePattern);

      if (match) {
        const dateStr = match[1];
        const version = match[3] ? parseInt(match[3]) : 1;

        weeklyTabs.push({
          tabName: name,
          date: dateStr,
          version: version,
          dateObj: new Date(dateStr)
        });
      }
    }

    // Sort by date descending (most recent first)
    weeklyTabs.sort((a, b) => b.dateObj - a.dateObj);

    // Mark the latest tab
    if (weeklyTabs.length > 0) {
      weeklyTabs[0].isLatest = true;
    }

    return weeklyTabs.map(tab => ({
      tabName: tab.tabName,
      date: tab.date,
      version: tab.version,
      isLatest: tab.isLatest || false
    }));

  } catch (error) {
    Logger.log(`Error in getWeeklyTabList: ${error.message}`);
    return [];
  }
}

/**
 * Gets ALL accounts for a specific weekly tab (not filtered by user)
 * @param {string} weeklyTabName - Name of the weekly tab
 * @returns {Object} Object with { growers, shrinkers, tabName, lastUpdated }
 */
function getAllAccounts(weeklyTabName) {
  try {
    const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
    const tab = targetSheet.getSheetByName(weeklyTabName);

    if (!tab) {
      throw new Error(`Tab "${weeklyTabName}" not found`);
    }

    // Read ALL accounts from tab (no filtering by CA)
    // Growers: rows 3-13 (header row 2, data rows 3-13)
    // Shrinkers: rows 17-27 (header row 16, data rows 17-27)
    const growers = extractAllAccountsFromSection(tab, 3, 13, 'Growers');
    const shrinkers = extractAllAccountsFromSection(tab, 17, 27, 'Shrinkers');

    return {
      growers: growers,
      shrinkers: shrinkers,
      tabName: weeklyTabName,
      lastUpdated: new Date().toISOString()
    };

  } catch (error) {
    Logger.log(`Error in getAllAccounts: ${error.message}`);
    return {
      growers: [],
      shrinkers: [],
      tabName: weeklyTabName,
      lastUpdated: new Date().toISOString(),
      error: error.message
    };
  }
}

/**
 * Helper function to extract ALL accounts from a section (no CA filtering)
 * @param {Sheet} sheet - The sheet to read from
 * @param {number} startRow - Start row (1-indexed)
 * @param {number} endRow - End row (1-indexed)
 * @param {string} section - Section name ('Growers' or 'Shrinkers')
 * @returns {Array} Array of account objects
 */
function extractAllAccountsFromSection(sheet, startRow, endRow, section) {
  const accounts = [];

  // Column structure based on actual sheet (1-indexed):
  // A: Link
  // B: Account Name (Biggest 10 Growers/Shrinkers MoM)
  // C: Area
  // D: Jan Committed ARR
  // E: Jan ARR Cons Rate
  // F: Jan Consumption ARR
  // G: WoW ECU Change
  // H: WoW % ECU Change
  // I: Feb vs Jan Tracked Change
  // J: Tracking MoM ARR Growth
  // K: Feb Tracking ARR
  // L: CA Forecast
  // M: CA Forecast Variance
  // N: CA Forecast Var %
  // O: Customer Architect
  // P: Notes

  for (let row = startRow; row <= endRow; row++) {
    const rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];  // Read columns A-P

    // Array is 0-indexed, so column N is at index N-1
    const link = rowData[0] ? rowData[0].toString().trim() : '';            // Column A
    const accountName = rowData[1] ? rowData[1].toString().trim() : '';     // Column B
    const area = rowData[2] ? rowData[2].toString().trim() : '';            // Column C
    const wowPercent = rowData[7];                                          // Column H (WoW % ECU Change)
    const momPercent = rowData[9];                                          // Column J (Tracking MoM ARR Growth)
    const caForecast = rowData[11];                                         // Column L
    const caForecastVar = rowData[12];                                      // Column M
    const caForecastVarPct = rowData[13];                                   // Column N
    const customerArchitect = rowData[14] ? rowData[14].toString().trim() : ''; // Column O
    const notes = rowData[15] ? rowData[15].toString().trim() : '';         // Column P

    // Include all accounts that have a name (no CA filtering)
    if (accountName) {
      accounts.push({
        link: link,
        accountName: accountName,
        area: area,
        wowPercent: wowPercent,
        momPercent: momPercent,
        caForecast: caForecast,
        caForecastVar: caForecastVar,
        caForecastVarPct: caForecastVarPct,
        customerArchitect: customerArchitect,  // Include CA name for filtering
        currentNote: notes,
        rowNumber: row,
        section: section
      });
    }
  }

  return accounts;
}

/**
 * TEST FUNCTION - Check CA-Lookup authorization
 * Run this from the Apps Script editor to see if your email is in CA-Lookup
 */
function testCAAuthorization() {
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(`Testing authorization for: ${userEmail}`);

  const caMapping = readCALookup();
  Logger.log(`Total entries in CA-Lookup: ${Object.keys(caMapping).length}`);

  let found = false;
  for (const [accountId, caInfo] of Object.entries(caMapping)) {
    if (caInfo.email && caInfo.email.toLowerCase() === userEmail.toLowerCase()) {
      Logger.log(`✓ FOUND! Account ID: ${accountId}, CA Name: ${caInfo.name}, Email: ${caInfo.email}`);
      found = true;
      break;
    }
  }

  if (!found) {
    Logger.log(`✗ NOT FOUND. Your email is not in CA-Lookup.`);
    Logger.log(`Sample CA-Lookup entries (first 10):`);
    Object.entries(caMapping).slice(0, 10).forEach(([accountId, caInfo]) => {
      Logger.log(`  - Account ${accountId}: ${caInfo.name} (${caInfo.email})`);
    });
  }

  const userInfo = getCurrentUserInfo();
  Logger.log(`\ngetCurrentUserInfo() result:`);
  Logger.log(JSON.stringify(userInfo, null, 2));
}

/**
 * Saves a note for a specific account row
 * @param {string} weeklyTabName - Name of the weekly tab
 * @param {number} rowNumber - Row number to update (1-indexed)
 * @param {string} noteText - Note text to save
 * @returns {Object} Result with { success, error, timestamp }
 */
function saveAccountNote(weeklyTabName, rowNumber, noteText) {
  try {
    // Validate inputs
    if (!weeklyTabName || !rowNumber || noteText === undefined) {
      throw new Error('Missing required parameters');
    }

    // Limit note length
    const MAX_NOTE_LENGTH = 2000;
    let sanitizedNote = noteText.toString().trim();
    if (sanitizedNote.length > MAX_NOTE_LENGTH) {
      sanitizedNote = sanitizedNote.substring(0, MAX_NOTE_LENGTH);
    }

    // Validate row number is in expected range (3-13 or 17-27)
    if (!((rowNumber >= 3 && rowNumber <= 13) || (rowNumber >= 17 && rowNumber <= 27))) {
      throw new Error('Invalid row number');
    }

    // Get user info for authorization
    const userInfo = getCurrentUserInfo();
    if (!userInfo.isAuthorized || !userInfo.caName) {
      throw new Error('User not authorized');
    }

    // Open sheet and verify user owns this row
    const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
    const tab = targetSheet.getSheetByName(weeklyTabName);

    if (!tab) {
      throw new Error(`Tab "${weeklyTabName}" not found`);
    }

    // Check Customer Architect column (column O = 15) matches user's CA name
    const customerArchitect = tab.getRange(rowNumber, 15).getValue()?.toString().trim();

    if (customerArchitect !== userInfo.caName) {
      throw new Error('Not authorized to edit this row');
    }

    // Update Notes column (column P = 16)
    tab.getRange(rowNumber, 16).setValue(sanitizedNote);

    return {
      success: true,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    Logger.log(`Error in saveAccountNote: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}
