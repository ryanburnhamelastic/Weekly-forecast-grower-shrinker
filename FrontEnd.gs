/**
 * Weekly Forecast - CA Notes Interface (Frontend)
 *
 * Server-side API functions for the HTML sidebar interface.
 * Provides authentication, data retrieval, and save operations.
 */

// ============================================
// AUTHENTICATION & AUTHORIZATION
// ============================================

/**
 * Gets current user information and authorization status
 * @returns {Object} User info with { email, displayName, caName, isAuthorized }
 */
function getCurrentUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const displayName = userEmail.split('@')[0].replace('.', ' ');

    // Read CA-Lookup to find user's CA name
    const caMapping = readCALookup();
    let caName = null;
    let isAuthorized = false;

    // Search for user email in CA-Lookup
    for (const [accountId, caInfo] of Object.entries(caMapping)) {
      if (caInfo.email && caInfo.email.toLowerCase() === userEmail.toLowerCase()) {
        caName = caInfo.name;
        isAuthorized = true;
        break;
      }
    }

    return {
      email: userEmail,
      displayName: displayName,
      caName: caName || 'Unknown',
      isAuthorized: isAuthorized
    };
  } catch (error) {
    Logger.log(`Error in getCurrentUserInfo: ${error.message}`);
    return {
      email: 'unknown@elastic.co',
      displayName: 'Unknown User',
      caName: null,
      isAuthorized: false,
      error: error.message
    };
  }
}

// ============================================
// DATA RETRIEVAL
// ============================================

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
 * Gets user's assigned accounts for a specific weekly tab
 * @param {string} weeklyTabName - Name of the weekly tab
 * @param {string} userEmail - User's email address
 * @returns {Object} Object with { growers, shrinkers, tabName, lastUpdated }
 */
function getUserAccounts(weeklyTabName, userEmail) {
  try {
    const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
    const tab = targetSheet.getSheetByName(weeklyTabName);

    if (!tab) {
      throw new Error(`Tab "${weeklyTabName}" not found`);
    }

    // Get user's CA name from CA-Lookup
    const caMapping = readCALookup();
    let userCAName = null;

    for (const [accountId, caInfo] of Object.entries(caMapping)) {
      if (caInfo.email && caInfo.email.toLowerCase() === userEmail.toLowerCase()) {
        userCAName = caInfo.name;
        break;
      }
    }

    if (!userCAName) {
      return {
        growers: [],
        shrinkers: [],
        tabName: weeklyTabName,
        lastUpdated: new Date().toISOString(),
        error: 'User not found in CA-Lookup'
      };
    }

    // Read data from tab
    // Growers: rows 3-13 (header row 2, data rows 3-13)
    // Shrinkers: rows 17-27 (header row 16, data rows 17-27)
    const growers = extractAccountsFromSection(tab, 3, 13, userCAName, 'Growers');
    const shrinkers = extractAccountsFromSection(tab, 17, 27, userCAName, 'Shrinkers');

    return {
      growers: growers,
      shrinkers: shrinkers,
      tabName: weeklyTabName,
      lastUpdated: new Date().toISOString()
    };

  } catch (error) {
    Logger.log(`Error in getUserAccounts: ${error.message}`);
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
 * Helper function to extract accounts from a section (Growers or Shrinkers)
 * @param {Sheet} sheet - The sheet to read from
 * @param {number} startRow - Start row (1-indexed)
 * @param {number} endRow - End row (1-indexed)
 * @param {string} caName - CA name to filter by
 * @param {string} section - Section name ('Growers' or 'Shrinkers')
 * @returns {Array} Array of account objects
 */
function extractAccountsFromSection(sheet, startRow, endRow, caName, section) {
  const accounts = [];

  // Column structure (1-indexed):
  // Columns 1-12: Base data from source
  // Column 13: CA Forecast
  // Column 14: CA Forecast Variance
  // Column 15: CA Forecast Var %
  // Column 16: Customer Architect
  // Column 17: Notes

  for (let row = startRow; row <= endRow; row++) {
    const rowData = sheet.getRange(row, 1, 1, 17).getValues()[0];

    // Array is 0-indexed, so column N is at index N-1
    const accountId = rowData[2] ? rowData[2].toString().trim() : '';       // Column 3
    const accountName = rowData[3] ? rowData[3].toString().trim() : '';     // Column 4
    const area = rowData[4] ? rowData[4].toString().trim() : '';            // Column 5
    const wowPercent = rowData[9];                                          // Column 10
    const momPercent = rowData[10];                                         // Column 11
    const caForecast = rowData[12];                                         // Column 13
    const caForecastVar = rowData[13];                                      // Column 14
    const caForecastVarPct = rowData[14];                                   // Column 15
    const customerArchitect = rowData[15] ? rowData[15].toString().trim() : ''; // Column 16
    const notes = rowData[16] ? rowData[16].toString().trim() : '';         // Column 17

    // Filter by Customer Architect column
    if (customerArchitect === caName && accountName) {
      accounts.push({
        accountId: accountId,
        accountName: accountName,
        area: area,
        wowPercent: wowPercent,
        momPercent: momPercent,
        caForecast: caForecast,
        caForecastVar: caForecastVar,
        caForecastVarPct: caForecastVarPct,
        currentNote: notes,
        rowNumber: row,
        section: section
      });
    }
  }

  return accounts;
}

// ============================================
// SAVE OPERATIONS
// ============================================

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

    // Check Customer Architect column (column 16) matches user's CA name
    const customerArchitect = tab.getRange(rowNumber, 16).getValue()?.toString().trim();

    if (customerArchitect !== userInfo.caName) {
      throw new Error('Not authorized to edit this row');
    }

    // Update Notes column (column 17)
    tab.getRange(rowNumber, 17).setValue(sanitizedNote);

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

// ============================================
// MENU INTEGRATION
// ============================================

/**
 * Shows the CA Notes Interface sidebar
 */
function showNotesInterface() {
  try {
    const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Weekly Forecast Notes')
      .setWidth(300);

    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.log(`Error in showNotesInterface: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error',
      `Failed to open interface: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function to include HTML files
 * @param {string} filename - Name of the file to include
 * @returns {string} File content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
