/**
 * Weekly Global Forecast Data Transfer Script
 *
 * This script:
 * 1. Reads Growers and Shrinkers tables from Global sheet
 * 2. Looks up Customer Architects from CA-Lookup (includes manager info)
 * 3. Creates a new dated tab in target sheet
 * 4. Writes combined data with CA assignments
 * 5. Adds @mention comments to notify CAs
 * 6. Groups accounts by CA manager and sends email alerts to managers
 *
 * Manager emails are dynamically pulled from CA-Lookup (Column F: CA Manager Email)
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  GLOBAL_SHEET_ID: '1YkOxHI5EZjVoCKJuifn2tZGZVj7gp_spxEv0YIAFVKw',
  TARGET_SHEET_ID: '1rsQUBs1nD8Ywve88910jGcObJRiZnxSL0gHEK9RMrBA',
  GLOBAL_TAB_NAME: 'Global',
  CA_LOOKUP_TAB_NAME: 'CA-Lookup',
  EXECUTION_LOG_TAB_NAME: 'ExecutionLog',
  EMAIL_DOMAIN: '@elastic.co',

  // TEST MODE - Set to true to only mention ryan.burnham@elastic.co
  TEST_MODE: true,
  TEST_EMAIL: 'ryan.burnham@elastic.co',

  // EMAIL CC RECIPIENTS
  SR_DIRECTOR_EMAIL: 'dan.owen@elastic.co',

  // Table ranges in Global sheet
  GROWERS_START_ROW: 25,
  GROWERS_END_ROW: 36,
  SHRINKERS_START_ROW: 39,
  SHRINKERS_END_ROW: 50,

  // Column range to copy (B through M)
  START_COLUMN: 2,  // Column B (1-indexed)
  END_COLUMN: 13,   // Column M (1-indexed)

  // Column indices (0-based after reading B-M)
  ACCOUNT_ID_COL: 1,      // Column C in original sheet, becomes index 1 after reading B-M
  ACCOUNT_NAME_COL: 2,    // Column D in original sheet, becomes index 2 after reading B-M
  AREA_COL: 3,            // Column E in original sheet, becomes index 3 after reading B-M
  WOW_PERCENT_COL: 8      // Column J in original sheet, becomes index 8 after reading B-M
};

// ============================================
// MAIN FUNCTION
// ============================================

/**
 * Main function to run weekly
 * Can be triggered manually or set up with a time-based trigger
 */
function weeklyForecastUpdate() {
  const startTime = new Date();
  let executionStatus = 'Running';
  let tabName = '';
  let totalAccounts = 0;
  let emailsSent = 0;
  let errorMessage = '';

  try {
    Logger.log('Starting weekly forecast update...');

    if (CONFIG.TEST_MODE) {
      Logger.log('⚠️ TEST MODE ENABLED - All @mentions will use ' + CONFIG.TEST_EMAIL);
    }

    // 1. Read CA-Lookup mapping
    Logger.log('Reading CA-Lookup...');
    const caMapping = readCALookup();
    Logger.log(`Found ${Object.keys(caMapping).length} CA mappings`);

    // 2. Read Global sheet data
    Logger.log('Reading Global sheet data...');
    const globalData = readGlobalSheetData();
    Logger.log(`Read ${globalData.growers.length} growers and ${globalData.shrinkers.length} shrinkers`);

    // 3. Create new tab with date name
    Logger.log('Creating new tab...');
    const newTab = createDatedTab();
    tabName = newTab.getName();
    Logger.log(`Created tab: ${tabName}`);

    // 4. Write combined data
    Logger.log('Writing data to new tab...');
    const writeResult = writeCombinedData(newTab, globalData, caMapping);
    Logger.log(`Wrote ${writeResult.rowCount} rows of data`);

    // 5. Add notes to cells
    Logger.log('Adding notes to cells...');
    const caAccounts = addNotesToCells(newTab, writeResult);
    Logger.log('Notes added successfully');

    // 6. Group accounts by CA manager
    Logger.log('Grouping accounts by CA manager...');
    const managerAccounts = groupAccountsByManager(caAccounts, caMapping);
    Logger.log(`Found ${Object.keys(managerAccounts).length} managers with accounts`);

    // 7. Send email notifications to CA managers
    Logger.log('Sending email notifications to CA managers...');
    sendManagerNotifications(newTab, managerAccounts);
    emailsSent = Object.keys(managerAccounts).length;
    totalAccounts = Object.values(managerAccounts).reduce((sum, accounts) => sum + accounts.length, 0);
    Logger.log('CA manager notifications sent successfully');

    executionStatus = 'Success';
    Logger.log('Weekly forecast update completed successfully!');
    showAlert('Success!',
      `Weekly forecast data has been copied to tab: ${tabName}\nNotifications sent to ${emailsSent} territory manager(s) for ${totalAccounts} accounts.`);

  } catch (error) {
    executionStatus = 'Failed';
    errorMessage = error.message;
    Logger.log(`Error in weeklyForecastUpdate: ${error.message}`);
    Logger.log(error.stack);
    showAlert('Error',
      `An error occurred: ${error.message}\n\nCheck the script logs for details.`);
    throw error;
  } finally {
    // Log execution to ExecutionLog sheet
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000; // seconds
    logExecution({
      timestamp: startTime,
      tabName: tabName,
      status: executionStatus,
      totalAccounts: totalAccounts,
      emailsSent: emailsSent,
      duration: duration,
      testMode: CONFIG.TEST_MODE,
      error: errorMessage
    });
  }
}

// ============================================
// DATA READING FUNCTIONS
// ============================================

/**
 * Reads CA-Lookup tab and returns mapping object with manager info
 * @returns {Object} Mapping of account_id to CA info (name, email, manager, managerEmail)
 */
function readCALookup() {
  const targetSheet = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const caLookupTab = targetSheet.getSheetByName(CONFIG.CA_LOOKUP_TAB_NAME);

  if (!caLookupTab) {
    throw new Error(`CA-Lookup tab not found in target sheet`);
  }

  const data = caLookupTab.getDataRange().getValues();
  const mapping = {};

  for (let i = 1; i < data.length; i++) {
    const accountId = data[i][0]?.toString().trim();       // Column A: Account ID
    const accountName = data[i][1]?.toString().trim();     // Column B: Account Name
    const caName = data[i][2]?.toString().trim();          // Column C: Customer Architect
    const caEmail = data[i][3]?.toString().trim();         // Column D: CA Email
    const caManager = data[i][4]?.toString().trim();       // Column E: CA Manager
    const caManagerEmail = data[i][5]?.toString().trim();  // Column F: CA Manager Email

    if (accountId && caName) {
      mapping[accountId] = {
        name: caName,
        email: caEmail || '',
        manager: caManager || '',
        managerEmail: caManagerEmail || ''
      };
    }
  }

  return mapping;
}

/**
 * Reads Growers and Shrinkers tables from Global sheet (columns B-M only)
 * @returns {Object} Object with growers and shrinkers data and formatting
 */
function readGlobalSheetData() {
  const globalSheet = SpreadsheetApp.openById(CONFIG.GLOBAL_SHEET_ID);
  const globalTab = globalSheet.getSheetByName(CONFIG.GLOBAL_TAB_NAME);

  if (!globalTab) {
    throw new Error(`Global tab not found in source sheet`);
  }

  const numColumns = CONFIG.END_COLUMN - CONFIG.START_COLUMN + 1;

  // Read growers table (including header) - columns B through M
  const growersRange = globalTab.getRange(
    CONFIG.GROWERS_START_ROW,
    CONFIG.START_COLUMN,
    CONFIG.GROWERS_END_ROW - CONFIG.GROWERS_START_ROW + 1,
    numColumns
  );
  const growersData = growersRange.getValues();
  const growersBackgrounds = growersRange.getBackgrounds();
  const growersFontColors = growersRange.getFontColors();
  const growersNumberFormats = growersRange.getNumberFormats();
  const growersFontWeights = growersRange.getFontWeights();

  // Read shrinkers table (including header) - columns B through M
  const shrinkersRange = globalTab.getRange(
    CONFIG.SHRINKERS_START_ROW,
    CONFIG.START_COLUMN,
    CONFIG.SHRINKERS_END_ROW - CONFIG.SHRINKERS_START_ROW + 1,
    numColumns
  );
  const shrinkersData = shrinkersRange.getValues();
  const shrinkersBackgrounds = shrinkersRange.getBackgrounds();
  const shrinkersFontColors = shrinkersRange.getFontColors();
  const shrinkersNumberFormats = shrinkersRange.getNumberFormats();
  const shrinkersFontWeights = shrinkersRange.getFontWeights();

  return {
    growers: growersData,
    growersFormatting: {
      backgrounds: growersBackgrounds,
      fontColors: growersFontColors,
      numberFormats: growersNumberFormats,
      fontWeights: growersFontWeights
    },
    shrinkers: shrinkersData,
    shrinkersFormatting: {
      backgrounds: shrinkersBackgrounds,
      fontColors: shrinkersFontColors,
      numberFormats: shrinkersNumberFormats,
      fontWeights: shrinkersFontWeights
    }
  };
}

// ============================================
// TAB CREATION FUNCTION
// ============================================

/**
 * Creates a new tab with date name, appending version if needed
 * @returns {Sheet} The newly created sheet
 */
function createDatedTab() {
  const targetSheet = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let tabName = dateStr;
  let version = 2;

  // Check if tab already exists, append version number if needed
  while (targetSheet.getSheetByName(tabName) !== null) {
    tabName = `${dateStr}_v${version}`;
    version++;
  }

  const newTab = targetSheet.insertSheet(tabName);

  // Move to first position (optional)
  targetSheet.setActiveSheet(newTab);
  targetSheet.moveActiveSheet(1);

  return newTab;
}

// ============================================
// DATA WRITING FUNCTION
// ============================================

/**
 * Writes combined data to the new tab with formatting
 * @param {Sheet} sheet - The target sheet
 * @param {Object} data - Object with growers, shrinkers data and formatting
 * @param {Object} caMapping - CA name mapping
 * @returns {Object} Information about written data
 */
function writeCombinedData(sheet, data, caMapping) {
  const combinedData = [];
  const combinedBackgrounds = [];
  const combinedFontColors = [];
  const combinedNumberFormats = [];
  const combinedFontWeights = [];

  // Calculate the correct number of columns (data columns + CA + Notes)
  const maxColumns = data.growers[0].length + 2;

  // Default formatting for new columns
  const defaultBackground = '#ffffff';
  const defaultFontColor = '#000000';
  const defaultNumberFormat = '@';  // Plain text
  const defaultFontWeight = 'normal';

  // Add Growers section header
  combinedData.push(['BIGGEST 10 GROWERS WoW']);
  combinedBackgrounds.push([defaultBackground]);
  combinedFontColors.push([defaultFontColor]);
  combinedNumberFormats.push([defaultNumberFormat]);
  combinedFontWeights.push(['bold']);

  // Add growers header row with CA and Notes columns
  const growersHeader = [...data.growers[0], 'Customer Architect', 'Notes'];
  combinedData.push(growersHeader);
  combinedBackgrounds.push([...data.growersFormatting.backgrounds[0], defaultBackground, defaultBackground]);
  combinedFontColors.push([...data.growersFormatting.fontColors[0], defaultFontColor, defaultFontColor]);
  combinedNumberFormats.push([...data.growersFormatting.numberFormats[0], defaultNumberFormat, defaultNumberFormat]);
  combinedFontWeights.push([...data.growersFormatting.fontWeights[0], 'bold', 'bold']);

  // Add growers data rows
  for (let i = 1; i < data.growers.length; i++) {
    const row = data.growers[i];
    const accountId = row[CONFIG.ACCOUNT_ID_COL]?.toString().trim();
    const accountName = row[CONFIG.ACCOUNT_NAME_COL];
    const caName = caMapping[accountId]?.name || '';

    Logger.log(`GROWER [${i}]: Account ID="${accountId}", Name="${accountName}", CA Found=${!!caName ? 'YES (' + caName + ')' : 'NO'}`);

    if (!caName && accountId) {
      Logger.log(`  ⚠️ No CA mapping found for this account ID`);
    }

    combinedData.push([...row, caName, '']);
    combinedBackgrounds.push([...data.growersFormatting.backgrounds[i], defaultBackground, defaultBackground]);
    combinedFontColors.push([...data.growersFormatting.fontColors[i], defaultFontColor, defaultFontColor]);
    combinedNumberFormats.push([...data.growersFormatting.numberFormats[i], defaultNumberFormat, defaultNumberFormat]);
    combinedFontWeights.push([...data.growersFormatting.fontWeights[i], defaultFontWeight, defaultFontWeight]);
  }

  // Add blank row separator
  combinedData.push([]);
  combinedBackgrounds.push([]);
  combinedFontColors.push([]);
  combinedNumberFormats.push([]);
  combinedFontWeights.push([]);

  // Add Shrinkers section header
  combinedData.push(['BIGGEST 10 SHRINKERS WoW']);
  combinedBackgrounds.push([defaultBackground]);
  combinedFontColors.push([defaultFontColor]);
  combinedNumberFormats.push([defaultNumberFormat]);
  combinedFontWeights.push(['bold']);

  // Add shrinkers header row with CA and Notes columns
  const shrinkersHeader = [...data.shrinkers[0], 'Customer Architect', 'Notes'];
  combinedData.push(shrinkersHeader);
  combinedBackgrounds.push([...data.shrinkersFormatting.backgrounds[0], defaultBackground, defaultBackground]);
  combinedFontColors.push([...data.shrinkersFormatting.fontColors[0], defaultFontColor, defaultFontColor]);
  combinedNumberFormats.push([...data.shrinkersFormatting.numberFormats[0], defaultNumberFormat, defaultNumberFormat]);
  combinedFontWeights.push([...data.shrinkersFormatting.fontWeights[0], 'bold', 'bold']);

  // Add shrinkers data rows
  for (let i = 1; i < data.shrinkers.length; i++) {
    const row = data.shrinkers[i];
    const accountId = row[CONFIG.ACCOUNT_ID_COL]?.toString().trim();
    const accountName = row[CONFIG.ACCOUNT_NAME_COL];
    const caName = caMapping[accountId]?.name || '';

    Logger.log(`SHRINKER [${i}]: Account ID="${accountId}", Name="${accountName}", CA Found=${!!caName ? 'YES (' + caName + ')' : 'NO'}`);

    if (!caName && accountId) {
      Logger.log(`  ⚠️ No CA mapping found for this account ID`);
    }

    combinedData.push([...row, caName, '']);
    combinedBackgrounds.push([...data.shrinkersFormatting.backgrounds[i], defaultBackground, defaultBackground]);
    combinedFontColors.push([...data.shrinkersFormatting.fontColors[i], defaultFontColor, defaultFontColor]);
    combinedNumberFormats.push([...data.shrinkersFormatting.numberFormats[i], defaultNumberFormat, defaultNumberFormat]);
    combinedFontWeights.push([...data.shrinkersFormatting.fontWeights[i], defaultFontWeight, defaultFontWeight]);
  }

  // Pad all rows to the same column width
  for (let i = 0; i < combinedData.length; i++) {
    while (combinedData[i].length < maxColumns) {
      combinedData[i].push('');
      combinedBackgrounds[i].push(defaultBackground);
      combinedFontColors[i].push(defaultFontColor);
      combinedNumberFormats[i].push(defaultNumberFormat);
      combinedFontWeights[i].push(defaultFontWeight);
    }
  }

  // Write all data to sheet
  const dataRange = sheet.getRange(1, 1, combinedData.length, maxColumns);
  dataRange.setValues(combinedData);

  // Apply formatting
  dataRange.setBackgrounds(combinedBackgrounds);
  dataRange.setFontColors(combinedFontColors);
  dataRange.setNumberFormats(combinedNumberFormats);
  dataRange.setFontWeights(combinedFontWeights);

  // Freeze header row
  sheet.setFrozenRows(2);

  return {
    rowCount: combinedData.length,
    growersDataStartRow: 4,  // Row 4 (first data row after headers)
    growersDataEndRow: 14,  // Process through row 14 to ensure we get all growers
    shrinkersDataStartRow: 18,  // Row 18 (first shrinker data row after headers)
    shrinkersDataEndRow: 30,  // Process through row 30 to ensure we get all shrinkers
    notesColumn: maxColumns  // Last column
  };
}

// ============================================
// NOTES AND NOTIFICATIONS
// ============================================

/**
 * Adds notes to cells and collects CA account assignments
 * @param {Sheet} sheet - The target sheet
 * @param {Object} writeResult - Information about written data
 * @returns {Object} Mapping of CA email to their accounts
 */
function addNotesToCells(sheet, writeResult) {
  const caAccounts = {};  // {email: [{name, row, section}]}

  // Process growers
  for (let i = writeResult.growersDataStartRow; i <= writeResult.growersDataEndRow; i++) {
    const accountInfo = addNoteForRow(sheet, i, writeResult.notesColumn, 'Growers');
    if (accountInfo) {
      if (!caAccounts[accountInfo.email]) {
        caAccounts[accountInfo.email] = [];
      }
      caAccounts[accountInfo.email].push(accountInfo);
    }
  }

  // Process shrinkers
  for (let i = writeResult.shrinkersDataStartRow; i <= writeResult.shrinkersDataEndRow; i++) {
    const accountInfo = addNoteForRow(sheet, i, writeResult.notesColumn, 'Shrinkers');
    if (accountInfo) {
      if (!caAccounts[accountInfo.email]) {
        caAccounts[accountInfo.email] = [];
      }
      caAccounts[accountInfo.email].push(accountInfo);
    }
  }

  return caAccounts;
}

/**
 * Adds a note for a specific row and returns account info
 * @param {Sheet} sheet - The target sheet
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed)
 * @param {string} section - 'Growers' or 'Shrinkers'
 * @returns {Object|null} Account info or null if skipped
 */
function addNoteForRow(sheet, row, col, section) {
  const accountNameCol = CONFIG.ACCOUNT_NAME_COL + 1;  // Convert to 1-indexed
  const areaCol = CONFIG.AREA_COL + 1;  // Convert to 1-indexed
  const wowPercentCol = CONFIG.WOW_PERCENT_COL + 1;  // Convert to 1-indexed
  const caCol = col - 1;  // CA is one column before Notes

  const area = sheet.getRange(row, areaCol).getValue()?.toString().trim();

  // Only process AMER region
  if (!area || !area.includes('AMER')) {
    Logger.log(`Skipping row ${row} - not in AMER region (area: ${area})`);
    return null;
  }

  const accountName = cleanCompanyName(sheet.getRange(row, accountNameCol).getValue());
  const caName = sheet.getRange(row, caCol).getValue()?.toString().trim();
  const wowPercent = sheet.getRange(row, wowPercentCol).getValue();

  if (caName) {
    // Use test email if in test mode, otherwise use actual CA email
    const email = CONFIG.TEST_MODE ? CONFIG.TEST_EMAIL : convertNameToEmail(caName);

    // Create note with @mention
    const noteText = CONFIG.TEST_MODE
      ? `[TEST MODE] @${email}\nPlease add notes for ${accountName}\n\nActual CA: ${caName}`
      : `@${email}\nPlease add notes for ${accountName}`;

    const notesCell = sheet.getRange(row, col);
    notesCell.setNote(noteText);

    Logger.log(`Added note for ${accountName} -> ${caName} (@${email})`);

    return {
      email: email,
      caName: caName,
      accountName: accountName,
      row: row,
      section: section,
      area: area,
      wowPercent: wowPercent
    };
  } else {
    Logger.log(`Skipping row ${row} - no CA name found`);
    return null;
  }
}

/**
 * Sends email notifications to CAs with their assigned accounts
 * @param {Sheet} sheet - The target sheet
 * @param {Object} caAccounts - Mapping of CA email to their accounts
 */
function sendCANotifications(sheet, caAccounts) {
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.TARGET_SHEET_ID}/edit#gid=${sheet.getSheetId()}`;
  const tabName = sheet.getName();

  for (const [email, accounts] of Object.entries(caAccounts)) {
    if (accounts.length === 0) continue;

    // Get unique CA name (should be same for all accounts)
    const caName = accounts[0].caName;

    // Group by section
    const growers = accounts.filter(a => a.section === 'Growers');
    const shrinkers = accounts.filter(a => a.section === 'Shrinkers');

    // Build CC list based on areas
    const ccList = buildCCList(accounts);

    // Build email body
    let emailBody = CONFIG.TEST_MODE
      ? `[TEST MODE - This email would normally go to ${caName}]\n\n`
      : '';

    emailBody += `Hi ${caName.split(' ')[0]},\n\n`;
    emailBody += `You have ${accounts.length} account${accounts.length > 1 ? 's' : ''} assigned in this week's forecast that need notes added.\n\n`;
    emailBody += `Please review and add notes in the sheet:\n${spreadsheetUrl}\n\n`;

    if (growers.length > 0) {
      emailBody += `📈 GROWERS (${growers.length}):\n`;
      growers.forEach(acc => {
        const percentDisplay = formatPercent(acc.wowPercent);
        emailBody += `  • ${acc.accountName} (${percentDisplay})\n`;
      });
      emailBody += '\n';
    }

    if (shrinkers.length > 0) {
      emailBody += `📉 SHRINKERS (${shrinkers.length}):\n`;
      shrinkers.forEach(acc => {
        const percentDisplay = formatPercent(acc.wowPercent);
        emailBody += `  • ${acc.accountName} (${percentDisplay})\n`;
      });
      emailBody += '\n';
    }

    emailBody += `Tab: ${tabName}\n\n`;
    emailBody += `Thank you!`;

    const subject = CONFIG.TEST_MODE
      ? `[TEST] Weekly Forecast - Action Required (${accounts.length} account${accounts.length > 1 ? 's' : ''})`
      : `Weekly Forecast - Action Required (${accounts.length} account${accounts.length > 1 ? 's' : ''})`;

    try {
      const emailParams = {
        to: email,
        subject: subject,
        body: emailBody
      };

      // Add CC in test mode or production
      if (ccList.length > 0) {
        emailParams.cc = CONFIG.TEST_MODE ? CONFIG.TEST_EMAIL : ccList.join(',');
        if (CONFIG.TEST_MODE) {
          emailBody = `[CC would include: ${ccList.join(', ')}]\n\n` + emailBody;
          emailParams.body = emailBody;
        }
      }

      MailApp.sendEmail(emailParams);
      Logger.log(`Sent email to ${email} (${caName}) for ${accounts.length} accounts, CC: ${ccList.join(', ')}`);
    } catch (error) {
      Logger.log(`Failed to send email to ${email}: ${error.message}`);
    }
  }

  Logger.log(`Sent notifications to ${Object.keys(caAccounts).length} CA(s)`);
}

/**
 * Builds CC list based on account areas
 * @param {Array} accounts - Array of account objects
 * @returns {Array} List of email addresses to CC
 */
function buildCCList(accounts) {
  const ccSet = new Set();

  // Always CC Sr Director
  ccSet.add(CONFIG.SR_DIRECTOR_EMAIL);

  // Add area managers for unique areas
  const uniqueAreas = [...new Set(accounts.map(a => a.area))];
  uniqueAreas.forEach(area => {
    if (CONFIG.AREA_MANAGERS[area]) {
      ccSet.add(CONFIG.AREA_MANAGERS[area]);
    }
  });

  return Array.from(ccSet);
}

/**
 * Groups accounts by CA manager email from CA-Lookup
 * @param {Object} caAccounts - Mapping of CA email to their accounts
 * @param {Object} caMapping - CA lookup mapping with manager info
 * @returns {Object} Mapping of manager email to accounts with CA info
 */
function groupAccountsByManager(caAccounts, caMapping) {
  const managerAccounts = {};

  // Iterate through all CA accounts and regroup by their manager
  for (const accounts of Object.values(caAccounts)) {
    accounts.forEach(account => {
      // Get manager email from CA-Lookup
      const caInfo = caMapping[account.accountId];
      const managerEmail = caInfo?.managerEmail;
      const managerName = caInfo?.manager || 'Manager';

      if (!managerEmail) {
        Logger.log(`No manager email found for account ${account.accountId} (${account.accountName})`);
        return;
      }

      if (!managerAccounts[managerEmail]) {
        managerAccounts[managerEmail] = {
          managerName: managerName,
          accounts: []
        };
      }

      managerAccounts[managerEmail].accounts.push(account);
    });
  }

  return managerAccounts;
}

/**
 * Sends email notifications to CA managers with their team's accounts
 * @param {Sheet} sheet - The target sheet
 * @param {Object} managerAccounts - Mapping of manager email to accounts
 */
function sendManagerNotifications(sheet, managerAccounts) {
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.TARGET_SHEET_ID}/edit#gid=${sheet.getSheetId()}`;
  const tabName = sheet.getName();

  for (const [managerEmail, data] of Object.entries(managerAccounts)) {
    const accounts = data.accounts;
    const managerName = data.managerName;

    if (accounts.length === 0) continue;

    // Use test email if in test mode
    const toEmail = CONFIG.TEST_MODE ? CONFIG.TEST_EMAIL : managerEmail;

    // Group by section
    const growers = accounts.filter(a => a.section === 'Growers');
    const shrinkers = accounts.filter(a => a.section === 'Shrinkers');

    // Group by territory for better organization
    const territoryCounts = {};
    accounts.forEach(acc => {
      territoryCounts[acc.area] = (territoryCounts[acc.area] || 0) + 1;
    });
    const territories = Object.keys(territoryCounts).sort();

    // Build email body
    let emailBody = CONFIG.TEST_MODE
      ? `[TEST MODE - This email would normally go to ${managerName} at ${managerEmail}]\n\n`
      : '';

    emailBody += `Hi ${managerName.split(' ')[0]},\n\n`;
    emailBody += `This is your weekly forecast alert for your team's accounts.\n\n`;
    emailBody += `Your team has ${accounts.length} account${accounts.length > 1 ? 's' : ''} that require attention this week`;

    if (territories.length > 0) {
      emailBody += ` across ${territories.length} territor${territories.length > 1 ? 'ies' : 'y'}: ${territories.join(', ')}`;
    }
    emailBody += `.\n\n`;

    emailBody += `Please review the full details in the sheet:\n${spreadsheetUrl}\n\n`;

    if (growers.length > 0) {
      emailBody += `📈 GROWERS (${growers.length}):\n`;
      growers.forEach(acc => {
        const percentDisplay = formatPercent(acc.wowPercent);
        emailBody += `  • ${acc.accountName} (${acc.area})\n`;
        emailBody += `    CA: ${acc.caName}\n`;
        emailBody += `    WoW: ${percentDisplay}\n`;
      });
      emailBody += '\n';
    }

    if (shrinkers.length > 0) {
      emailBody += `📉 SHRINKERS (${shrinkers.length}):\n`;
      shrinkers.forEach(acc => {
        const percentDisplay = formatPercent(acc.wowPercent);
        emailBody += `  • ${acc.accountName} (${acc.area})\n`;
        emailBody += `    CA: ${acc.caName}\n`;
        emailBody += `    WoW: ${percentDisplay}\n`;
      });
      emailBody += '\n';
    }

    emailBody += `Tab: ${tabName}\n\n`;
    emailBody += `Thank you!`;

    const subject = CONFIG.TEST_MODE
      ? `[TEST] Weekly Forecast Alert - Your Team (${accounts.length} account${accounts.length > 1 ? 's' : ''})`
      : `Weekly Forecast Alert - Your Team (${accounts.length} account${accounts.length > 1 ? 's' : ''})`;

    try {
      const emailParams = {
        to: toEmail,
        subject: subject,
        body: emailBody
      };

      // CC Sr Director
      emailParams.cc = CONFIG.TEST_MODE ? CONFIG.TEST_EMAIL : CONFIG.SR_DIRECTOR_EMAIL;

      if (CONFIG.TEST_MODE) {
        emailBody = `[CC would include: ${CONFIG.SR_DIRECTOR_EMAIL}]\n\n` + emailBody;
        emailParams.body = emailBody;
      }

      MailApp.sendEmail(emailParams);
      Logger.log(`Sent email to ${toEmail} for ${managerName} (${accounts.length} accounts)`);
    } catch (error) {
      Logger.log(`Failed to send email to ${managerName}: ${error.message}`);
    }
  }

  Logger.log(`Sent notifications to ${Object.keys(managerAccounts).length} manager(s)`);
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Cleans company name by removing asterisk prefix
 * @param {string} name - Company name
 * @returns {string} Cleaned name
 */
function cleanCompanyName(name) {
  if (!name) return '';
  return name.toString().trim().replace(/^\*\s*/, '');
}

/**
 * Converts CA name to email address
 * @param {string} name - CA name (e.g., "John Smith")
 * @returns {string} Email address (e.g., "john.smith@elastic.co")
 */
function convertNameToEmail(name) {
  if (!name) return '';

  const parts = name.trim().toLowerCase().split(/\s+/);

  if (parts.length === 0) return '';
  if (parts.length === 1) return `${parts[0]}${CONFIG.EMAIL_DOMAIN}`;

  // firstname.lastname format
  const firstName = parts[0];
  const lastName = parts[parts.length - 1];

  return `${firstName}.${lastName}${CONFIG.EMAIL_DOMAIN}`;
}

/**
 * Safely shows an alert dialog (works in both manual and trigger contexts)
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showAlert(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (error) {
    // Running from trigger context where UI is not available
    Logger.log(`Alert [${title}]: ${message}`);
  }
}

/**
 * Formats a percentage value for display
 * @param {number|string} value - The percentage value (stored as decimal in Google Sheets)
 * @returns {string} Formatted percentage string
 */
function formatPercent(value) {
  if (value === null || value === undefined || value === '') {
    return 'N/A';
  }

  // If it's already a string with %, return as is
  if (typeof value === 'string' && value.includes('%')) {
    return value;
  }

  // Convert to number
  const numValue = typeof value === 'number' ? value : parseFloat(value);

  if (isNaN(numValue)) {
    return 'N/A';
  }

  // Google Sheets stores percentages as decimals, so always multiply by 100
  // 10.25 in sheet = 1025%, 0.53 = 53%, 0.05 = 5%
  const percentValue = numValue * 100;

  // Format with sign
  const sign = percentValue >= 0 ? '+' : '';
  return `${sign}${percentValue.toFixed(0)}%`;
}

// ============================================
// EXECUTION LOGGING
// ============================================

/**
 * Logs execution details to ExecutionLog sheet
 * @param {Object} logData - Execution data to log
 */
function logExecution(logData) {
  try {
    const targetSheet = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
    let logTab = targetSheet.getSheetByName(CONFIG.EXECUTION_LOG_TAB_NAME);

    // Create ExecutionLog tab if it doesn't exist
    if (!logTab) {
      logTab = targetSheet.insertSheet(CONFIG.EXECUTION_LOG_TAB_NAME);

      // Add headers
      const headers = [
        'Timestamp',
        'Tab Name',
        'Status',
        'Total Accounts',
        'Emails Sent',
        'Duration (sec)',
        'Test Mode',
        'Error Message'
      ];
      logTab.getRange(1, 1, 1, headers.length).setValues([headers]);
      logTab.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      logTab.setFrozenRows(1);
    }

    // Append new log entry
    const row = [
      logData.timestamp,
      logData.tabName || 'N/A',
      logData.status,
      logData.totalAccounts || 0,
      logData.emailsSent || 0,
      logData.duration ? logData.duration.toFixed(2) : 0,
      logData.testMode ? 'Yes' : 'No',
      logData.error || ''
    ];

    logTab.appendRow(row);

    // Format timestamp column
    const lastRow = logTab.getLastRow();
    logTab.getRange(lastRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    // Color code status
    const statusCell = logTab.getRange(lastRow, 3);
    if (logData.status === 'Success') {
      statusCell.setBackground('#d9ead3'); // Light green
    } else if (logData.status === 'Failed') {
      statusCell.setBackground('#f4cccc'); // Light red
    }

    Logger.log(`Execution logged to ${CONFIG.EXECUTION_LOG_TAB_NAME}`);

  } catch (error) {
    Logger.log(`Failed to log execution: ${error.message}`);
    // Don't throw - logging failure shouldn't break the main execution
  }
}

// ============================================
// MENU FUNCTION (Optional)
// ============================================

/**
 * Creates custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Weekly Forecast')
    .addItem('Run Weekly Update', 'weeklyForecastUpdate')
    .addSeparator()
    .addItem('Test CA Lookup', 'testCALookup')
    .addToUi();
}

/**
 * Test function to verify CA lookup is working
 */
function testCALookup() {
  try {
    const mapping = readCALookup();
    Logger.log('CA Lookup Test Results:');
    Logger.log(`Found ${Object.keys(mapping).length} mappings`);

    for (const [accountId, ca] of Object.entries(mapping)) {
      const email = convertNameToEmail(ca);
      Logger.log(`${accountId} → ${ca} (${email})`);
    }

    showAlert('Test Complete',
      `Found ${Object.keys(mapping).length} CA mappings.\n\nCheck the script logs for details.`);
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    showAlert('Error', error.message);
  }
}
