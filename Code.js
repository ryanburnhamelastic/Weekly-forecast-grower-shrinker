/**
 * Weekly Global Forecast - Combined Script
 *
 * This script combines two operations:
 * 1. SPREADSHEET: Reads Growers and Shrinkers from Global sheet, creates dated tab, sends CA notifications
 * 2. DOCUMENT: Populates Google Doc with growers/shrinkers data organized by area
 *
 * Execution order: Spreadsheet (primary) → Document (supplementary)
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  // Source spreadsheet
  GLOBAL_SHEET_ID: '1YkOxHI5EZjVoCKJuifn2tZGZVj7gp_spxEv0YIAFVKw',
  GLOBAL_TAB_NAME: 'Global',

  // Target destinations (PRODUCTION)
  TARGET_SHEET_ID: '1rsQUBs1nD8Ywve88910jGcObJRiZnxSL0gHEK9RMrBA',
  TARGET_DOCUMENT_ID: '1TPysiit9ZQ5YRvMIaJkb5JV_2LlMXUvD7Jf8nDbL2i4',

  // Test destinations (used when TEST_MODE = true)
  TEST_TARGET_SHEET_ID: '1qr62xOa6OJ5y9dW-Ge5147IgqeetHOXmtWSmYdagSdY',
  TEST_TARGET_DOCUMENT_ID: '114Gg__2GMoRYMBqKQyN69znfaKGPL7LhygGkRlKXY5w',

  // Spreadsheet tabs
  CA_LOOKUP_TAB_NAME: 'CA-Lookup',
  EXECUTION_LOG_TAB_NAME: 'ExecutionLog',

  // Email configuration
  EMAIL_DOMAIN: '@elastic.co',
  TEST_MODE: false,
  TEST_EMAIL: 'ryan.burnham@elastic.co',
  SR_DIRECTOR_EMAIL: 'dan.owen@elastic.co',

  // Area code to document section mapping
  AREA_SECTION_MAP: {
    'AMER_STRATEGICS': 'Strat',
    'AMER_ENT_EAST': 'Ent-East / Canada / LATAM',
    'AMER_CAN': 'Ent-East / Canada / LATAM',
    'AMER_LATAM': 'Ent-East / Canada / LATAM',
    'AMER_ENT_WEST': 'Ent-West',
    'AMER_MIDMKT_GENBUS': 'MM / GB'
  },

  // Table ranges in Global sheet (shared by both operations)
  GROWERS_START_ROW: 25,
  GROWERS_END_ROW: 36,
  SHRINKERS_START_ROW: 39,
  SHRINKERS_END_ROW: 50,

  // Column range (B through R to include CA Forecast columns)
  START_COLUMN: 2,
  END_COLUMN: 18,

  // Column indices (0-based after reading from column B)
  ACCOUNT_ID_COL: 1,
  ACCOUNT_NAME_COL: 2,
  AREA_COL: 3,
  WOW_PERCENT_COL: 8,
  MOM_PERCENT_COL: 9,       // Column K (MoM % ECU Change)
  CA_FORECAST_COL: 14,      // Column P
  CA_FORECAST_VAR_COL: 16   // Column R
};

// ============================================
// HELPER FUNCTIONS - TEST MODE
// ============================================

/**
 * Get the target spreadsheet ID based on TEST_MODE
 * @return {string} Spreadsheet ID
 */
function getTargetSpreadsheetId() {
  if (CONFIG.TEST_MODE) {
    Logger.log('🧪 TEST MODE: Using test spreadsheet');
    return CONFIG.TEST_TARGET_SHEET_ID;
  }
  return CONFIG.TARGET_SHEET_ID;
}

/**
 * Get the target document ID based on TEST_MODE
 * @return {string} Document ID
 */
function getTargetDocumentId() {
  if (CONFIG.TEST_MODE) {
    Logger.log('🧪 TEST MODE: Using test document');
    return CONFIG.TEST_TARGET_DOCUMENT_ID;
  }
  return CONFIG.TARGET_DOCUMENT_ID;
}

// ============================================
// MAIN EXECUTION FUNCTION
// ============================================

/**
 * Main combined function to run weekly
 * Executes both spreadsheet and document operations
 */
function weeklyForecastUpdateCombined() {
  const startTime = new Date();
  let executionStatus = 'Running';
  let spreadsheetTabName = '';
  let totalAccounts = 0;
  let emailsSent = 0;
  let documentUpdated = false;
  let errorMessage = '';

  try {
    Logger.log('========================================');
    Logger.log('Starting combined weekly forecast update...');
    Logger.log('========================================');

    if (CONFIG.TEST_MODE) {
      Logger.log('⚠️ TEST MODE ENABLED - All @mentions will use ' + CONFIG.TEST_EMAIL);
    }

    // STEP 1: Read Global sheet data ONCE (shared by both operations)
    Logger.log('\n[1/5] Reading Global sheet data...');
    const globalData = readGlobalSheetData();
    Logger.log(`✓ Read ${globalData.growers.length} growers and ${globalData.shrinkers.length} shrinkers`);

    // STEP 2: Read CA mapping (for spreadsheet operations)
    Logger.log('\n[2/5] Reading CA-Lookup...');
    const caMapping = readCALookup();
    Logger.log(`✓ Found ${Object.keys(caMapping).length} CA mappings`);

    // STEP 3: SPREADSHEET OPERATIONS (Critical - must succeed)
    Logger.log('\n[3/5] Executing spreadsheet operations...');
    const spreadsheetResult = executeSpreadsheetOperations(globalData, caMapping);
    spreadsheetTabName = spreadsheetResult.tabName;
    totalAccounts = spreadsheetResult.totalAccounts;
    emailsSent = spreadsheetResult.emailsSent;
    Logger.log(`✓ Spreadsheet operations complete: Tab "${spreadsheetTabName}", ${emailsSent} emails sent`);

    // STEP 4: DOCUMENT OPERATIONS (Supplementary - can fail gracefully)
    Logger.log('\n[4/5] Executing document operations...');
    try {
      const documentResult = executeDocumentOperations(globalData, caMapping);
      documentUpdated = documentResult.success;
      Logger.log(`✓ Document operations complete: ${documentUpdated ? 'Success' : 'Failed'}`);
    } catch (docError) {
      Logger.log(`⚠️ Document operations failed (non-critical): ${docError.message}`);
      Logger.log('Continuing with spreadsheet success...');
      documentUpdated = false;
    }

    executionStatus = 'Success';
    Logger.log('\n========================================');
    Logger.log('✓ Combined update completed successfully!');
    Logger.log('========================================');

    showAlert('Success!',
      `Weekly forecast completed:\n\n` +
      `• Spreadsheet tab: ${spreadsheetTabName}\n` +
      `• Emails sent: ${emailsSent}\n` +
      `• Document updated: ${documentUpdated ? 'Yes' : 'No'}`);

  } catch (error) {
    executionStatus = 'Failed';
    errorMessage = error.message;
    Logger.log(`\n❌ Error in combined update: ${error.message}`);
    Logger.log(error.stack);
    showAlert('Error',
      `An error occurred: ${error.message}\n\nCheck the script logs for details.`);
    throw error;
  } finally {
    // STEP 5: Log execution
    Logger.log('\n[5/5] Logging execution...');
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    logExecution({
      timestamp: startTime,
      tabName: spreadsheetTabName,
      status: executionStatus,
      totalAccounts: totalAccounts,
      emailsSent: emailsSent,
      documentUpdated: documentUpdated,
      duration: duration,
      testMode: CONFIG.TEST_MODE,
      error: errorMessage
    });
    Logger.log('✓ Execution logged');
  }
}

// ============================================
// SPREADSHEET OPERATIONS
// ============================================

/**
 * Executes all spreadsheet operations
 * @param {Object} globalData - Data from Global sheet
 * @param {Object} caMapping - CA lookup mapping
 * @returns {Object} Result with tabName, totalAccounts, emailsSent
 */
function executeSpreadsheetOperations(globalData, caMapping) {
  // 1. Create new dated tab
  const newTab = createDatedTab();
  const tabName = newTab.getName();
  Logger.log(`  → Created tab: ${tabName}`);

  // 2. Write combined data with formatting
  const writeResult = writeCombinedData(newTab, globalData, caMapping);
  Logger.log(`  → Wrote ${writeResult.rowCount} rows of data`);

  // 3. Add @mention notes to cells
  const caAccounts = addNotesToCells(newTab, writeResult, caMapping);
  Logger.log(`  → Added notes for ${Object.keys(caAccounts).length} CAs`);

  // 4. Send email notifications
  sendCANotifications(newTab, caAccounts, caMapping);
  const emailsSent = Object.keys(caAccounts).length;
  const totalAccounts = Object.values(caAccounts).reduce((sum, accounts) => sum + accounts.length, 0);
  Logger.log(`  → Sent ${emailsSent} notification emails`);

  return {
    tabName: tabName,
    totalAccounts: totalAccounts,
    emailsSent: emailsSent
  };
}

/**
 * Creates a new tab with date name, appending version if needed
 * @returns {Sheet} The newly created sheet
 */
function createDatedTab() {
  const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let tabName = dateStr;
  let version = 2;

  while (targetSheet.getSheetByName(tabName) !== null) {
    tabName = `${dateStr}_v${version}`;
    version++;
  }

  const newTab = targetSheet.insertSheet(tabName);
  targetSheet.setActiveSheet(newTab);
  targetSheet.moveActiveSheet(1);

  return newTab;
}

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

  // 12 base columns (B-M) + 3 new columns (CA Forecast, Variance, Var%) + CA + Notes = 17
  const baseColumnCount = 12;
  const maxColumns = 17;
  const defaultBackground = '#ffffff';
  const defaultFontColor = '#000000';
  const defaultNumberFormat = '@';
  const currencyFormat = '"$"#,##0.0,"K"';
  const percentFormat = '0%';
  const defaultFontWeight = 'normal';

  // Add Growers section
  combinedData.push(['BIGGEST 10 GROWERS MoM']);
  combinedBackgrounds.push([defaultBackground]);
  combinedFontColors.push([defaultFontColor]);
  combinedNumberFormats.push([defaultNumberFormat]);
  combinedFontWeights.push(['bold']);

  // Build growers header: base columns + 3 new columns + CA + Notes
  const growersBaseHeader = data.growers[0].slice(0, baseColumnCount);
  const growersHeader = [...growersBaseHeader, 'CA Forecast', 'CA Forecast Variance', 'CA Forecast Var %', 'Customer Architect', 'Notes'];
  combinedData.push(growersHeader);
  // Use the header background and font color from source for consistency across all columns
  const growersHeaderColor = data.growersFormatting.backgrounds[0][0] || defaultBackground;
  const growersHeaderFontColor = data.growersFormatting.fontColors[0][0] || defaultFontColor;
  const growersHeaderBg = [...data.growersFormatting.backgrounds[0].slice(0, baseColumnCount), growersHeaderColor, growersHeaderColor, growersHeaderColor, growersHeaderColor, growersHeaderColor];
  const growersHeaderFc = [...data.growersFormatting.fontColors[0].slice(0, baseColumnCount), growersHeaderFontColor, growersHeaderFontColor, growersHeaderFontColor, growersHeaderFontColor, growersHeaderFontColor];
  const growersHeaderNf = [...data.growersFormatting.numberFormats[0].slice(0, baseColumnCount), currencyFormat, currencyFormat, percentFormat, defaultNumberFormat, defaultNumberFormat];
  const growersHeaderFw = [...data.growersFormatting.fontWeights[0].slice(0, baseColumnCount), 'bold', 'bold', 'bold', 'bold', 'bold'];
  combinedBackgrounds.push(growersHeaderBg);
  combinedFontColors.push(growersHeaderFc);
  combinedNumberFormats.push(growersHeaderNf);
  combinedFontWeights.push(growersHeaderFw);

  for (let i = 1; i < data.growers.length; i++) {
    const row = data.growers[i];
    const accountId = row[CONFIG.ACCOUNT_ID_COL]?.toString().trim();
    const caName = caMapping[accountId]?.name || '';
    const caForecast = row[CONFIG.CA_FORECAST_COL];
    const caForecastVar = row[CONFIG.CA_FORECAST_VAR_COL];
    const caForecastVarPct = calculateVariancePercent(caForecastVar, caForecast);

    const baseData = row.slice(0, baseColumnCount);
    combinedData.push([...baseData, caForecast, caForecastVar, caForecastVarPct, caName, '']);
    const rowBg = [...data.growersFormatting.backgrounds[i].slice(0, baseColumnCount), defaultBackground, defaultBackground, defaultBackground, defaultBackground, defaultBackground];
    const rowFc = [...data.growersFormatting.fontColors[i].slice(0, baseColumnCount), defaultFontColor, defaultFontColor, defaultFontColor, defaultFontColor, defaultFontColor];
    const rowNf = [...data.growersFormatting.numberFormats[i].slice(0, baseColumnCount), currencyFormat, currencyFormat, percentFormat, defaultNumberFormat, defaultNumberFormat];
    const rowFw = [...data.growersFormatting.fontWeights[i].slice(0, baseColumnCount), defaultFontWeight, defaultFontWeight, defaultFontWeight, defaultFontWeight, defaultFontWeight];
    combinedBackgrounds.push(rowBg);
    combinedFontColors.push(rowFc);
    combinedNumberFormats.push(rowNf);
    combinedFontWeights.push(rowFw);
  }

  // Blank row separator
  combinedData.push([]);
  combinedBackgrounds.push([]);
  combinedFontColors.push([]);
  combinedNumberFormats.push([]);
  combinedFontWeights.push([]);

  // Add Shrinkers section
  combinedData.push(['BIGGEST 10 SHRINKERS MoM']);
  combinedBackgrounds.push([defaultBackground]);
  combinedFontColors.push([defaultFontColor]);
  combinedNumberFormats.push([defaultNumberFormat]);
  combinedFontWeights.push(['bold']);

  // Build shrinkers header: base columns + 3 new columns + CA + Notes
  const shrinkersBaseHeader = data.shrinkers[0].slice(0, baseColumnCount);
  const shrinkersHeader = [...shrinkersBaseHeader, 'CA Forecast', 'CA Forecast Variance', 'CA Forecast Var %', 'Customer Architect', 'Notes'];
  combinedData.push(shrinkersHeader);
  // Use the header background and font color from source for consistency across all columns
  const shrinkersHeaderColor = data.shrinkersFormatting.backgrounds[0][0] || defaultBackground;
  const shrinkersHeaderFontColor = data.shrinkersFormatting.fontColors[0][0] || defaultFontColor;
  const shrinkersHeaderBg = [...data.shrinkersFormatting.backgrounds[0].slice(0, baseColumnCount), shrinkersHeaderColor, shrinkersHeaderColor, shrinkersHeaderColor, shrinkersHeaderColor, shrinkersHeaderColor];
  const shrinkersHeaderFc = [...data.shrinkersFormatting.fontColors[0].slice(0, baseColumnCount), shrinkersHeaderFontColor, shrinkersHeaderFontColor, shrinkersHeaderFontColor, shrinkersHeaderFontColor, shrinkersHeaderFontColor];
  const shrinkersHeaderNf = [...data.shrinkersFormatting.numberFormats[0].slice(0, baseColumnCount), currencyFormat, currencyFormat, percentFormat, defaultNumberFormat, defaultNumberFormat];
  const shrinkersHeaderFw = [...data.shrinkersFormatting.fontWeights[0].slice(0, baseColumnCount), 'bold', 'bold', 'bold', 'bold', 'bold'];
  combinedBackgrounds.push(shrinkersHeaderBg);
  combinedFontColors.push(shrinkersHeaderFc);
  combinedNumberFormats.push(shrinkersHeaderNf);
  combinedFontWeights.push(shrinkersHeaderFw);

  for (let i = 1; i < data.shrinkers.length; i++) {
    const row = data.shrinkers[i];
    const accountId = row[CONFIG.ACCOUNT_ID_COL]?.toString().trim();
    const caName = caMapping[accountId]?.name || '';
    const caForecast = row[CONFIG.CA_FORECAST_COL];
    const caForecastVar = row[CONFIG.CA_FORECAST_VAR_COL];
    const caForecastVarPct = calculateVariancePercent(caForecastVar, caForecast);

    const baseData = row.slice(0, baseColumnCount);
    combinedData.push([...baseData, caForecast, caForecastVar, caForecastVarPct, caName, '']);
    const rowBg = [...data.shrinkersFormatting.backgrounds[i].slice(0, baseColumnCount), defaultBackground, defaultBackground, defaultBackground, defaultBackground, defaultBackground];
    const rowFc = [...data.shrinkersFormatting.fontColors[i].slice(0, baseColumnCount), defaultFontColor, defaultFontColor, defaultFontColor, defaultFontColor, defaultFontColor];
    const rowNf = [...data.shrinkersFormatting.numberFormats[i].slice(0, baseColumnCount), currencyFormat, currencyFormat, percentFormat, defaultNumberFormat, defaultNumberFormat];
    const rowFw = [...data.shrinkersFormatting.fontWeights[i].slice(0, baseColumnCount), defaultFontWeight, defaultFontWeight, defaultFontWeight, defaultFontWeight, defaultFontWeight];
    combinedBackgrounds.push(rowBg);
    combinedFontColors.push(rowFc);
    combinedNumberFormats.push(rowNf);
    combinedFontWeights.push(rowFw);
  }

  // Pad rows
  for (let i = 0; i < combinedData.length; i++) {
    while (combinedData[i].length < maxColumns) {
      combinedData[i].push('');
      combinedBackgrounds[i].push(defaultBackground);
      combinedFontColors[i].push(defaultFontColor);
      combinedNumberFormats[i].push(defaultNumberFormat);
      combinedFontWeights[i].push(defaultFontWeight);
    }
  }

  // Write to sheet
  const dataRange = sheet.getRange(1, 1, combinedData.length, maxColumns);
  dataRange.setValues(combinedData);
  dataRange.setBackgrounds(combinedBackgrounds);
  dataRange.setFontColors(combinedFontColors);
  dataRange.setNumberFormats(combinedNumberFormats);
  dataRange.setFontWeights(combinedFontWeights);
  sheet.setFrozenRows(2);

  return {
    rowCount: combinedData.length,
    growersDataStartRow: 3,
    growersDataEndRow: 13,
    shrinkersDataStartRow: 16,
    shrinkersDataEndRow: 26,
    notesColumn: maxColumns
  };
}

/**
 * Adds notes to cells and collects CA account assignments
 * @param {Sheet} sheet - The target sheet
 * @param {Object} writeResult - Information about written data
 * @param {Object} caMapping - CA lookup mapping
 * @returns {Object} Mapping of CA email to their accounts
 */
function addNotesToCells(sheet, writeResult, caMapping) {
  const caAccounts = {};

  for (let i = writeResult.growersDataStartRow; i <= writeResult.growersDataEndRow; i++) {
    const accountInfo = addNoteForRow(sheet, i, writeResult.notesColumn, 'Growers', caMapping);
    if (accountInfo) {
      if (!caAccounts[accountInfo.email]) {
        caAccounts[accountInfo.email] = [];
      }
      caAccounts[accountInfo.email].push(accountInfo);
    }
  }

  for (let i = writeResult.shrinkersDataStartRow; i <= writeResult.shrinkersDataEndRow; i++) {
    const accountInfo = addNoteForRow(sheet, i, writeResult.notesColumn, 'Shrinkers', caMapping);
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
 * @param {Object} caMapping - CA lookup mapping
 * @returns {Object|null} Account info or null if skipped
 */
function addNoteForRow(sheet, row, col, section, caMapping) {
  const accountIdCol = CONFIG.ACCOUNT_ID_COL + 1;
  const accountNameCol = CONFIG.ACCOUNT_NAME_COL + 1;
  const areaCol = CONFIG.AREA_COL + 1;
  const wowPercentCol = CONFIG.WOW_PERCENT_COL + 1;
  const momPercentCol = CONFIG.MOM_PERCENT_COL + 1;
  const caForecastVarPctCol = 15;  // Column O: CA Forecast Var %
  const caCol = col - 1;

  const area = sheet.getRange(row, areaCol).getValue()?.toString().trim();

  if (!area || !area.includes('AMER')) {
    return null;
  }

  const accountId = sheet.getRange(row, accountIdCol).getValue()?.toString().trim();
  const accountName = cleanCompanyName(sheet.getRange(row, accountNameCol).getValue());
  const caName = sheet.getRange(row, caCol).getValue()?.toString().trim();
  const wowPercent = sheet.getRange(row, wowPercentCol).getValue();
  const momPercent = sheet.getRange(row, momPercentCol).getValue();
  const caForecastVarPct = sheet.getRange(row, caForecastVarPctCol).getValue();

  if (caName) {
    // Prioritize email from CA-Lookup, fall back to name-to-email conversion
    let email;
    if (CONFIG.TEST_MODE) {
      email = CONFIG.TEST_EMAIL;
    } else {
      const lookupEmail = caMapping[accountId]?.email;
      email = lookupEmail || convertNameToEmail(caName);
    }
    const noteText = CONFIG.TEST_MODE
      ? `[TEST MODE] @${email}\nPlease add notes for ${accountName}\n\nActual CA: ${caName}`
      : `@${email}\nPlease add notes for ${accountName}`;

    sheet.getRange(row, col).setNote(noteText);

    return {
      accountId: accountId,
      email: email,
      caName: caName,
      accountName: accountName,
      row: row,
      section: section,
      area: area,
      wowPercent: wowPercent,
      momPercent: momPercent,
      caForecastVarPct: caForecastVarPct
    };
  }

  return null;
}

/**
 * Sends email notifications to CAs with their assigned accounts
 * @param {Sheet} sheet - The target sheet
 * @param {Object} caAccounts - Mapping of CA email to their accounts
 * @param {Object} caMapping - CA lookup mapping with manager info
 */
function sendCANotifications(sheet, caAccounts, caMapping) {
  const webAppUrl = 'https://ela.st/weekly-consumption';
  const tabName = sheet.getName();

  for (const [email, accounts] of Object.entries(caAccounts)) {
    if (accounts.length === 0) continue;

    const caName = accounts[0].caName;
    const growers = accounts.filter(a => a.section === 'Growers');
    const shrinkers = accounts.filter(a => a.section === 'Shrinkers');
    const ccList = buildCCList(accounts, caMapping);

    let emailBody = CONFIG.TEST_MODE
      ? `[TEST MODE - This email would normally go to ${caName}]\n\n`
      : '';

    emailBody += `Hi ${caName.split(' ')[0]},\n\n`;
    emailBody += `You have ${accounts.length} account${accounts.length > 1 ? 's' : ''} assigned in this week's forecast that need notes added.\n\n`;
    emailBody += `Please review and add notes using the web app:\n${webAppUrl}\n\n`;
    emailBody += `In your notes, please provide:\n`;
    emailBody += `  • A summary of your recent activities with the account\n`;
    emailBody += `  • Why the account is in the grower/shrinker category\n`;
    emailBody += `  • If CA Forecast Variance is above 10%, explain the variance\n\n`;

    if (growers.length > 0) {
      emailBody += `📈 GROWERS (${growers.length}):\n`;
      growers.forEach(acc => {
        const varPctDisplay = formatCAForecastVarPct(acc.caForecastVarPct);
        emailBody += `  • ${acc.accountName} (WoW: ${formatPercent(acc.wowPercent)}, MoM: ${formatPercent(acc.momPercent)}, CA Forecast Var: ${varPctDisplay})\n`;
      });
      emailBody += '\n';
    }

    if (shrinkers.length > 0) {
      emailBody += `📉 SHRINKERS (${shrinkers.length}):\n`;
      shrinkers.forEach(acc => {
        const varPctDisplay = formatCAForecastVarPct(acc.caForecastVarPct);
        emailBody += `  • ${acc.accountName} (WoW: ${formatPercent(acc.wowPercent)}, MoM: ${formatPercent(acc.momPercent)}, CA Forecast Var: ${varPctDisplay})\n`;
      });
      emailBody += '\n';
    }

    emailBody += `Tab: ${tabName}\n\nThank you!`;

    const subject = CONFIG.TEST_MODE
      ? `[TEST] Weekly Forecast - Action Required (${accounts.length} account${accounts.length > 1 ? 's' : ''})`
      : `Weekly Forecast - Action Required (${accounts.length} account${accounts.length > 1 ? 's' : ''})`;

    try {
      const emailParams = {
        to: email,
        subject: subject,
        body: emailBody
      };

      if (ccList.length > 0) {
        emailParams.cc = CONFIG.TEST_MODE ? CONFIG.TEST_EMAIL : ccList.join(',');
        if (CONFIG.TEST_MODE) {
          emailBody = `[CC would include: ${ccList.join(', ')}]\n\n` + emailBody;
          emailParams.body = emailBody;
        }
      }

      MailApp.sendEmail(emailParams);
    } catch (error) {
      Logger.log(`Failed to send email to ${email}: ${error.message}`);
    }
  }
}

/**
 * Builds CC list using CA manager info from lookup
 * @param {Array} accounts - Array of account objects
 * @param {Object} caMapping - CA lookup mapping with manager info
 * @returns {Array} List of email addresses to CC
 */
function buildCCList(accounts, caMapping) {
  const ccSet = new Set();
  ccSet.add(CONFIG.SR_DIRECTOR_EMAIL);

  // Add CA managers from lookup data
  accounts.forEach(account => {
    if (account.accountId) {
      const caInfo = caMapping[account.accountId];
      if (caInfo && caInfo.managerEmail) {
        ccSet.add(caInfo.managerEmail);
      }
    }
  });

  return Array.from(ccSet);
}

// ============================================
// DOCUMENT OPERATIONS
// ============================================

/**
 * Executes all document operations
 * @param {Object} globalData - Data from Global sheet
 * @param {Object} caMapping - CA lookup mapping
 * @returns {Object} Result with success status
 */
function executeDocumentOperations(globalData, caMapping) {
  // Extract companies from global data (includes CA and WoW%)
  const growers = extractCompaniesFromGlobalData(globalData.growers, caMapping);
  const shrinkers = extractCompaniesFromGlobalData(globalData.shrinkers, caMapping);

  Logger.log(`  → Extracted ${growers.length} growers and ${shrinkers.length} shrinkers for document`);

  // Group by area (preserves full company objects with CA and WoW%)
  const growersByArea = groupCompaniesByArea(growers);
  const shrinkersByArea = groupCompaniesByArea(shrinkers);

  // Update document
  updateDocument(growersByArea, shrinkersByArea);

  return { success: true };
}

/**
 * Extract companies from global data format
 * @param {Array} dataRows - Rows from globalData
 * @param {Object} caMapping - CA lookup mapping
 * @returns {Array} Array of company objects {name, area, ca, wowPercent, momPercent}
 */
function extractCompaniesFromGlobalData(dataRows, caMapping) {
  const companies = [];

  // Skip header row (index 0), process data rows
  for (let i = 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    const accountId = row[CONFIG.ACCOUNT_ID_COL]?.toString().trim();
    const companyNameRaw = row[CONFIG.ACCOUNT_NAME_COL];
    const areaCode = row[CONFIG.AREA_COL] ? row[CONFIG.AREA_COL].toString().trim() : '';
    const wowPercent = row[CONFIG.WOW_PERCENT_COL];
    const momPercent = row[CONFIG.MOM_PERCENT_COL];

    const companyName = cleanCompanyName(companyNameRaw);
    const caName = caMapping[accountId]?.name || 'No CA';

    if (companyName && areaCode) {
      companies.push({
        name: companyName,
        area: areaCode,
        ca: caName,
        wowPercent: wowPercent,
        momPercent: momPercent
      });
    }
  }

  return companies;
}

/**
 * Group companies by their area code using the mapping
 * @param {Array} companies - Array of company objects
 * @return {Object} Object with section names as keys and arrays of company objects as values
 */
function groupCompaniesByArea(companies) {
  const grouped = {};

  for (const company of companies) {
    const sectionName = CONFIG.AREA_SECTION_MAP[company.area];

    if (!sectionName) {
      Logger.log(`  ⚠️ No mapping found for area code: ${company.area}`);
      continue;
    }

    if (!grouped[sectionName]) {
      grouped[sectionName] = [];
    }

    // Push the full company object (includes name, ca, wowPercent)
    grouped[sectionName].push(company);
  }

  return grouped;
}

/**
 * Update the document with growers and shrinkers
 * @param {Object} growersByArea - Growers grouped by section
 * @param {Object} shrinkersByArea - Shrinkers grouped by section
 */
function updateDocument(growersByArea, shrinkersByArea) {
  const doc = DocumentApp.openById(getTargetDocumentId());
  const tabs = doc.getTabs();
  let targetTab = null;

  for (let tab of tabs) {
    if (tab.getTitle() === 'Weekly notes') {
      targetTab = tab;
      break;
    }
  }

  if (!targetTab && tabs.length > 0) {
    targetTab = tabs[0];
  }

  if (!targetTab) {
    throw new Error('No suitable tab found in document');
  }

  const body = targetTab.asDocumentTab().getBody();
  const allSections = new Set([
    ...Object.keys(growersByArea),
    ...Object.keys(shrinkersByArea)
  ]);

  for (const sectionName of allSections) {
    const growers = growersByArea[sectionName] || [];
    const shrinkers = shrinkersByArea[sectionName] || [];
    updateSection(body, sectionName, growers, shrinkers);
  }

  Logger.log(`  → Updated ${allSections.size} sections in document`);
}

/**
 * Find and update a specific section in the document
 * @param {GoogleAppsScript.Document.Body} body - Document body
 * @param {string} sectionName - Name of the section to update
 * @param {Array} growers - Array of grower company names
 * @param {Array} shrinkers - Array of shrinker company names
 */
function updateSection(body, sectionName, growers, shrinkers) {
  try {
    const numElements = body.getNumChildren();
    let sectionFound = false;

    for (let i = 0; i < numElements; i++) {
      const element = body.getChild(i);
      const elementType = element.getType();

      if (elementType === DocumentApp.ElementType.TABLE) {
        const table = element.asTable();
        const numRows = table.getNumRows();

        for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
          const row = table.getRow(rowIndex);
          const numCells = row.getNumCells();

          if (numCells > 0) {
            const firstCell = row.getCell(0);
            const cellText = firstCell.getText().trim().replace(/\n/g, ' ');

            if (cellText.includes(sectionName)) {
              sectionFound = true;

              if (numCells > 1) {
                const contentCell = row.getCell(1);
                updateCellContent(contentCell, sectionName, growers, shrinkers);
              }

              break;
            }
          }
        }

        if (sectionFound) break;
      }

      if (elementType === DocumentApp.ElementType.PARAGRAPH) {
        const text = element.asParagraph().getText().trim();

        if (text.includes(sectionName)) {
          sectionFound = true;
          updateParagraphContent(body, i, numElements, sectionName, growers, shrinkers);
          break;
        }
      }
    }
  } catch (error) {
    Logger.log(`  ⚠️ Error updating section "${sectionName}": ${error.message}`);
  }
}

/**
 * Update content within a table cell
 * Format: **CompanyName** (CA Name) : WoW ±%, MoM ±%:
 * @param {GoogleAppsScript.Document.TableCell} cell - The cell to update
 * @param {string} sectionName - Section name for logging
 * @param {Array} growers - Grower company objects {name, ca, wowPercent, momPercent}
 * @param {Array} shrinkers - Shrinker company objects {name, ca, wowPercent, momPercent}
 */
function updateCellContent(cell, sectionName, growers, shrinkers) {
  // Clear cell
  while (cell.getNumChildren() > 0) {
    cell.removeChild(cell.getChild(0));
  }

  // Add growers with formatting: • **CompanyName** (CA) : WoW ±%, MoM ±%:
  if (growers.length > 0) {
    for (const grower of growers) {
      const companyText = grower.name;
      const caText = ` (${grower.ca}) : WoW ${formatPercent(grower.wowPercent)}, MoM ${formatPercent(grower.momPercent)}:`;
      const fullText = `• ${companyText}${caText}`;

      const para = cell.appendParagraph(fullText);
      // Make company name bold (after "• " which is 2 characters)
      const startBold = 2;
      const endBold = startBold + companyText.length - 1;
      para.editAsText().setBold(startBold, endBold, true);
    }
  }

  // Add shrinkers with formatting: • **CompanyName** (CA) : WoW ±%, MoM ±%:
  if (shrinkers.length > 0) {
    for (const shrinker of shrinkers) {
      const companyText = shrinker.name;
      const caText = ` (${shrinker.ca}) : WoW ${formatPercent(shrinker.wowPercent)}, MoM ${formatPercent(shrinker.momPercent)}:`;
      const fullText = `• ${companyText}${caText}`;

      const para = cell.appendParagraph(fullText);
      // Make company name bold (after "• " which is 2 characters)
      const startBold = 2;
      const endBold = startBold + companyText.length - 1;
      para.editAsText().setBold(startBold, endBold, true);
    }
  }
}

/**
 * Update content in paragraphs following a section header
 * Format: **CompanyName** (CA Name) : WoW ±%, MoM ±%:
 * @param {GoogleAppsScript.Document.Body} body - Document body
 * @param {number} startIndex - Index where section was found
 * @param {number} numElements - Total number of elements
 * @param {string} sectionName - Section name for logging
 * @param {Array} growers - Grower company objects {name, ca, wowPercent, momPercent}
 * @param {Array} shrinkers - Shrinker company objects {name, ca, wowPercent, momPercent}
 */
function updateParagraphContent(body, startIndex, numElements, sectionName, growers, shrinkers) {
  // For paragraph-based sections, we'll replace or insert formatted text
  // This is a fallback for non-table sections

  let currentIndex = startIndex + 1;

  // Add growers
  if (growers.length > 0) {
    for (const grower of growers) {
      const companyText = grower.name;
      const caText = ` (${grower.ca}) : WoW ${formatPercent(grower.wowPercent)}, MoM ${formatPercent(grower.momPercent)}:`;
      const fullText = `• ${companyText}${caText}`;

      // Check if we can reuse an existing paragraph or need to insert
      if (currentIndex < numElements) {
        const element = body.getChild(currentIndex);
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const para = element.asParagraph();
          para.setText(fullText);
          // Make company name bold (after "• ")
          const startBold = 2;
          const endBold = startBold + companyText.length - 1;
          para.editAsText().setBold(startBold, endBold, true);
          currentIndex++;
        }
      } else {
        const para = body.insertParagraph(currentIndex, fullText);
        // Make company name bold (after "• ")
        const startBold = 2;
        const endBold = startBold + companyText.length - 1;
        para.editAsText().setBold(startBold, endBold, true);
        currentIndex++;
      }
    }
  }

  // Add shrinkers
  if (shrinkers.length > 0) {
    for (const shrinker of shrinkers) {
      const companyText = shrinker.name;
      const caText = ` (${shrinker.ca}) : WoW ${formatPercent(shrinker.wowPercent)}, MoM ${formatPercent(shrinker.momPercent)}:`;
      const fullText = `• ${companyText}${caText}`;

      if (currentIndex < numElements) {
        const element = body.getChild(currentIndex);
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const para = element.asParagraph();
          para.setText(fullText);
          // Make company name bold (after "• ")
          const startBold = 2;
          const endBold = startBold + companyText.length - 1;
          para.editAsText().setBold(startBold, endBold, true);
          currentIndex++;
        }
      } else {
        const para = body.insertParagraph(currentIndex, fullText);
        // Make company name bold (after "• ")
        const startBold = 2;
        const endBold = startBold + companyText.length - 1;
        para.editAsText().setBold(startBold, endBold, true);
        currentIndex++;
      }
    }
  }
}

// ============================================
// SHARED DATA READING FUNCTIONS
// ============================================

/**
 * Reads CA-Lookup tab and returns mapping object
 * @returns {Object} Mapping of account_id to CA name
 */
function readCALookup() {
  const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
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

  // Read growers table
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

  // Read shrinkers table
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
// SHARED HELPER FUNCTIONS
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
  const firstName = parts[0];
  const lastName = parts[parts.length - 1];
  return `${firstName}.${lastName}${CONFIG.EMAIL_DOMAIN}`;
}

/**
 * Safely shows an alert dialog
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showAlert(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log(`Alert [${title}]: ${message}`);
  }
}

/**
 * Formats CA Forecast Variance percentage for display in emails
 * @param {number|string} value - The variance percentage value (as decimal, e.g., 0.10 for 10%)
 * @returns {string} Formatted percentage string with sign
 */
function formatCAForecastVarPct(value) {
  if (value === null || value === undefined || value === '' || value === 'N/A') {
    return 'N/A';
  }

  const numValue = typeof value === 'number' ? value : parseFloat(value);

  if (isNaN(numValue)) {
    return 'N/A';
  }

  const percentValue = numValue * 100;
  const sign = percentValue >= 0 ? '+' : '';
  return `${sign}${percentValue.toFixed(0)}%`;
}

/**
 * Calculates variance percentage (Variance / Forecast * 100)
 * @param {number|string} variance - The variance value
 * @param {number|string} forecast - The forecast value
 * @returns {number|string} Percentage value or 'N/A' if cannot calculate
 */
function calculateVariancePercent(variance, forecast) {
  if (forecast === null || forecast === undefined || forecast === '' || forecast === 0) {
    return 'N/A';
  }

  const numVariance = typeof variance === 'number' ? variance : parseFloat(variance);
  const numForecast = typeof forecast === 'number' ? forecast : parseFloat(forecast);

  if (isNaN(numVariance) || isNaN(numForecast) || numForecast === 0) {
    return 'N/A';
  }

  return numVariance / numForecast;
}

/**
 * Formats a percentage value for display
 * @param {number|string} value - The percentage value
 * @returns {string} Formatted percentage string
 */
function formatPercent(value) {
  if (value === null || value === undefined || value === '') {
    return 'N/A';
  }

  if (typeof value === 'string' && value.includes('%')) {
    return value;
  }

  const numValue = typeof value === 'number' ? value : parseFloat(value);

  if (isNaN(numValue)) {
    return 'N/A';
  }

  const percentValue = numValue * 100;
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
    const targetSheet = SpreadsheetApp.openById(getTargetSpreadsheetId());
    let logTab = targetSheet.getSheetByName(CONFIG.EXECUTION_LOG_TAB_NAME);

    if (!logTab) {
      logTab = targetSheet.insertSheet(CONFIG.EXECUTION_LOG_TAB_NAME);

      const headers = [
        'Timestamp',
        'Tab Name',
        'Status',
        'Total Accounts',
        'Emails Sent',
        'Document Updated',
        'Duration (sec)',
        'Test Mode',
        'Error Message'
      ];
      logTab.getRange(1, 1, 1, headers.length).setValues([headers]);
      logTab.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      logTab.setFrozenRows(1);
    }

    const row = [
      logData.timestamp,
      logData.tabName || 'N/A',
      logData.status,
      logData.totalAccounts || 0,
      logData.emailsSent || 0,
      logData.documentUpdated ? 'Yes' : 'No',
      logData.duration ? logData.duration.toFixed(2) : 0,
      logData.testMode ? 'Yes' : 'No',
      logData.error || ''
    ];

    logTab.appendRow(row);

    const lastRow = logTab.getLastRow();
    logTab.getRange(lastRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    const statusCell = logTab.getRange(lastRow, 3);
    if (logData.status === 'Success') {
      statusCell.setBackground('#d9ead3');
    } else if (logData.status === 'Failed') {
      statusCell.setBackground('#f4cccc');
    }

  } catch (error) {
    Logger.log(`Failed to log execution: ${error.message}`);
  }
}

// ============================================
// MENU & TEST FUNCTIONS
// ============================================

/**
 * Creates custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Weekly Forecast')
    .addItem('Run Combined Update', 'weeklyForecastUpdateCombined')
    .addSeparator()
    .addItem('CA Notes Interface', 'showNotesInterface')
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
