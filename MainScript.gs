/**
 * ===================================================================
 * FINAL ROSTER SYNC SCRIPT - Runs Daily
 * ===================================================================
 *
 * This script runs on a daily trigger. It reads the 'Engineers'
 * roster, finds today's date, and updates the existing Notion
 * database rows with the correct 'Person' tags.
 *
 * * v2.2 - Added hash-checking to prevent re-syncing unchanged data.
 * It compares the current roster data for a day against the hash
 * stored from the last successful sync of that same day.
 */

// --- SCRIPT CONFIGURATION ---
const ROSTER_SHEET_NAME = "Engineers"; // The name of your roster tab
const ROSTER_HEADER_ROW = 2; // Row number with dates (e.g., "4-Nov")
const ROSTER_START_ROW = 3; // Row number where engineer list starts
const DATE_FORMAT = "d-MMM"; // Format of dates in header (e.g., "4-Nov")
const TIMEZONE = "Asia/Kolkata"; // Your local timezone

const ENGINEER_NAME_COL = 3; // Column C
const DESIGNATION_COL = 4; // Column D

const CONTROL_SHEET_NAME = "Sync Control";
const MAPPER_SHEET_NAME = "ID Mapper";
const LOG_SHEET_NAME = "Shift Engg Logs";
const HASH_SHEET_NAME = "Sync Hashes"; // New sheet to store hashes
// --- END CONFIGURATION ---

/**
 * Creates a master menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Roster Sync Menu"); // Master menu

  menu.addItem("Sync TODAY_S Roster", "runDailySync"); // Main action
  menu.addSeparator();

  // Sub-menu for helper/setup functions
  const setupSubMenu = ui.createMenu("Setup Tools (Run Once)");
  setupSubMenu.addItem("Find User IDs", "findUserIDs");
  setupSubMenu.addItem("Find Shift Page IDs", "findShiftPageIDs");
  menu.addSubMenu(setupSubMenu);

  menu.addSeparator();
  menu.addItem("Clear Debug Logs", "clearLogSheet");

  menu.addToUi();
}

/**
 * Main function. Runs daily via trigger or manually from menu.
 */
function runDailySync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = ss.getSheetByName(CONTROL_SHEET_NAME);

  if (!controlSheet) {
    SpreadsheetApp.getUi().alert(
      `Error: Sheet "${CONTROL_SHEET_NAME}" not found.`
    );
    return;
  }

  const statusCell = controlSheet.getRange("B3");
  statusCell.setValue("Syncing...");

  try {
    // 1. Clear old logs and prepare log sheet
    clearLogSheet();
    Logger.log("Log sheet cleared.");

    // 2. Load Mappers
    const shiftPageIdMap = loadShiftPageIdMap(controlSheet);
    if (shiftPageIdMap.size === 0)
      throw new Error("Shift Page ID map is empty!");

    const engineerIdMap = loadEngineerIdMap(
      ss.getSheetByName(MAPPER_SHEET_NAME)
    );
    if (engineerIdMap.size === 0) throw new Error("Engineer ID map is empty!");
    Logger.log(`Loaded ${engineerIdMap.size} Engineer User IDs.`);

    // 3. Get Roster Data
    const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!rosterSheet)
      throw new Error(`Sheet "${ROSTER_SHEET_NAME}" not found.`);

    const todayStr = Utilities.formatDate(new Date(), TIMEZONE, DATE_FORMAT);
    Logger.log(`Today's date key: ${todayStr}`);

    const dateColumn = findDateColumn(rosterSheet, todayStr);
    if (!dateColumn)
      throw new Error(
        `Could not find today's date "${todayStr}" in roster header (Row ${ROSTER_HEADER_ROW}).`
      );
    Logger.log(`Found date column: ${dateColumn}`);

    // 4. Process the Roster
    const shiftData = processRoster(
      rosterSheet,
      dateColumn,
      engineerIdMap,
      shiftPageIdMap
    );
    Logger.log("Finished processing roster.");

    // 5. *** NEW HASHING LOGIC ***
    const currentDataHash = generateDataHash(shiftData);
    const storedHash = getStoredHash(todayStr);

    if (currentDataHash === storedHash) {
      const skipMsg = `No changes detected for ${todayStr}. Sync skipped.`;
      Logger.log(skipMsg);
      statusCell.setValue(skipMsg);
      // Write to log sheet
      logToSheet(["N/A", "N/A", "N/A", skipMsg]);
      return; // Stop the script
    }
    Logger.log(
      `New hash ${currentDataHash} does not match stored hash ${storedHash}. Proceeding with sync.`
    );
    // *** END HASHING LOGIC ***

    // 6. Update Notion (Only runs if hashes don't match)
    const { NOTION_API_KEY } = getConfig();
    let updatedCount = 0;

    for (let [shiftName, pageId] of shiftPageIdMap.entries()) {
      const dataForShift = shiftData[shiftName] || { L1: [], L2: [] };
      updateNotionPage(
        NOTION_API_KEY,
        pageId,
        dataForShift.L1,
        dataForShift.L2
      );
      Logger.log(
        `Successfully updated page for "${shiftName}" (ID: ${pageId})`
      );
      updatedCount++;
    }

    // 7. Report Success and SAVE the new hash
    setStoredHash(todayStr, currentDataHash); // Save the new hash
    const successMsg = `Sync complete for ${todayStr}. Updated ${updatedCount} shift rows in Notion. New hash saved.`;
    statusCell.setValue(successMsg);
    Logger.log(successMsg);
  } catch (e) {
    // 8. Report Error
    Logger.log(e);
    statusCell.setValue(`Error: ${e.message}`);
    SpreadsheetApp.getUi().alert(e.message);
  }
}

// ==================================================================
// NEW HASHING FUNCTIONS
// ==================================================================

/**
 * Creates/finds the Sync Hashes sheet.
 */
function getHashSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HASH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HASH_SHEET_NAME);
    sheet.appendRow(["Date String", "Data Hash"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Generates a stable MD5 hash from the shift data object.
 */
function generateDataHash(shiftData) {
  // Sort the object keys to ensure the string is always in the same order
  const sortedKeys = Object.keys(shiftData).sort();
  const sortedData = {};
  for (const key of sortedKeys) {
    sortedData[key] = shiftData[key];
  }
  // Convert the stable, sorted object to a JSON string
  const dataString = JSON.stringify(sortedData);
  // Compute the hash
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dataString)
    .map((byte) => (byte & 0xff).toString(16).padStart(2, "0"))
    .join("");
}

/**
 * Finds the stored hash for a given date.
 */
function getStoredHash(dateStr) {
  const sheet = getHashSheet();
  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  for (const row of data) {
    if (row[0] === dateStr) {
      return row[1]; // Found it
    }
  }
  return null; // Not found
}

/**
 * Saves a new hash for a given date, overwriting the old one.
 */
function setStoredHash(dateStr, hash) {
  const sheet = getHashSheet();
  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  let found = false;
  // Try to update existing row
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === dateStr) {
      sheet.getRange(i + 2, 2).setValue(hash); // i+2 because data is 0-indexed and starts at row 2
      found = true;
      break;
    }
  }
  // If not found, add a new row
  if (!found) {
    sheet.appendRow([dateStr, hash]);
  }
}

// ==================================================================
// EXISTING FUNCTIONS (No changes below this line)
// ==================================================================

/**
 * Creates/clears the log sheet and adds headers.
 */
function clearLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
  }
  sheet.clear(); // Clear all old data
  const headers = [
    "Timestamp",
    "Engineer Name",
    "Designation (from Sheet)",
    "Shift (from Sheet)",
    "Script Action",
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Appends a log message to the log sheet.
 */
function logToSheet(logMessage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOG_SHEET_NAME);
    const timestamp = Utilities.formatDate(
      new Date(),
      TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );
    sheet.appendRow([timestamp, ...logMessage]);
  } catch (e) {
    Logger.log(`Failed to write to log sheet: ${e.message}`);
  }
}

/**
 * Gets API Key and DB ID from Script Properties
 */
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const NOTION_API_KEY = scriptProperties.getProperty("NOTION_API_KEY");
  const NOTION_DATABASE_ID = scriptProperties.getProperty("NOTION_DATABASE_ID");

  if (!NOTION_API_KEY || !NOTION_DATABASE_ID) {
    throw new Error("API Key or Database ID not set in Script Properties.");
  }
  return { NOTION_API_KEY, NOTION_DATABASE_ID };
}

/**
 * Reads the 'Sync Control' sheet and returns a Map of
 * { "Shift 1" -> "page-id-abc...", "Shift 2" -> "page-id-xyz..." }
 */
function loadShiftPageIdMap(controlSheet) {
  const map = new Map();
  // Start from row 3 to skip headers
  const data = controlSheet
    .getRange(`D3:E${controlSheet.getLastRow()}`)
    .getValues();
  for (const row of data) {
    const shiftName = row[0]; // e.g., "Shift 1"
    const pageId = row[1]; // e.g., "abc-123..."
    if (shiftName && pageId) {
      map.set(shiftName.trim(), pageId.trim()); // Trim spaces to be safe
    }
  }
  return map;
}

/**
 * Reads the 'ID Mapper' sheet and returns a Map of
 * { "Adish" -> "user-id-123...", "Archie" -> "user-id-456..." }
 */
function loadEngineerIdMap(mapperSheet) {
  const map = new Map();
  if (!mapperSheet) {
    throw new Error(`Sheet "${MAPPER_SHEET_NAME}" not found.`);
  }
  // Start from row 2 to skip header
  const data = mapperSheet
    .getRange(`A2:B${mapperSheet.getLastRow()}`)
    .getValues();
  for (const row of data) {
    const engineerName = row[0]; // e.g., "Adish"
    const userId = row[1]; // e.g., "user-id-123..."
    if (engineerName && userId) {
      map.set(engineerName, userId);
    }
  }
  return map;
}

/**
 * Scans the header row for a specific date string (e.g., "4-Nov")
 * and returns the column number (1-indexed).
 */
function findDateColumn(sheet, dateStr) {
  const headerRow = sheet
    .getRange(ROSTER_HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    // Check if cellValue is a Date object and format it
    if (cellValue instanceof Date) {
      if (Utilities.formatDate(cellValue, TIMEZONE, DATE_FORMAT) === dateStr) {
        return i + 1; // +1 because columns are 1-indexed
      }
    }
    // Check if cellValue is a string and matches
    else if (typeof cellValue === "string" && cellValue.trim() === dateStr) {
      return i + 1;
    }
  }
  return null; // Not found
}

/**
 * Loops through the roster for the given date column and sorts
 * engineers into an object grouped by shift and designation.
 */
function processRoster(sheet, dateColumn, engineerIdMap, shiftPageIdMap) {
  const numRows = sheet.getLastRow() - ROSTER_START_ROW + 1;
  if (numRows <= 0) {
    return {}; // No data rows
  }
  const range = sheet.getRange(
    ROSTER_START_ROW,
    1,
    numRows,
    sheet.getLastColumn()
  );
  const values = range.getValues();

  const shifts = {}; // This will hold our final data: { "Shift 1": { L1: [], L2: [] } }

  for (const row of values) {
    const engineerName = row[ENGINEER_NAME_COL - 1]; // -1 for 0-indexing
    const designation = row[DESIGNATION_COL - 1]
      .toString()
      .toUpperCase()
      .replace("*", ""); // "L1", "L2"
    let shiftName = row[dateColumn - 1]; // The value in the date column, e.g., "Shift 1" or "OH"

    // 1. Clean up shiftName
    if (typeof shiftName === "string") {
      shiftName = shiftName.trim();
    } else {
      shiftName = String(shiftName); // Convert numbers/other types to string
    }

    // 2. Skip if no engineer name
    if (!engineerName) {
      // Don't log, this is likely a blank formatting row
      continue;
    }

    // 3. VALIDATION: Is this a shift we sync?
    if (!shiftPageIdMap.has(shiftName)) {
      logToSheet([
        engineerName,
        designation,
        shiftName,
        "Skipped (Shift not mapped in Sync Control)",
      ]);
      continue;
    }

    // 4. Find Notion User ID
    const notionUserId = engineerIdMap.get(engineerName);
    if (!notionUserId) {
      Logger.log(
        `Warning: Skipping "${engineerName}" - not found in ID Mapper.`
      );
      logToSheet([
        engineerName,
        designation,
        shiftName,
        "Skipped (No User ID in Mapper)",
      ]);
      continue;
    }

    // 5. Prepare Notion Person object
    const notionPerson = { id: notionUserId };

    // 6. Add them to the 'shifts' object
    if (!shifts[shiftName]) {
      shifts[shiftName] = { L1: [], L2: [] };
    }

    if (designation === "L1") {
      shifts[shiftName].L1.push(notionPerson);
      logToSheet([
        engineerName,
        designation,
        shiftName,
        "Added to L1 bucket for " + shiftName,
      ]);
    } else if (designation === "L2") {
      shifts[shiftName].L2.push(notionPerson);
      logToSheet([
        engineerName,
        designation,
        shiftName,
        "Added to L2 bucket for " + shiftName,
      ]);
    } else {
      logToSheet([
        engineerName,
        designation,
        shiftName,
        "Skipped (Designation not L1 or L2)",
      ]);
    }
  }
  return shifts;
}

/**
 * Sends the final API call to Notion to update a single page (row)
 * with the new lists of L1 and L2 people.
 */
function updateNotionPage(apiKey, pageId, l1People, l2People) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = {
    properties: {
      L1: {
        type: "people",
        people: l1People,
      },
      L2: {
        type: "people",
        people: l2People,
      },
    },
  };

  const options = {
    method: "patch",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Notion-Version": "2022-06-28",
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log(
      `Failed to update page ${pageId}. Response: ${response.getContentText()}`
    );
    throw new Error(`Notion API Error (PATCH): ${response.getContentText()}`);
  }
}
