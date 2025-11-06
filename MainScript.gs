/**
 * ===================================================================
 * FINAL ROSTER SYNC SCRIPT - Runs Daily
 * ===================================================================
 *
 * This script runs on a daily trigger. It reads the 'Engineers'
 * roster, finds today's date, and updates the existing Notion
 * database rows with the correct 'Person' tags.
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

// List of text in roster that means the person is NOT working
const LEAVE_TERMS = ["Week Off", "EL", "NH", "COMP OFF", "#REF!"];
// --- END CONFIGURATION ---

/**
 * Creates a master menu in the Google Sheet.
 * This one menu can run both the daily sync AND the setup tools.
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

  // Set status in control panel
  const statusCell = controlSheet.getRange("B3");
  statusCell.setValue("Syncing...");

  try {
    // 1. Load Mappers
    const shiftPageIdMap = loadShiftPageIdMap(controlSheet);
    if (shiftPageIdMap.size === 0) {
      throw new Error(
        'Shift Page ID map is empty! Run the "Find Shift Page IDs" helper first.'
      );
    }

    const engineerIdMap = loadEngineerIdMap(
      ss.getSheetByName(MAPPER_SHEET_NAME)
    );
    if (engineerIdMap.size === 0) {
      throw new Error(
        'Engineer ID map is empty! Fill out the "ID Mapper" tab.'
      );
    }
    Logger.log(`Loaded ${engineerIdMap.size} Engineer User IDs.`);

    // 2. Get Roster Data
    const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!rosterSheet) {
      throw new Error(`Sheet "${ROSTER_SHEET_NAME}" not found.`);
    }

    // Use Utilities.formatDate, NOT SpreadsheetApp.Utilities.formatDate
    const todayStr = Utilities.formatDate(new Date(), TIMEZONE, DATE_FORMAT);
    Logger.log(`Today's date key: ${todayStr}`);

    const dateColumn = findDateColumn(rosterSheet, todayStr);
    if (!dateColumn) {
      throw new Error(
        `Could not find today's date "${todayStr}" in roster header (Row ${ROSTER_HEADER_ROW}). Check format.`
      );
    }
    Logger.log(`Found date column: ${dateColumn}`);

    // 3. Process the Roster
    const shiftData = processRoster(rosterSheet, dateColumn, engineerIdMap);
    Logger.log("Finished processing roster.");

    // 4. Update Notion
    const { NOTION_API_KEY } = getConfig();
    let updatedCount = 0;

    // Loop through our list of shifts (from the control panel map)
    for (let [shiftName, pageId] of shiftPageIdMap.entries()) {
      // Find the data for this shift
      const dataForShift = shiftData[shiftName] || { L1: [], L2: [] }; // Default to empty

      // Call the API to update the page
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

    // 5. Report Success
    const successMsg = `Sync complete for ${todayStr}. Updated ${updatedCount} shift rows in Notion.`;
    statusCell.setValue(successMsg);
    Logger.log(successMsg);
  } catch (e) {
    // 6. Report Error
    Logger.log(e);
    statusCell.setValue(`Error: ${e.message}`);
    SpreadsheetApp.getUi().alert(e.message);
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
      map.set(shiftName, pageId);
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
function processRoster(sheet, dateColumn, engineerIdMap) {
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
    const shiftName = row[dateColumn - 1]; // The value in the date column, e.g., "Shift 1"

    // 1. Skip if no shift or engineer
    if (!shiftName || !engineerName || !designation) {
      continue;
    }

    // 2. Skip if on leave
    if (
      typeof shiftName !== "string" ||
      LEAVE_TERMS.includes(shiftName.toUpperCase())
    ) {
      continue;
    }

    // 3. Find Notion User ID
    const notionUserId = engineerIdMap.get(engineerName);
    if (!notionUserId) {
      Logger.log(
        `Warning: Skipping "${engineerName}" - not found in ID Mapper.`
      );
      continue;
    }

    // 4. Prepare Notion Person object
    const notionPerson = { id: notionUserId };

    // 5. Add them to the 'shifts' object
    if (!shifts[shiftName]) {
      // If "Shift 1" isn't an entry yet, create it
      shifts[shiftName] = { L1: [], L2: [] };
    }

    if (designation === "L1") {
      shifts[shiftName].L1.push(notionPerson);
    } else if (designation === "L2") {
      shifts[shiftName].L2.push(notionPerson);
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

  // This payload assumes your Notion properties are named "L1" and "L2"
  const payload = {
    properties: {
      L1: {
        type: "people",
        people: l1People, // e.g., [{id: "user-1"}, {id: "user-2"}]
      },
      L2: {
        type: "people",
        people: l2People, // e.g., [{id: "user-3"}]
      },
    },
  };

  const options = {
    method: "patch", // 'patch' means UPDATE, not create
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
