/**
 * ===================================================================
 * HELPER SCRIPT - Run ONCE
 * ===================================================================
 *
 * This file contains the one-time setup functions.
 * They are called from the 'Roster Sync Menu' > 'Setup Tools' menu
 * defined in Code.gs.
 */

/**
 * Fetches all users in your Notion workspace.
 * Creates a new sheet named 'User IDs' with the results.
 */
function findUserIDs() {
  try {
    // This function only needs the API Key
    const scriptProperties = PropertiesService.getScriptProperties();
    const NOTION_API_KEY = scriptProperties.getProperty("NOTION_API_KEY");
    if (!NOTION_API_KEY) {
      throw new Error(
        "NOTION_API_KEY is not set in Script Properties. Please check Project Settings (⚙️)."
      );
    }

    const url = "https://api.notion.com/v1/users";
    const options = {
      method: "get",
      headers: {
        Authorization: `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": "2022-06-28",
      },
      muteHttpExceptions: true,
    };

    let allUsers = [];
    let nextCursor = undefined;

    // Loop to handle pagination (if you have > 100 users)
    do {
      let fullUrl = url;
      if (nextCursor) {
        fullUrl += `?start_cursor=${nextCursor}`;
      }
      const response = UrlFetchApp.fetch(fullUrl, options);
      const data = JSON.parse(response.getContentText());

      if (response.getResponseCode() !== 200) {
        throw new Error(
          `Notion API Error (Users): ${response.getContentText()}`
        );
      }

      allUsers = allUsers.concat(data.results);
      nextCursor = data.next_cursor;
    } while (nextCursor);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("User IDs");
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet("User IDs");
    }

    sheet.appendRow(["Name", "Email", "User ID (COPY THIS)"]);

    allUsers.forEach((user) => {
      if (user.type === "person" && user.person) {
        sheet.appendRow([user.name, user.person.email, user.id]);
      }
    });

    sheet.autoResizeColumns(1, 3);
    SpreadsheetApp.getUi().alert(
      'Finished! A tab "User IDs" has been created/updated with all users from your workspace.'
    );
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

/**
 * Fetches all pages in your database (e.g., "Shift 1").
 * Prints the results onto the 'Sync Control' sheet.
 */
function findShiftPageIDs() {
  try {
    // This function needs both keys, so it uses the main getConfig()
    const { NOTION_API_KEY, NOTION_DATABASE_ID } = getConfig();
    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;
    const options = {
      method: "post",
      headers: {
        Authorization: `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": "2022-06-28",
      },
      payload: JSON.stringify({}), // Empty payload queries all pages
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `Notion API Error (Database Query): ${response.getContentText()}`
      );
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONTROL_SHEET_NAME); // CONTROL_SHEET_NAME is from Code.gs
    if (!sheet) {
      sheet = ss.insertSheet(CONTROL_SHEET_NAME);
    }

    // Clear old data first
    sheet.getRange("D1:E100").clearContent(); // Clear a large range

    // Add headers
    sheet.getRange("D1").setValue("--- Shift Page IDs ---");
    sheet.getRange("D2").setValue("Shift Name (Title)");
    sheet.getRange("E2").setValue("Page ID");

    let row = 3;
    data.results.forEach((page) => {
      try {
        // This finds the 'Title' property automatically
        const titleProperty = Object.values(page.properties).find(
          (prop) => prop.type === "title"
        );
        const shiftName = titleProperty.title[0].plain_text;

        sheet.getRange(row, 4).setValue(shiftName); // Col D
        sheet.getRange(row, 5).setValue(page.id); // Col E
        row++;
      } catch (e) {
        // Ignore pages without a title or with incorrect property name
        Logger.log(
          `Skipping a page. Could not find title property. Error: ${e.message}`
        );
      }
    });

    sheet.autoResizeColumns(4, 5);
    SpreadsheetApp.getUi().alert(
      `Finished! Check the "${CONTROL_SHEET_NAME}" tab (column D & E) for your Page IDs.`
    );
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}
