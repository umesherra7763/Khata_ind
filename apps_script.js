// ============================================================
//  Khata Index — Google Apps Script Web App
//  Deploy this as a Web App (Execute as: Me, Access: Anyone)
//  and paste the resulting URL into index.html
// ============================================================

// ⚠️ PASTE YOUR GOOGLE SHEET ID HERE (from the URL of your sheet)
// e.g. https://docs.google.com/spreadsheets/d/THIS_PART_HERE/edit
const SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const SHEET_NAME = 'Entries'; // Name of the tab/sheet

// ------------------------------------------------------------
//  GET handler — returns all entries as JSON
// ------------------------------------------------------------
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getAll') {
      return jsonResponse(getAllEntries());
    }
    return jsonResponse({ error: 'Unknown action' }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ------------------------------------------------------------
//  POST handler — add, update, or delete an entry
// ------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'add') {
      return jsonResponse(addEntry(data));
    } else if (action === 'update') {
      return jsonResponse(updateEntry(data));
    } else if (action === 'delete') {
      return jsonResponse(deleteEntry(data.id));
    }
    return jsonResponse({ error: 'Unknown action' }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ------------------------------------------------------------
//  CRUD helpers
// ------------------------------------------------------------

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function getAllEntries() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // only header row

  return data.slice(1).map(row => ({
    id:         String(row[0]),
    letter:     String(row[1]),
    name:       String(row[2]),
    pageNumber: Number(row[3])
  }));
}

function addEntry(data) {
  const sheet = getSheet();
  const id = String(Date.now());
  const letter = data.name.trim().charAt(0).toUpperCase();
  sheet.appendRow([id, letter, data.name.trim(), parseInt(data.pageNumber, 10)]);
  return { success: true, id, letter };
}

function updateEntry(data) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      const letter = data.name.trim().charAt(0).toUpperCase();
      sheet.getRange(i + 1, 1, 1, 4).setValues([
        [String(data.id), letter, data.name.trim(), parseInt(data.pageNumber, 10)]
      ]);
      return { success: true };
    }
  }
  return { error: 'Entry not found' };
}

function deleteEntry(id) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Entry not found' };
}

// ------------------------------------------------------------
//  Utility — return JSON with CORS headers
// ------------------------------------------------------------
function jsonResponse(obj, statusCode) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
