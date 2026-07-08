// ICON T-Shirt Vote — Google Apps Script backend (single-choice version)
// -----------------------------------------------------------------------
// Paste this into the Apps Script editor of a Google Sheet
// (Extensions > Apps Script), then deploy as a Web App.

var SHEET_NAME = "Votes";

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Name", "Choice"]);
  }
  return sheet;
}

// Handles vote submission: POST { name, choice: "design_5" }
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var name = (body.name || "Anonymous").toString().substring(0, 100);
    var choice = (body.choice || "").toString();

    var sheet = getSheet_();
    sheet.appendRow([new Date(), name, choice]);

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.toString() });
  }
}

// Handles results retrieval: GET request
// Returns vote counts per design + total voter count
function doGet(e) {
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var counts = {}; // designId -> count
    var voterCount = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var choice = row[2];
      if (!choice) continue;
      voterCount++;
      counts[choice] = (counts[choice] || 0) + 1;
    }

    return jsonResponse_({ ok: true, voterCount: voterCount, counts: counts });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.toString() });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
