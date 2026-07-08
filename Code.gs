// ==========================================
// ICON T-Shirt Vote Backend
// Google Apps Script
// ==========================================

const SHEET_NAME = "Votes";

// Allowed design IDs
const ALLOWED_CHOICES = [
  "design_1",
  "design_2",
  "design_3",
  "design_4",
  "design_5"
];

// ------------------------------------------
// Get or create sheet
// ------------------------------------------
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp",
      "Name",
      "Choice"
    ]);
  }

  return sheet;
}

// ------------------------------------------
// Submit Vote
// ------------------------------------------
function doPost(e) {

  try {

    if (!e.postData || !e.postData.contents) {
      return jsonResponse_(false, "No POST data received.");
    }

    let body;

    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonResponse_(false, "Invalid JSON.");
    }

    const name = (body.name || "").trim().substring(0, 100);
    const choice = (body.choice || "").trim();

    if (!name) {
      return jsonResponse_(false, "Name is required.");
    }

    if (ALLOWED_CHOICES.indexOf(choice) === -1) {
      return jsonResponse_(false, "Invalid design selected.");
    }

    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();

    // Prevent duplicate voting
    for (let i = 1; i < data.length; i++) {

      const existingName = String(data[i][1]).trim().toLowerCase();

      if (existingName === name.toLowerCase()) {
        return jsonResponse_(false, "You have already voted.");
      }
    }

    sheet.appendRow([
      new Date(),
      name,
      choice
    ]);

    return jsonResponse_(true, "Vote submitted successfully.");

  } catch (err) {

    return jsonResponse_(false, err.toString());

  }

}

// ------------------------------------------
// Get Results
// ------------------------------------------
function doGet(e) {

  try {

    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();

    const counts = {};
    let voterCount = 0;

    ALLOWED_CHOICES.forEach(function(choice) {
      counts[choice] = 0;
    });

    for (let i = 1; i < data.length; i++) {

      const choice = String(data[i][2]);

      if (counts.hasOwnProperty(choice)) {
        counts[choice]++;
        voterCount++;
      }

    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        voterCount: voterCount,
        counts: counts
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return jsonResponse_(false, err.toString());

  }

}

// ------------------------------------------
// JSON Response Helper
// ------------------------------------------
function jsonResponse_(ok, message) {

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: ok,
      message: message
    }))
    .setMimeType(ContentService.MimeType.JSON);

}
