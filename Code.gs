/**
 * Webinar Registration - Google Apps Script backend
 *
 * Handles two course groups: MPC and BIPC.
 * Each has its own registration sheet and YouTube link.
 *
 * Config sheet layout:
 *   A1: "MPC YouTube Link:"   B1: <mpc link>
 *   A2: "BIPC YouTube Link:"  B2: <bipc link>
 */

var SPREADSHEET_ID = "1vsap_kKyT5YoOd-dk77AFi_YJYIF2q4jPFA87KwABng";

var CONFIG_SHEET = "Config";
var MPC_SHEET    = "MPC Registrations";
var BIPC_SHEET   = "BIPC Registrations";

var MPC_LINK_CELL  = "B1";
var BIPC_LINK_CELL = "B2";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var course = String(data.course || "").toUpperCase();

    if (course !== "MPC" && course !== "BIPC") {
      return json({ status: "error", message: "Invalid course. Must be MPC or BIPC." });
    }

    var ss = getSpreadsheet();

    // Save registration to the correct sheet
    var sheetName = course === "MPC" ? MPC_SHEET : BIPC_SHEET;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["Timestamp", "Parent Name", "Parent Number", "Student Name", "Student Number"]);
    }

    var existing = sheet.getDataRange().getValues().slice(1);
    var isDuplicate = existing.some(function(row) {
      return String(row[2]).trim() === String(data.parentNumber || "").trim();
    });
    if (!isDuplicate) {
      sheet.appendRow([
        new Date(),
        data.parentName    || "",
        data.parentNumber  || "",
        data.studentName   || "",
        data.studentNumber || ""
      ]);
    }

    var link = getYouTubeLink(ss, course);
    return json({ status: "success", meetLink: link });

  } catch (err) {
    return json({ status: "error", message: String(err) });
  }
}

/** Handles prefetch GET requests — ?course=MPC or ?course=BIPC */
function doGet(e) {
  var course = (e && e.parameter && e.parameter.course)
    ? String(e.parameter.course).toUpperCase()
    : "MPC";
  var ss = getSpreadsheet();
  return json({ status: "ok", meetLink: getYouTubeLink(ss, course) });
}

function getYouTubeLink(ss, course) {
  var config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) {
    config = ss.insertSheet(CONFIG_SHEET);
    config.getRange("A1").setValue("MPC YouTube Link:");
    config.getRange(MPC_LINK_CELL).setValue("https://youtu.be/KmSYJBn6sMU");
    config.getRange("A2").setValue("BIPC YouTube Link:");
    config.getRange(BIPC_LINK_CELL).setValue("https://youtu.be/iITGn3-Ge0s");
  }
  var cell = course === "MPC" ? MPC_LINK_CELL : BIPC_LINK_CELL;
  return config.getRange(cell).getValue();
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
