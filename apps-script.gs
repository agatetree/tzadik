// Yahrtzeit Tracker - Google Apps Script
// Paste this entire file into: script.google.com -> New Project

// CHANGE THIS to your own secret password before deploying
var SECRET_KEY = 'Tishu@Box6';

// Email address that receives the password recovery email
var RECOVERY_EMAIL = 'agateestates@gmail.com';

// Sheet names
var SHEET_ADDED       = 'נוסף';
var SHEET_SUGGESTIONS = 'הצעות';
var SHEET_HIDDEN      = 'מוסתר';
var SHEET_MESSAGE     = 'הודעה';
var SHEET_FULL        = 'כל הרשימה';

var ADDED_HEADERS   = ['חודש','יום','עברי','לועזי','שם','בר','כינוי','פרטי','רבי'];
var SUGGEST_HEADERS = ['תאריך','חודש','יום','עברי','לועזי','שם','בר','כינוי','הערה'];
var HIDDEN_HEADERS  = ['מפתח'];
var FULL_HEADERS    = ['חודש','יום','עברי','לועזי','שם','בר','כינוי'];

// GET: return all added entries + hidden keys + message
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var addedSheet = ss.getSheetByName(SHEET_ADDED);
  var entries = [];
  if (addedSheet && addedSheet.getLastRow() >= 2) {
    var rows = addedSheet.getRange(2, 1, addedSheet.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var entry = {
        month:      String(r[0]).trim(),
        day:        String(r[1]).trim(),
        hebrewYear: String(r[2]).trim(),
        gregYear:   String(r[3]).trim(),
        name:       String(r[4]).trim(),
        fatherName: String(r[5]).trim(),
        title:      String(r[6]).trim(),
        private:    String(r[7]).trim(),
        rabbi:      String(r[8]).trim() === 'false' ? false : true
      };
      if (entry.month && entry.day && entry.name) {
        entries.push(entry);
      }
    }
  }

  var hiddenSheet = ss.getSheetByName(SHEET_HIDDEN);
  var hidden = [];
  if (hiddenSheet && hiddenSheet.getLastRow() >= 2) {
    var hRows = hiddenSheet.getRange(2, 1, hiddenSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < hRows.length; j++) {
      var k = String(hRows[j][0]).trim();
      if (k) hidden.push(k);
    }
  }

  var message = '';
  var msgSheet = ss.getSheetByName(SHEET_MESSAGE);
  if (msgSheet && msgSheet.getLastRow() >= 2) {
    message = String(msgSheet.getRange(2, 1).getValue()).trim();
  }

  return respond({ entries: entries, hidden: hidden, message: message });
}

// POST: handle all write actions
function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ error: 'Invalid JSON' });
  }

  // Validate password (login check, no side effects)
  if (data.action === 'validate') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    return respond({ success: true });
  }

  // Password recovery: send SECRET_KEY to RECOVERY_EMAIL
  if (data.action === 'recover') {
    try {
      MailApp.sendEmail({
        to: RECOVERY_EMAIL,
        subject: 'Yahrtzeit Tracker - Password Recovery',
        body: 'Your admin password is: ' + SECRET_KEY
      });
      return respond({ success: true });
    } catch (mailErr) {
      return respond({ error: 'mail_failed', detail: mailErr.toString() });
    }
  }

  // Add a new entry
  if (data.action === 'add') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var sheet = getOrCreateSheet(SHEET_ADDED, ADDED_HEADERS);
    sheet.appendRow([
      data.month      || '',
      data.day        || '',
      data.hebrewYear || '',
      data.gregYear   || '',
      data.name       || '',
      data.fatherName || '',
      data.title      || '',
      data.private    ? 'true' : '',
      data.rabbi === false ? 'false' : ''
    ]);
    return respond({ success: true });
  }

  // Update an existing sheet entry (replaces old row, keeps sheet clean)
  if (data.action === 'update') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var sheet = getOrCreateSheet(SHEET_ADDED, ADDED_HEADERS);
    var oldKey = String(data.oldKey || '').trim();
    if (oldKey && sheet.getLastRow() >= 2) {
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        var r = rows[i];
        var rKey = [
          String(r[0]).trim(), String(r[1]).trim(), String(r[4]).trim(),
          String(r[5]).trim(), String(r[2]).trim(), String(r[3]).trim(),
          String(r[6]).trim()
        ].join('|');
        if (String(r[8]).trim() === 'false') rKey = rKey + '|norabbi';
        if (rKey === oldKey) sheet.deleteRow(i + 2);
      }
    }
    sheet.appendRow([
      data.month      || '',
      data.day        || '',
      data.hebrewYear || '',
      data.gregYear   || '',
      data.name       || '',
      data.fatherName || '',
      data.title      || '',
      data.private    ? 'true' : '',
      data.rabbi === false ? 'false' : ''
    ]);
    return respond({ success: true });
  }

  // Suggest (public)
  if (data.action === 'suggest') {
    var sheet = getOrCreateSheet(SHEET_SUGGESTIONS, SUGGEST_HEADERS);
    sheet.appendRow([
      new Date(),
      data.month      || '',
      data.day        || '',
      data.hebrewYear || '',
      data.gregYear   || '',
      data.name       || '',
      data.fatherName || '',
      data.title      || '',
      data.note       || ''
    ]);
    return respond({ success: true });
  }

  // Hide an entry
  if (data.action === 'hide') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var sheet = getOrCreateSheet(SHEET_HIDDEN, HIDDEN_HEADERS);
    var key = String(data.entryKey || '').trim();
    if (key) {
      var existing = [];
      if (sheet.getLastRow() >= 2) {
        var eRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        for (var i = 0; i < eRows.length; i++) {
          existing.push(String(eRows[i][0]).trim());
        }
      }
      if (existing.indexOf(key) === -1) {
        sheet.appendRow([key]);
      }
    }
    return respond({ success: true });
  }

  // Unhide an entry
  if (data.action === 'unhide') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hiddenSheet = ss.getSheetByName(SHEET_HIDDEN);
    if (hiddenSheet && hiddenSheet.getLastRow() >= 2) {
      var target = String(data.entryKey || '').trim();
      var hRows = hiddenSheet.getRange(2, 1, hiddenSheet.getLastRow() - 1, 1).getValues();
      for (var i = hRows.length - 1; i >= 0; i--) {
        if (String(hRows[i][0]).trim() === target) {
          hiddenSheet.deleteRow(i + 2);
        }
      }
    }
    return respond({ success: true });
  }

  // Set message
  if (data.action === 'setMessage') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_MESSAGE);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_MESSAGE);
      sheet.appendRow([SHEET_MESSAGE]);
      sheet.getRange(1, 1).setFontWeight('bold').setBackground('#243b55').setFontColor('#ffffff');
    }
    if (sheet.getLastRow() < 2) sheet.appendRow(['']);
    sheet.getRange(2, 1).setValue(data.message || '');
    return respond({ success: true });
  }

  // Populate full list sheet
  if (data.action === 'populate') {
    if (data.key !== SECRET_KEY) return respond({ error: 'Unauthorized' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_FULL);
    if (!sheet) {
      var first = ss.getSheets()[0];
      var isUnused = first.getLastRow() <= 1
                  && first.getName() !== SHEET_ADDED
                  && first.getName() !== SHEET_SUGGESTIONS;
      if (isUnused) {
        first.setName(SHEET_FULL);
        sheet = first;
      } else {
        sheet = ss.insertSheet(SHEET_FULL);
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(1);
      }
    }
    sheet.clearContents();
    sheet.appendRow(FULL_HEADERS);
    sheet.getRange(1, 1, 1, FULL_HEADERS.length)
         .setFontWeight('bold').setBackground('#243b55').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    var entries = data.entries || [];
    if (entries.length > 0) {
      var rows = [];
      for (var i = 0; i < entries.length; i++) {
        var en = entries[i];
        rows.push([en.month, en.day, en.hebrewYear, en.gregYear, en.name, en.fatherName, en.title]);
      }
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }
    return respond({ success: true, count: entries.length });
  }

  return respond({ error: 'Unknown action' });
}

// Helpers
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#243b55')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
