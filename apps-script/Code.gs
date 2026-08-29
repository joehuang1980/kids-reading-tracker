/**
 * Kids Reading Tracker — Google Apps Script Web App
 * Stage 6: 在 Google Sheet 建立名為「Kids Reading Records」的試算表
 *          新增工作表「Reading_Log」，欄位如下：
 *          A: Record_ID  B: ISBN  C: Book_Title  D: Reading_Date
 *          E: Reader     F: Created_At  G: Thumbnail
 *
 * Stage 7: 部署步驟
 *  1. 開啟 https://script.google.com，新增專案
 *  2. 貼上此程式碼，存檔
 *  3. 「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *     - 以以下身分執行：「我」
 *     - 誰可以存取：「所有人」
 *  4. 複製 Web App URL，填入 js/config.js 的 APPS_SCRIPT_URL
 *
 * Stage 10: Excel 匯出 — 在 Google Sheet 選「檔案」→「下載」→「Microsoft Excel (.xlsx)」
 */

var SHEET_NAME = 'Reading_Log';

function getSheet() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Record_ID', 'ISBN', 'Book_Title', 'Reading_Date', 'Reader', 'Created_At', 'Thumbnail']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Stage 8: Receive reading record from frontend
function doPost(e) {
  try {
    var sheet = getSheet();
    var data  = JSON.parse(e.postData.contents);

    // Validation
    if (!data.title || !data.readingDate || !data.reader) {
      return _json({ success: false, message: 'Missing required fields' });
    }

    // Stage 9: Duplicate check — same ISBN + reader + date within 5 minutes
    if (data.isbn) {
      var lr = sheet.getLastRow();
      if (lr > 1) {
        var checkRows = Math.min(20, lr - 1);
        var rows = sheet.getRange(Math.max(2, lr - checkRows + 1), 1, checkRows, 7).getValues();
        var now  = new Date();
        for (var i = 0; i < rows.length; i++) {
          var r       = rows[i];
          var rowTime = r[5] instanceof Date ? r[5] : new Date(r[5]);
          // r[3] may be a Date object; normalise to yyyy-MM-dd string for comparison
          var rowDate = r[3] instanceof Date
            ? Utilities.formatDate(r[3], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[3]);
          if (String(r[1]) === String(data.isbn) && rowDate === data.readingDate && r[4] === data.reader) {
            if ((now - rowTime) < 300000) {  // within 5 minutes
              return _json({ success: false, duplicate: true, message: '重複紀錄' });
            }
          }
        }
      }
    }

    var recordId = 'R-' +
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd') + '-' +
      Math.random().toString(36).substr(2, 4).toUpperCase();

    sheet.appendRow([
      recordId,
      data.isbn      || '',
      data.title,
      data.readingDate,
      data.reader,
      new Date(),
      data.thumbnail || '',
    ]);

    return _json({ success: true, recordId: recordId });
  } catch (err) {
    return _json({ success: false, message: String(err) });
  }
}

// Stage 11–12: Return all records for history / dashboard pages
function doGet(e) {
  try {
    var sheet = getSheet();
    var lr    = sheet.getLastRow();
    if (lr < 2) return _json({ records: [] });

    var rows    = sheet.getRange(2, 1, lr - 1, 7).getValues();
    var records = rows
      .filter(function(r) { return r[2]; })  // skip rows without title
      .map(function(r) {
        return {
          id:          r[0],
          isbn:        String(r[1] || ''),  // Sheets stores ISBN as number
          title:       r[2],
          readingDate: r[3] instanceof Date
            ? Utilities.formatDate(r[3], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[3]),
          reader:      r[4],
          createdAt:   r[5] instanceof Date ? r[5].toISOString() : String(r[5]),
          thumbnail:   r[6] || '',
        };
      });

    return _json({ records: records });
  } catch (err) {
    return _json({ success: false, message: String(err) });
  }
}

// One-time helper: run manually in Apps Script editor to clear all data rows
function clearAllRecords() {
  var sheet = getSheet();
  var lr = sheet.getLastRow();
  if (lr > 1) sheet.deleteRows(2, lr - 1);
  Logger.log('Cleared ' + (lr - 1) + ' records.');
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
