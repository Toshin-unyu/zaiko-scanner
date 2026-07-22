const SHEET_NAME = 'マスタ';
const HEADER = ['JANコード', '商品ページURL', '登録日時'];
const ALLOWED_HOSTNAME = 'solution.soloel.com';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowByJan_(sheet, jan) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(jan)) {
      return { rowIndex: i + 1, url: values[i][1] };
    }
  }
  return null;
}

function doGet(e) {
  const jan = e.parameter.jan;
  if (!jan) {
    return jsonResponse_({ success: false, message: 'jan パラメータが必要です' });
  }
  const sheet = getSheet_();
  const found = findRowByJan_(sheet, jan);
  if (!found) {
    return jsonResponse_({ success: true, found: false });
  }
  return jsonResponse_({ success: true, found: true, url: found.url });
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ success: false, message: 'リクエストの形式が不正です' });
  }

  const jan = payload.jan;
  const url = payload.url;
  if (!jan || !url) {
    return jsonResponse_({ success: false, message: 'jan と url が必要です' });
  }

  const match = String(url).match(/^https:\/\/([^\/]+)(\/.*)?$/);
  if (!match || match[1] !== ALLOWED_HOSTNAME) {
    return jsonResponse_({ success: false, message: `登録できるのは https://${ALLOWED_HOSTNAME}/ 配下のURLのみです` });
  }

  const sheet = getSheet_();
  const existing = findRowByJan_(sheet, jan);
  const now = new Date();
  if (existing) {
    sheet.getRange(existing.rowIndex, 2).setValue(url);
    sheet.getRange(existing.rowIndex, 3).setValue(now);
    return jsonResponse_({ success: true, updated: true });
  }

  sheet.appendRow([jan, url, now]);
  return jsonResponse_({ success: true, updated: false });
}
