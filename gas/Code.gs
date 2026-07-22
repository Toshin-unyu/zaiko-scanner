const SHEET_NAME = 'マスタ';
const HEADER = ['JANコード', '商品ページURL', '登録日時'];
const DOMAIN_SHEET_NAME = '許可ドメイン';
const DOMAIN_HEADER = ['ドメイン'];
const DEFAULT_ALLOWED_DOMAIN = 'solution.soloel.com';
const ADMIN_KEY_PROPERTY = 'ADMIN_KEY';

function getAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty(ADMIN_KEY_PROPERTY);
}

function timingSafeEqual_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function isValidAdminKey_(key) {
  const adminKey = getAdminKey_();
  return !!adminKey && typeof key === 'string' && timingSafeEqual_(key, adminKey);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function getDomainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DOMAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DOMAIN_SHEET_NAME);
    sheet.appendRow(DOMAIN_HEADER);
    sheet.appendRow([DEFAULT_ALLOWED_DOMAIN]);
  }
  return sheet;
}

function getAllowedDomains_() {
  const sheet = getDomainSheet_();
  const values = sheet.getDataRange().getValues();
  const domains = [];
  for (let i = 1; i < values.length; i++) {
    const domain = String(values[i][0]).trim();
    if (domain) domains.push(domain);
  }
  return domains;
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

function findDomainRow_(sheet, domain) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === domain.toLowerCase()) {
      return i + 1;
    }
  }
  return null;
}

function extractHostname_(url) {
  const match = String(url).match(/^https:\/\/([^\/:?#]+)(\/.*)?$/);
  return match ? match[1] : null;
}

function isAllowedUrl_(url) {
  const hostname = extractHostname_(url);
  if (!hostname) return false;
  const allowed = getAllowedDomains_();
  return allowed.some(d => d.toLowerCase() === hostname.toLowerCase());
}

function doGet(e) {
  if (e.parameter.action === 'domains') {
    return jsonResponse_({ success: true, domains: getAllowedDomains_() });
  }

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

  if (payload.action === 'addDomain') {
    return addDomain_(payload.domain, payload.adminKey);
  }
  if (payload.action === 'removeDomain') {
    return removeDomain_(payload.domain, payload.adminKey);
  }
  return registerProduct_(payload);
}

function addDomain_(domain, adminKey) {
  if (!isValidAdminKey_(adminKey)) {
    return jsonResponse_({ success: false, message: '管理キーが正しくありません' });
  }
  domain = String(domain || '').trim();
  if (!domain || /[\/:]/.test(domain)) {
    return jsonResponse_({ success: false, message: 'ドメイン名の形式が不正です（例: solution.soloel.com）' });
  }
  const sheet = getDomainSheet_();
  if (findDomainRow_(sheet, domain)) {
    return jsonResponse_({ success: false, message: 'そのドメインは既に登録されています' });
  }
  sheet.appendRow([domain]);
  return jsonResponse_({ success: true });
}

function removeDomain_(domain, adminKey) {
  if (!isValidAdminKey_(adminKey)) {
    return jsonResponse_({ success: false, message: '管理キーが正しくありません' });
  }
  domain = String(domain || '').trim();
  const sheet = getDomainSheet_();
  const rowIndex = findDomainRow_(sheet, domain);
  if (!rowIndex) {
    return jsonResponse_({ success: false, message: 'そのドメインは登録されていません' });
  }
  const remainingCount = getAllowedDomains_().length - 1;
  if (remainingCount < 1) {
    return jsonResponse_({ success: false, message: '最後の1件は削除できません（登録が全て不可になるため）' });
  }
  sheet.deleteRow(rowIndex);
  return jsonResponse_({ success: true });
}

function registerProduct_(payload) {
  const jan = payload.jan;
  const url = payload.url;
  if (!jan || !url) {
    return jsonResponse_({ success: false, message: 'jan と url が必要です' });
  }

  if (!isAllowedUrl_(url)) {
    return jsonResponse_({ success: false, message: '登録できるのは許可済みドメイン配下のhttps URLのみです' });
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
