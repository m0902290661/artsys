/**
 * QR Code 流量統計與跳轉 API 程式庫 (直接以 URL 為 Key)
 * 工作表結構:
 * 1. QR_Analytics (Original URL, Total Views, Last Accessed, Note)
 * 2. Access_Logs  (Timestamp, Target URL, User Agent)
 */

const SHEET_NAME = 'QR_Analytics';
const LOG_SHEET_NAME = 'Access_Logs';

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let qrSheet = ss.getSheetByName(SHEET_NAME);
  if (!qrSheet) {
    qrSheet = ss.insertSheet(SHEET_NAME);
    qrSheet.appendRow(['Original URL', 'Total Views', 'Last Accessed', 'Note']);
    qrSheet.getRange("1:1").setFontWeight("bold").setBackground("#e0e0e0");
  }

  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(['Timestamp', 'Target URL', 'User Agent']);
    logSheet.getRange("1:1").setFontWeight("bold").setBackground("#e0e0e0");
  }
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action;
  let result = { success: false, message: 'Invalid action' };

  try {
    switch (action) {
      case 'logAndGetUrl':
        result = logAndGetUrl(params.targetUrl, params.ua || '', params.note || '');
        break;
      case 'getStats':
        result = getStats(params.targetUrl);
        break;
      case 'getAllStats':
        result = getAllStats();
        break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 紀錄網址點擊流量 (若網址不存在會自動建立)
 */
function logAndGetUrl(targetUrl, userAgent, note) {
  if (!targetUrl) return { success: false, message: 'Missing targetUrl parameter' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  let currentViews = 0;

  // 使用完整的 URL 直接進行比對
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetUrl).trim()) {
      targetRow = i + 1;
      currentViews = Number(data[i][1]) || 0;
      break;
    }
  }

  const now = new Date();

  if (targetRow !== -1) {
    // 已存在網址：次數 + 1，更新時間
    sheet.getRange(targetRow, 2).setValue(currentViews + 1);
    sheet.getRange(targetRow, 3).setValue(now);
    if (note) sheet.getRange(targetRow, 4).setValue(note);
  } else {
    // 首次存取網址：自動新增紀錄
    sheet.appendRow([targetUrl, 1, now, note || '']);
    currentViews = 0;
  }

  // 寫入詳細存取 Log
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (logSheet) {
    logSheet.appendRow([now, targetUrl, userAgent]);
  }

  return {
    success: true,
    targetUrl: targetUrl,
    totalViews: currentViews + 1
  };
}

/**
 * 查詢特定 URL 的統計
 */
function getStats(targetUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: 'Sheet not initialized' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetUrl).trim()) {
      return {
        success: true,
        data: {
          targetUrl: data[i][0],
          totalViews: data[i][1],
          lastAccessed: data[i][2],
          note: data[i][3]
        }
      };
    }
  }
  return { success: false, message: 'Target URL not found' };
}

/**
 * 取得所有流量統計
 */
function getAllStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: 'Sheet not initialized' };

  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      targetUrl: data[i][0],
      totalViews: data[i][1],
      lastAccessed: data[i][2],
      note: data[i][3]
    });
  }
  return { success: true, count: list.length, data: list };
}