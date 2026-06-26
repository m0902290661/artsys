const SPREADSHEET_ID = "1O2wUFvNN9mnIaI09kxaeyHI_UcDy77u5E1TTVlz82uA";
const MAIN_FOLDER_ID = "1yrORQBOk72ghb5vVPRhMs0gSOVy0HEPi";
const SESSION_TTL_HOURS = 8;
const USER_HEADERS = [
  "studentId", "name", "className", "club", "password", "sessionToken", "sessionExpiresAt"
];

function authorizeDriveAndSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);
  return "授權成功：" + ss.getName() + " / " + folder.getName();
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getParams(e) {
  if (e && e.postData && e.postData.contents) {
    var text = e.postData.contents;
    if (text.trim().charAt(0) === "{") {
      return JSON.parse(text);
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (action === "getSchedule") {
    return jsonOutput(getScheduleRows_(ss));
  }

  if (action === "getSubmissions") {
    var adminSession = requireSession_(ss, params, true);
    if (!adminSession.success) return jsonOutput(adminSession);
    return jsonOutput(getSubmissionRows_(ss));
  }

  if (action === "getMySubmissions") {
    var mySession = requireSession_(ss, params, false);
    if (!mySession.success) return jsonOutput(mySession);
    if (mySession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return jsonOutput(getSubmissionRows_(ss).filter(function(item) {
      return item.studentId.toString() === (params.studentId || "").toString();
    }));
  }

  if (action === "getUsers") {
    var usersSession = requireSession_(ss, params, true);
    if (!usersSession.success) return jsonOutput(usersSession);
    return jsonOutput(getSheetObjects_(ss, "users", USER_HEADERS).map(function(user) {
      delete user.password;
      delete user.sessionToken;
      delete user.sessionExpiresAt;
      return user;
    }));
  }

  if (action === "validateSession") {
    return jsonOutput(requireSession_(ss, params, false));
  }

  if (action === "authCheck") {
    return authCheck_();
  }

  return jsonOutput({ success: false, message: "未知的 GET action" });
}

function doPost(e) {
  var params = getParams(e);
  var action = params.action;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var userSheet = ss.getSheetByName("users");
  ensureUserSessionColumns_(userSheet);
  var userData = userSheet.getDataRange().getValues();

  if (action === "login") {
    return login_(userSheet, userData, params);
  }

  if (action === "logout") {
    return logout_(userSheet, userData, params);
  }

  if (action === "signup") {
    return signup_(userSheet, userData, params);
  }

  if (action === "uploadPDF") {
    var uploadSession = requireSession_(ss, params, false);
    if (!uploadSession.success) return jsonOutput(uploadSession);
    if (uploadSession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return uploadPdf_(ss, params);
  }

  if (action === "addSchedule") {
    var addSession = requireSession_(ss, params, true);
    if (!addSession.success) return jsonOutput(addSession);
    return addSchedule_(ss, params);
  }

  if (action === "deleteSchedule") {
    var deleteSession = requireSession_(ss, params, true);
    if (!deleteSession.success) return jsonOutput(deleteSession);
    return deleteSchedule_(ss, params);
  }

  if (action === "revokeSubmission") {
    var revokeSession = requireSession_(ss, params, false);
    if (!revokeSession.success) return jsonOutput(revokeSession);
    if (revokeSession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return revokeSubmission_(ss, params);
  }

  if (action === "clearSubmissions") {
    var clearSession = requireSession_(ss, params, true);
    if (!clearSession.success) return jsonOutput(clearSession);
    return clearSubmissions_(ss, params);
  }

  if (action === "addUser") {
    var addUserSession = requireSession_(ss, params, true);
    if (!addUserSession.success) return jsonOutput(addUserSession);
    return addUser_(userSheet, userData, params);
  }

  if (action === "updateUserPassword") {
    var passwordSession = requireSession_(ss, params, true);
    if (!passwordSession.success) return jsonOutput(passwordSession);
    return updateUserPassword_(userSheet, userData, params);
  }

  return jsonOutput({ success: false, message: "未知的 POST action" });
}

function login_(userSheet, userData, params) {
  var headerMap = getHeaderMap_(userSheet);
  var tokenCol = headerMap.sessionToken + 1;
  var expiresCol = headerMap.sessionExpiresAt + 1;
  var studentId = (params.studentId || "").toString().trim();
  var password = (params.password || "").toString();

  if (!studentId || !password) {
    return jsonOutput({ success: false, message: "請輸入學號與密碼。" });
  }

  for (var i = 1; i < userData.length; i++) {
    if (
      userData[i][0].toString() === studentId &&
      userData[i][4].toString() === password
    ) {
      var token = Utilities.getUuid() + "-" + new Date().getTime();
      var expiresAt = new Date(new Date().getTime() + SESSION_TTL_HOURS * 60 * 60 * 1000);
      userSheet.getRange(i + 1, tokenCol).setValue(token);
      userSheet.getRange(i + 1, expiresCol).setValue(expiresAt);
      SpreadsheetApp.flush();
      return jsonOutput({
        success: true,
        message: "登入成功",
        sessionToken: token,
        sessionExpiresAt: expiresAt,
        user: {
          studentId: userData[i][0],
          name: userData[i][1],
          className: userData[i][2],
          club: userData[i][3],
          sessionToken: token,
          sessionExpiresAt: expiresAt
        }
      });
    }
  }
  return jsonOutput({ success: false, message: "學號或密碼錯誤" });
}

function logout_(userSheet, userData, params) {
  var headerMap = getHeaderMap_(userSheet);
  var studentId = (params.studentId || params.adminId || "").toString().trim();
  var sessionToken = (params.sessionToken || "").toString();

  if (!studentId || !sessionToken) {
    return jsonOutput({ success: true, message: "已登出。" });
  }

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][headerMap.studentId].toString() !== studentId) continue;
    if ((userData[i][headerMap.sessionToken] || "").toString() === sessionToken) {
      userSheet.getRange(i + 1, headerMap.sessionToken + 1).setValue("");
      userSheet.getRange(i + 1, headerMap.sessionExpiresAt + 1).setValue("");
      SpreadsheetApp.flush();
    }
    return jsonOutput({ success: true, message: "已登出。" });
  }

  return jsonOutput({ success: true, message: "已登出。" });
}

function signup_(userSheet, userData, params) {
  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === params.studentId.toString()) {
      if (userData[i][4] !== "") {
        return jsonOutput({ success: false, message: "該學號已開通" });
      }
      userSheet.getRange(i + 1, 5).setValue(params.password);
      return jsonOutput({ success: true, message: "帳號開通成功！" });
    }
  }
  return jsonOutput({ success: false, message: "找不到此學號" });
}

function addUser_(userSheet, userData, params) {
  var studentId = (params.newStudentId || params.targetStudentId || "").toString().trim();
  var name = (params.name || "").toString().trim();
  var className = (params.className || "").toString().trim();
  var club = (params.club || "").toString().trim();
  var password = (params.password || "").toString();

  if (!studentId || !name || !className) {
    return jsonOutput({ success: false, message: "請填寫學號、姓名與班級。" });
  }

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === studentId) {
      return jsonOutput({ success: false, message: "此學號已存在，未新增。" });
    }
  }

  userSheet.appendRow([studentId, name, className, club, password, "", ""]);
  return jsonOutput({ success: true, message: "人員已新增。" });
}

function updateUserPassword_(userSheet, userData, params) {
  var studentId = (params.targetStudentId || "").toString().trim();
  var password = (params.password || "").toString();

  if (!studentId || !password) {
    return jsonOutput({ success: false, message: "缺少學號或新密碼。" });
  }

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === studentId) {
      userSheet.getRange(i + 1, 5).setValue(password);
      userSheet.getRange(i + 1, 6).setValue("");
      userSheet.getRange(i + 1, 7).setValue("");
      return jsonOutput({ success: true, message: "密碼已更新，該帳號需重新登入。" });
    }
  }

  return jsonOutput({ success: false, message: "找不到此學號。" });
}

function ensureUserSessionColumns_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), USER_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  for (var i = 0; i < USER_HEADERS.length; i++) {
    if (headers[i] !== USER_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(USER_HEADERS[i]);
    }
  }
}

function getHeaderMap_(sheet) {
  ensureUserSessionColumns_(sheet);
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), USER_HEADERS.length)).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) map[headers[i]] = i;
  }
  return map;
}

function requireSession_(ss, params, adminOnly) {
  var studentId = (params.studentId || params.adminId || "").toString();
  var sessionToken = (params.sessionToken || "").toString();

  if (!studentId || !sessionToken) {
    return { success: false, message: "登入已失效，請重新登入。" };
  }

  var userSheet = ss.getSheetByName("users");
  ensureUserSessionColumns_(userSheet);
  var headerMap = getHeaderMap_(userSheet);
  var data = userSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][headerMap.studentId].toString() !== studentId) continue;

    var storedToken = (data[i][headerMap.sessionToken] || "").toString();
    var expiresAt = data[i][headerMap.sessionExpiresAt];
    var expiresDate = expiresAt instanceof Date ? expiresAt : new Date(expiresAt);

    if (!storedToken || storedToken !== sessionToken) {
      return { success: false, message: "登入密鑰不一致，請重新登入。" };
    }
    if (!expiresAt || Number.isNaN(expiresDate.getTime()) || expiresDate.getTime() <= new Date().getTime()) {
      return { success: false, message: "登入已逾時，請重新登入。" };
    }

    var user = {
      studentId: data[i][headerMap.studentId],
      name: data[i][headerMap.name],
      className: data[i][headerMap.className],
      club: data[i][headerMap.club],
      sessionToken: storedToken,
      sessionExpiresAt: expiresDate
    };

    if (adminOnly && user.className !== "管理員") {
      return { success: false, message: "權限不足。" };
    }

    return {
      success: true,
      message: "登入狀態有效",
      user: user,
      sessionExpiresAt: expiresDate
    };
  }

  return { success: false, message: "找不到登入帳號，請重新登入。" };
}

function uploadPdf_(ss, params) {
  try {
    var studentId = params.studentId;
    var fileName = params.filename || (studentId + "_upload.pdf");
    var fileData = params.fileBase64 || params.file || "";
    var base64 = fileData.indexOf(",") >= 0 ? fileData.split(",")[1] : fileData;

    if (!studentId || !base64) {
      return jsonOutput({ success: false, message: "缺少學號或檔案資料" });
    }

    var task = params.scheduleTask || params.task || "未分類項目";
    var activeSubmission = findActiveSubmission_(ss, studentId, task);
    if (activeSubmission) {
      return jsonOutput({
        success: false,
        message: "此項目已繳交，請先到「已繳交內容」撤銷後再重新提交。"
      });
    }

    var taskName = sanitizeFolderName_(task);
    var mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    var taskFolder = getOrCreateFolder_(mainFolder, taskName);
    var targetFolder = getOrCreateFolder_(taskFolder, studentId);

    var decoded = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(decoded, "application/pdf", fileName);
    var file = targetFolder.createFile(blob);

    appendSubmission_(ss, params, file);

    return jsonOutput({
      success: true,
      message: "檔案上傳成功！",
      url: file.getUrl()
    });
  } catch (err) {
    return jsonOutput({ success: false, message: "上傳失敗: " + err.toString() });
  }
}

function addSchedule_(ss, params) {
  var sheet = ss.getSheetByName("schedule");
  sheet.appendRow([params.task, params.deadline, params.note || ""]);
  return jsonOutput({ success: true, message: "時程已新增" });
}

function deleteSchedule_(ss, params) {
  var sheet = ss.getSheetByName("schedule");
  var rowNumber = Number(params.id) + 2;
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要刪除的時程" });
  }
  sheet.deleteRow(rowNumber);
  return jsonOutput({ success: true, message: "時程已刪除" });
}

function getScheduleRows_(ss) {
  var sheet = ss.getSheetByName("schedule");
  var data = sheet.getDataRange().getValues();
  var scheduleList = [];
  for (var i = 1; i < data.length; i++) {
    scheduleList.push({
      id: i - 1,
      task: data[i][0],
      deadline: data[i][1],
      note: data[i][2]
    });
  }
  return scheduleList;
}

function getSheetObjects_(ss, sheetName, keys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var item = {};
    for (var j = 0; j < keys.length; j++) {
      item[keys[j]] = data[i][j];
    }
    rows.push(item);
  }
  return rows;
}

function appendSubmission_(ss, params, file) {
  var sheet = getSubmissionSheet_(ss);
  sheet.appendRow([
    params.scheduleTask || params.task || "",
    params.studentId || "",
    params.studentName || "",
    params.className || "",
    params.filename || file.getName(),
    file.getUrl(),
    new Date(),
    "active",
    ""
  ]);
}

function getSubmissionSheet_(ss) {
  var sheet = ss.getSheetByName("submissions");
  var headers = ["task", "studentId", "studentName", "className", "filename", "url", "createdAt", "status", "revokedAt"];
  if (!sheet) {
    sheet = ss.insertSheet("submissions");
    sheet.appendRow(headers);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return sheet;
  }

  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (existing[i] !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  return sheet;
}

function getSubmissionRows_(ss) {
  var sheet = getSubmissionSheet_(ss);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1] && !data[i][4]) continue;
    rows.push({
      id: i + 1,
      task: data[i][0],
      studentId: data[i][1],
      studentName: data[i][2],
      className: data[i][3],
      filename: data[i][4],
      url: data[i][5],
      createdAt: data[i][6],
      status: data[i][7] || "active",
      revokedAt: data[i][8] || ""
    });
  }
  return rows;
}

function findActiveSubmission_(ss, studentId, task) {
  var rows = getSubmissionRows_(ss);
  for (var i = 0; i < rows.length; i++) {
    if (
      rows[i].studentId.toString() === studentId.toString() &&
      rows[i].task.toString() === task.toString() &&
      rows[i].status !== "revoked"
    ) {
      return rows[i];
    }
  }
  return null;
}

function revokeSubmission_(ss, params) {
  var sheet = getSubmissionSheet_(ss);
  var rowNumber = Number(params.id);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要撤銷的繳交紀錄" });
  }

  var row = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];
  if (row[1].toString() !== (params.studentId || "").toString()) {
    return jsonOutput({ success: false, message: "只能撤銷自己的繳交紀錄" });
  }
  if ((row[7] || "active") === "revoked") {
    return jsonOutput({ success: false, message: "此繳交紀錄已撤銷" });
  }

  var deleteResult = trashDriveFileByUrl_(row[5]);
  if (!deleteResult.success) {
    return jsonOutput(deleteResult);
  }

  sheet.getRange(rowNumber, 8).setValue("revoked");
  sheet.getRange(rowNumber, 9).setValue(new Date());
  sheet.getRange(rowNumber, 6).setValue("");
  return jsonOutput({ success: true, message: "已撤銷繳交紀錄並刪除雲端檔案，可以重新提交此項目。" });
}

function trashDriveFileByUrl_(url) {
  if (!url) {
    return { success: true, message: "沒有檔案連結，僅撤銷紀錄。" };
  }

  var fileId = extractDriveFileId_(url);
  if (!fileId) {
    return { success: false, message: "無法解析雲端檔案連結，未撤銷繳交紀錄。" };
  }

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true, message: "檔案已移至垃圾桶。" };
  } catch (err) {
    return { success: false, message: "雲端檔案刪除失敗，未撤銷繳交紀錄: " + err.toString() };
  }
}

function extractDriveFileId_(url) {
  var text = String(url || "");
  var patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match && match[1]) return match[1];
  }

  if (/^[a-zA-Z0-9_-]{20,}$/.test(text)) return text;
  return "";
}

function clearSubmissions_(ss, params) {
  if ((params.password || "").toString() !== "990607") {
    return jsonOutput({ success: false, message: "密碼錯誤，未清除繳交紀錄。" });
  }

  var sheet = getSubmissionSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return jsonOutput({ success: true, message: "繳交紀錄已清除。" });
}

function getOrCreateFolder_(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function sanitizeFolderName_(name) {
  return String(name || "未分類項目")
    .replace(/[\\/:*?"<>|#%{}~&]/g, "_")
    .trim() || "未分類項目";
}

function authCheck_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    return jsonOutput({
      success: true,
      message: "Drive 與 Sheet 權限正常",
      spreadsheet: ss.getName(),
      folder: folder.getName()
    });
  } catch (err) {
    return jsonOutput({
      success: false,
      message: "授權檢查失敗: " + err.toString()
    });
  }
}

function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}
