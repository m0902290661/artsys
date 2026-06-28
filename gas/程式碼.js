/**
 * ==========================================
 * 弘明實驗高級中學 - 美術歷程管理系統核心後端 (升級版)
 * ==========================================
 * 建立日期：2026
 * 功能：支援管理員後台時程、解鎖、問題回報、學生帳號管理(全資料修改+動態選單)，以及學生端功能。
 */

// 1. 全域環境變數配置 (請根據你的雲端硬碟實際 ID 調整)
const SPREADSHEET_ID = "1O2wUFvNN9mnIaI09kxaeyHI_UcDy77u5E1TTVlz82uA";
const MAIN_FOLDER_ID = "1yrORQBOk72ghb5vVPRhMs0gSOVy0HEPi";
const SESSION_TTL_HOURS = 8;

// 2. 資料庫欄位結構定義 (Headers)
const USER_HEADERS = [
  "studentId", "name", "className", "club", "password", "sessionToken", "sessionExpiresAt"
];
const REPORT_HEADERS = [
  "createdAt", "reportType", "studentId", "studentName", "className", "contact", "subject", "description", "status", "handledAt", "handledBy"
];
const SUBMISSION_HEADERS = [
  "task", "studentId", "studentName", "className", "filename", "url", "createdAt", "status", "revokedAt"
];
const UNLOCK_HEADERS = [
  "studentId", "task"
];
const SCHEDULE_HEADERS = [
  "task", "deadline", "note"
];
const OPTION_HEADERS = [
  "type", "value" // type: 'class' 或 'club'
];

/**
 * ==========================================
 * 主進入點：doGet (處理前端所有 GET 請求)
 * ==========================================
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action;
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(err) {
    return jsonOutput({ success: false, message: "資料庫連線失敗，請檢查 SPREADSHEET_ID 是否正確。" });
  }

  // 路由分流
  if (action === "getSchedule") {
    return jsonOutput(getScheduleRows_(ss));
  }

  if (action === "getOptions") {
    return jsonOutput(getOptionRows_(ss));
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
    return jsonOutput(getSubmissionRows_(ss).filter(function (item) {
      return item.studentId.toString() === (params.studentId || "").toString();
    }));
  }

  if (action === "getUsers") {
    var usersSession = requireSession_(ss, params, true);
    if (!usersSession.success) return jsonOutput(usersSession);
    
    // 管理員需要讀取密碼來做編輯時的預填，安全起見 sessionToken 不洩漏
    return jsonOutput(getSheetObjects_(ss, "users", USER_HEADERS).map(function (user) {
      delete user.sessionToken;
      delete user.sessionExpiresAt;
      return user;
    }));
  }

  if (action === "getReports") {
    var reportsSession = requireSession_(ss, params, true);
    if (!reportsSession.success) return jsonOutput(reportsSession);
    return jsonOutput(getReportRows_(ss));
  }

  if (action === "validateSession") {
    return jsonOutput(requireSession_(ss, params, false));
  }

  if (action === "authCheck") {
    return authCheck_();
  }

  return jsonOutput({ success: false, message: "未知的 GET action" });
}

/**
 * ==========================================
 * 主進入點：doPost (處理前端所有 POST 請求)
 * ==========================================
 */
function doPost(e) {
  var params = getParams(e);
  var action = params.action;
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(err) {
    return jsonOutput({ success: false, message: "資料庫連線失敗，請檢查 SPREADSHEET_ID 是否正確。" });
  }
  
  var userSheet = getSheet_(ss, "users", USER_HEADERS);
  var userData = userSheet.getDataRange().getValues();

  // 無需 Session 權限的指令
  if (action === "login") {
    return login_(userSheet, userData, params);
  }
  if (action === "signup") {
    return signup_(userSheet, userData, params);
  }
  if (action === "logout") {
    return logout_(userSheet, userData, params);
  }

  // 管理員專用指令 (需要管理員 Session 權限)
  if (action === "addSchedule") {
    var addSession = requireSession_(ss, params, true);
    if (!addSession.success) return jsonOutput(addSession);
    return addSchedule_(ss, params);
  }

  if (action === "editSchedule") {
    var editSession = requireSession_(ss, params, true);
    if (!editSession.success) return jsonOutput(editSession);
    return editSchedule_(ss, params);
  }

  if (action === "deleteSchedule") {
    var deleteSession = requireSession_(ss, params, true);
    if (!deleteSession.success) return jsonOutput(deleteSession);
    return deleteSchedule_(ss, params);
  }

  if (action === "addUnlock") {
    var unlockSession = requireSession_(ss, params, true);
    if (!unlockSession.success) return jsonOutput(unlockSession);
    return addUnlock_(ss, params);
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

  if (action === "editUser") {
    var editUserSession = requireSession_(ss, params, true);
    if (!editUserSession.success) return jsonOutput(editUserSession);
    return editUser_(userSheet, userData, params);
  }

  if (action === "updateUserPassword") {
    var passwordSession = requireSession_(ss, params, true);
    if (!passwordSession.success) return jsonOutput(passwordSession);
    return updateUserPassword_(userSheet, userData, params);
  }

  if (action === "updateReportStatus") {
    var reportStatusSession = requireSession_(ss, params, true);
    if (!reportStatusSession.success) return jsonOutput(reportStatusSession);
    return updateReportStatus_(ss, params, reportStatusSession.user);
  }

  // 學生/通用功能指令 (一般登入身分即可)
  if (action === "uploadPDF") {
    var uploadSession = requireSession_(ss, params, false);
    if (!uploadSession.success) return jsonOutput(uploadSession);
    if (uploadSession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return uploadPdf_(ss, params);
  }

  if (action === "revokeSubmission") {
    var revokeSession = requireSession_(ss, params, false);
    if (!revokeSession.success) return jsonOutput(revokeSession);
    if (revokeSession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return revokeSubmission_(ss, params);
  }

  if (action === "submitReport") {
    var reportSession = requireSession_(ss, params, false);
    if (!reportSession.success) return jsonOutput(reportSession);
    if (reportSession.user.studentId.toString() !== (params.studentId || "").toString()) {
      return jsonOutput({ success: false, message: "登入狀態不符，請重新登入。" });
    }
    return submitReport_(ss, params, reportSession.user);
  }

  return jsonOutput({ success: false, message: "未知的 POST action" });
}

/**
 * ==========================================
 * 功能函數實作分區
 * ==========================================
 */

// 讀取班級與美術社團選單設定值
function getOptionRows_(ss) {
  var sheet = getSheet_(ss, "options", OPTION_HEADERS);
  var data = sheet.getDataRange().getValues();
  var classes = [];
  var clubs = [];
  
  // 如果是全新表格，預填預設資料防呆
  if (data.length <= 1) {
    var defaultOptions = [
      ["class", "高一智"], ["class", "高一仁"], ["class", "高一勇"],
      ["club", "素描社"], ["club", "水彩社"], ["club", "油畫社"], ["club", "水墨社"],
      ["club", "書法社"], ["club", "雕塑社"], ["club", "漫畫社"], ["club", "設計社"]
    ];
    defaultOptions.forEach(function(opt) { sheet.appendRow(opt); });
    data = sheet.getDataRange().getValues();
  }

  for (var i = 1; i < data.length; i++) {
    var type = data[i][0].toString().trim();
    var val = data[i][1].toString().trim();
    if (type === "class") classes.push(val);
    if (type === "club") clubs.push(val);
  }
  return { classes: classes, clubs: clubs };
}

// 學生與管理員登入
function login_(userSheet, userData, params) {
  var headerMap = getHeaderMap_(userSheet, USER_HEADERS);
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

// 安全登出機制
function logout_(userSheet, userData, params) {
  var headerMap = getHeaderMap_(userSheet, USER_HEADERS);
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

// 學生首次自主註冊開通
function signup_(userSheet, userData, params) {
  var studentId = (params.studentId || "").toString().trim();
  var password = (params.password || "").toString();
  
  if (!studentId || !password) {
    return jsonOutput({ success: false, message: "缺少學號或密碼配置。" });
  }

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === studentId) {
      if (userData[i][4] !== "") {
        return jsonOutput({ success: false, message: "該學號已開通" });
      }
      userSheet.getRange(i + 1, 5).setValue(password);
      return jsonOutput({ success: true, message: "帳號開通成功！" });
    }
  }
  return jsonOutput({ success: false, message: "找不到此學號，請聯繫管理員建立基本資料。" });
}

// 新增時程項目
function addSchedule_(ss, params) {
  var sheet = getSheet_(ss, "schedule", SCHEDULE_HEADERS);
  sheet.appendRow([params.task, params.deadline, params.note || ""]);
  return jsonOutput({ success: true, message: "時程已新增" });
}

// 編輯變更既有時程
function editSchedule_(ss, params) {
  var sheet = getSheet_(ss, "schedule", SCHEDULE_HEADERS);
  var rowNumber = Number(params.id) + 2; 

  if (isNaN(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要編輯的時程項目" });
  }

  sheet.getRange(rowNumber, 1).setValue(params.task);
  sheet.getRange(rowNumber, 2).setValue(params.deadline);
  sheet.getRange(rowNumber, 3).setValue(params.note || "");
  return jsonOutput({ success: true, message: "時程項目已更新" });
}

// 刪除時程項目
function deleteSchedule_(ss, params) {
  var sheet = getSheet_(ss, "schedule", SCHEDULE_HEADERS);
  var rowNumber = Number(params.id) + 2;
  if (isNaN(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要刪除的時程" });
  }
  sheet.deleteRow(rowNumber);
  return jsonOutput({ success: true, message: "時程已刪除" });
}

// 獲取所有時程數據清單
function getScheduleRows_(ss) {
  var sheet = getSheet_(ss, "schedule", SCHEDULE_HEADERS);
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

// 授予遲交解鎖權限 (由管理員呼叫)
function addUnlock_(ss, params) {
  var sheet = getSheet_(ss, "unlocks", UNLOCK_HEADERS);
  sheet.appendRow([params.targetStudentId, params.task]);
  return jsonOutput({ success: true, message: "已授予該生遲交權限。" });
}

// 檢查並核對遲交狀態 (讀取後隨即銷毀，實現單次授權)
function isUnlocked_(ss, studentId, task) {
  var sheet = getSheet_(ss, "unlocks", UNLOCK_HEADERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === studentId.toString() && data[i][1].toString() === task.toString()) {
      sheet.deleteRow(i + 1); 
      return true;
    }
  }
  return false;
}

// 學生 PDF 上傳整合系統
function uploadPdf_(ss, params) {
  try {
    var studentId = (params.studentId || "").toString().trim();
    var task = (params.scheduleTask || params.task || "未分類項目").toString().trim();
    var fileName = params.filename || (studentId + "_" + task + ".pdf");
    var fileData = params.fileBase64 || params.file || "";
    
    if (!studentId || !fileData) {
      return jsonOutput({ success: false, message: "缺少學號或檔案數據。" });
    }
    
    var base64 = fileData.indexOf(",") >= 0 ? fileData.split(",")[1] : fileData;

    // 遲交限制時間阻斷與解鎖校驗
    var sched = getScheduleRows_(ss);
    var item = sched.find(function(r) { return r.task === task; });
    if (item && item.deadline) {
      var deadlineDate = new Date(item.deadline);
      if (deadlineDate.getHours() === 0 && deadlineDate.getMinutes() === 0) {
        deadlineDate.setHours(23, 59, 59, 999); 
      }
      if (new Date() > deadlineDate) {
        if (!isUnlocked_(ss, studentId, task)) {
          return jsonOutput({ success: false, message: "已超過截止日期，且未獲管理員解鎖權限。" });
        }
      }
    }

    // 重複上傳有效件檢查
    var activeSubmission = findActiveSubmission_(ss, studentId, task);
    if (activeSubmission) {
      return jsonOutput({
        success: false,
        message: "此項目已繳交。請先至「已繳交內容」撤銷後，再重新提交。"
      });
    }

    // Drive 資料夾動態樹狀建立與儲存
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

// 新增繳交紀錄數據
function appendSubmission_(ss, params, file) {
  var sheet = getSheet_(ss, "submissions", SUBMISSION_HEADERS);
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

// 檢索所有檔案提交紀錄
function getSubmissionRows_(ss) {
  var sheet = getSheet_(ss, "submissions", SUBMISSION_HEADERS);
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

// 搜尋特定的有效上傳物件
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

// 學生自主撤銷檔案機制與雲端連動刪除
function revokeSubmission_(ss, params) {
  var sheet = getSheet_(ss, "submissions", SUBMISSION_HEADERS);
  var rowNumber = Number(params.id);
  if (isNaN(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要撤銷的繳交紀錄" });
  }

  var row = sheet.getRange(rowNumber, 1, 1, SUBMISSION_HEADERS.length).getValues()[0];
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

// 刪除 Google Drive 上的實體實檔案
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

// 正則表達式解析網址中的 File ID
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

// 建立全新學生資料
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
  return jsonOutput({ success: true, message: "人員帳號建立完畢。" });
}

// 管理員變更、覆蓋用戶所有欄位資料
function editUser_(userSheet, userData, params) {
  var originalStudentId = (params.id || "").toString().trim(); // 原本的學號 (識別行用)
  var newStudentId = (params.newStudentId || "").toString().trim();
  var name = (params.name || "").toString().trim();
  var className = (params.className || "").toString().trim();
  var club = (params.club || "").toString().trim();
  var password = (params.password || "").toString();

  if (!originalStudentId || !newStudentId || !name || !className) {
    return jsonOutput({ success: false, message: "學號、姓名與班級皆為必填欄位。" });
  }

  // 檢查是否將學號改成與他人重複
  if (originalStudentId !== newStudentId) {
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][0].toString() === newStudentId) {
        return jsonOutput({ success: false, message: "新的學號已被其他學生使用，未變更。" });
      }
    }
  }

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === originalStudentId) {
      userSheet.getRange(i + 1, 1).setValue(newStudentId);
      userSheet.getRange(i + 1, 2).setValue(name);
      userSheet.getRange(i + 1, 3).setValue(className);
      userSheet.getRange(i + 1, 4).setValue(club);
      if (password) {
        userSheet.getRange(i + 1, 5).setValue(password);
        // 密碼有改則強制把舊的登入 Session 清空
        userSheet.getRange(i + 1, 6).setValue("");
        userSheet.getRange(i + 1, 7).setValue("");
      }
      SpreadsheetApp.flush();
      return jsonOutput({ success: true, message: "學生資料已完全更新。" });
    }
  }
  return jsonOutput({ success: false, message: "找不到指定的學生帳號。" });
}

// 強制蓋寫、覆蓋特定帳號密碼 (保留舊有的獨立接口，供防呆調配)
function updateUserPassword_(userSheet, userData, params) {
  var studentId = (params.targetStudentId || "").toString().trim();
  var password = (params.password || "").toString();

  if (!studentId || !password) {
    return jsonOutput({ success: false, message: "缺少學號或新密碼參數。" });
  }
  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0].toString() === studentId) {
      userSheet.getRange(i + 1, 5).setValue(password);
      userSheet.getRange(i + 1, 6).setValue(""); 
      userSheet.getRange(i + 1, 7).setValue(""); 
      return jsonOutput({ success: true, message: "密碼已更新，該帳號需要重新登入。" });
    }
  }
  return jsonOutput({ success: false, message: "找不到此學號。" });
}

// 學生端發送系統異常或密碼重設單
function submitReport_(ss, params, user) {
  var reportType = (params.reportType || "").toString().trim();
  var subject = (params.subject || "").toString().trim();
  var description = (params.description || "").toString().trim();
  var contact = (params.contact || "").toString().trim();

  if (!reportType || !subject || !description) {
    return jsonOutput({ success: false, message: "請填寫問題類型、主旨與問題說明。" });
  }

  var sheet = getSheet_(ss, "reports", REPORT_HEADERS);
  sheet.appendRow([
    new Date(), reportType, user.studentId || "", user.name || "", user.className || "",
    contact, subject, description, "pending", "", ""
  ]);
  return jsonOutput({ success: true, message: "問題回報已送出，請等待管理員處理。" });
}

// 獲取所有問題單清單
function getReportRows_(ss) {
  var sheet = getSheet_(ss, "reports", REPORT_HEADERS);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][2] && !data[i][6]) continue;
    rows.push({
      id: i + 1,
      createdAt: data[i][0],
      reportType: data[i][1],
      studentId: data[i][2],
      studentName: data[i][3],
      className: data[i][4],
      contact: data[i][5],
      subject: data[i][6],
      description: data[i][7],
      status: data[i][8] || "pending",
      handledAt: data[i][9] || "",
      handledBy: data[i][10] || ""
    });
  }
  return rows;
}

// 管理員變更單據狀態 (待處理/處理中/已完成)
function updateReportStatus_(ss, params, adminUser) {
  var sheet = getSheet_(ss, "reports", REPORT_HEADERS);
  var rowNumber = Number(params.id);
  var status = (params.status || "").toString();
  var allowedStatuses = ["pending", "processing", "done"];

  if (isNaN(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return jsonOutput({ success: false, message: "找不到要更新的回報單據。" });
  }
  if (allowedStatuses.indexOf(status) === -1) {
    return jsonOutput({ success: false, message: "變更狀態值不合法。" });
  }

  sheet.getRange(rowNumber, 9).setValue(status);
  if (status === "done") {
    sheet.getRange(rowNumber, 10).setValue(new Date());
    sheet.getRange(rowNumber, 11).setValue(adminUser.name || adminUser.studentId || "管理員");
  } else {
    sheet.getRange(rowNumber, 10).setValue("");
    sheet.getRange(rowNumber, 11).setValue("");
  }
  return jsonOutput({ success: true, message: "回報狀態已更新。" });
}

// 刪除清空全校繳交資料表 (後台二次密碼確認)
function clearSubmissions_(ss, params) {
  if ((params.password || "").toString() !== "990607") {
    return jsonOutput({ success: false, message: "管理確認碼錯誤，未清除繳交紀錄。" });
  }
  var sheet = getSheet_(ss, "submissions", SUBMISSION_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return jsonOutput({ success: true, message: "全校美術檔案繳交紀錄已清空。" });
}

/**
 * ==========================================
 * 核心安全阻斷器與底層 Helper 函數群
 * ==========================================
 */

// 強制校驗 Token 與逾時，並提供管理員權限級別篩選
function requireSession_(ss, params, adminOnly) {
  var studentId = (params.studentId || params.adminId || "").toString();
  var sessionToken = (params.sessionToken || "").toString();

  if (!studentId || !sessionToken) {
    return { success: false, message: "登入驗證已失效，請重新登入。" };
  }

  var userSheet = getSheet_(ss, "users", USER_HEADERS);
  var headerMap = getHeaderMap_(userSheet, USER_HEADERS);
  var data = userSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][headerMap.studentId].toString() !== studentId) continue;

    var storedToken = (data[i][headerMap.sessionToken] || "").toString();
    var expiresAt = data[i][headerMap.sessionExpiresAt];
    var expiresDate = expiresAt instanceof Date ? expiresAt : new Date(expiresAt);

    if (!storedToken || storedToken !== sessionToken) {
      return { success: false, message: "帳號在其他地方登入或密鑰不符，請重新登入。" };
    }
    if (!expiresAt || Number.isNaN(expiresDate.getTime()) || expiresDate.getTime() <= new Date().getTime()) {
      return { success: false, message: "登入已超過安全時效，請重新登入。" };
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
      return { success: false, message: "您不具備管理員權限，系統已拒絕存取。" };
    }

    return {
      success: true,
      message: "驗證通過",
      user: user,
      sessionExpiresAt: expiresDate
    };
  }
  return { success: false, message: "無此帳號資料，請重新登入。" };
}

// 工作表自動建構子與標頭自我修復函數
function getSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return sheet;
  }
  var existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (existingHeaders[i] !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  return sheet;
}

// 建立欄位名稱字串與陣列索引的對照
function getHeaderMap_(sheet, expectedHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) map[headers[i]] = i;
  }
  return map;
}

// 把試算表二維陣列轉換成標準 JSON 物件陣列
function getSheetObjects_(ss, sheetName, keys) {
  var sheet = getSheet_(ss, sheetName, keys);
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

// 檔案夾階層檢查與動態生成
function getOrCreateFolder_(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

// 過濾不合規的資料夾特殊命名符號
function sanitizeFolderName_(name) {
  return String(name || "未分類項目")
    .replace(/[\\/:*?"<>|#%{}~&]/g, "_")
    .trim() || "未分類項目";
}

// 格式化輸出 JSON
function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 解析相容各式前端發送的引數形態 (URLSearchParams / Payload JSON)
function getParams(e) {
  if (e && e.postData && e.postData.contents) {
    var text = e.postData.contents;
    if (text.trim().charAt(0) === "{") {
      return JSON.parse(text);
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}

// 提供跨域預檢處理 (CORS Header Options)
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// 首次手動綁定授權與除錯檢查環境專用函數
function authCheck_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    return jsonOutput({
      success: true,
      message: "雲端 Sheet 與 Drive 授權連接狀態：完全正常",
      spreadsheet: ss.getName(),
      folder: folder.getName()
    });
  } catch (err) {
    return jsonOutput({
      success: false,
      message: "權限檢查時發生錯誤: " + err.toString()
    });
  }
}