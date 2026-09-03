/**
 * =========================================================================
 * BACKEND GOOGLE APPS SCRIPT - CÔNG AN XÃ ĐỨC HỢP, TỈNH HƯNG YÊN
 * Hòm thư tiếp nhận tin báo tố giác tội phạm & Quản lý hồ sơ nghiệp vụ ANTT
 * =========================================================================
 */

var CONFIG = {
  UNIT_NAME: "CÔNG AN XÃ ĐỨC HỢP",
  PROVINCE: "CÔNG AN TỈNH HƯNG YÊN",
  HOTLINE: "02213.815.999",
  EMAIL_RECEIVER: "conganxaduchopdangbai@gmail.com",
  DRIVE_FOLDER_NAME: "CAX_DucHop_HoSoToGiac_DinhKem",
  VILLAGES: [
    "Đức An", "Đức Trung", "Phú Ninh", "Nho Lâm", "Hạnh Lâm", 
    "Vân Nghệ", "Trung Hòa", "Phú Cường", "Quảng Lạc", "Bắc Nam Phú", "Tây Thịnh"
  ],
  SHEETS: {
    CASES: "CASES",
    USERS: "USERS",
    ATTACHMENTS: "ATTACHMENTS",
    PROCESS_HISTORY: "PROCESS_HISTORY",
    AUDIT_LOG: "AUDIT_LOG"
  }
};

function doGet(e) {
  initSystemDatabase();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Hòm Thư Tố Giác Tội Phạm - Công An Xã Đức Hợp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty("SPREADSHEET_ID");
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create("QL_ToGiac_CongAnXa_DucHop");
      props.setProperty("SPREADSHEET_ID", ss.getId());
    }
  }
  return ss;
}

function getOrCreateUploadFolder() {
  var folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  var newFolder = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

function ensureSheetHeaders(sheet, requiredHeaders) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return;
  }
  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  for (var i = 0; i < requiredHeaders.length; i++) {
    var colName = requiredHeaders[i];
    if (currentHeaders.indexOf(colName) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(colName);
      currentHeaders.push(colName);
    }
  }
}

function initSystemDatabase() {
  var ss = getSpreadsheet();
  
  // 1. Sheet USERS
  var userHeaders = ["userId", "username", "passwordHash", "fullName", "badgeNumber", "rank", "position", "role", "phone", "assignedVillages", "status", "createdAt"];
  var userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS) || ss.insertSheet(CONFIG.SHEETS.USERS);
  ensureSheetHeaders(userSheet, userHeaders);
  
  if (userSheet.getLastRow() <= 1) {
    var defaultOfficers = [
      ["USR-001", "admin", "admin@123", "Quản Trị Viên Hệ Thống", "343-000", "Thượng tá", "Quản trị hệ thống", "SUPER_ADMIN", "02213.815.999", "Đức An, Đức Trung, Phú Ninh, Nho Lâm, Hạnh Lâm, Vân Nghệ, Trung Hòa, Phú Cường, Quảng Lạc, Bắc Nam Phú, Tây Thịnh", "ACTIVE", new Date().toISOString()],
      ["USR-002", "Quang343001", "1", "Đoàn Ngọc Quang", "343-001", "Thượng tá", "Trưởng CAX", "CHIEF", "0983.892.222", "Đức An, Đức Trung, Phú Ninh, Nho Lâm, Hạnh Lâm, Vân Nghệ, Trung Hòa, Phú Cường, Quảng Lạc, Bắc Nam Phú, Tây Thịnh", "ACTIVE", new Date().toISOString()],
      ["USR-003", "Hai343002", "1", "Phạm Văn Hài", "343-002", "Thiếu tá", "Phó CAX PCTP", "DEPUTY_CHIEF", "0986.106.548", "Đức Trung, Phú Ninh, Tây Thịnh", "ACTIVE", new Date().toISOString()],
      ["USR-004", "Thu343003", "1", "Vũ Văn Thu", "343-003", "Thiếu tá", "Phó CAX AN", "DEPUTY_CHIEF", "0987.827.336", "Đức An, Nho Lâm, Quảng Lạc", "ACTIVE", new Date().toISOString()],
      ["USR-005", "Ngoc343004", "1", "Đặng Hồng Ngọc", "343-004", "Trung tá", "Phó CAX Tổng hợp", "DEPUTY_CHIEF", "0944.061.666", "Hạnh Lâm, Vân Nghệ, Bắc Nam Phú", "ACTIVE", new Date().toISOString()],
      ["USR-006", "ThuCSTT343005", "1", "Vũ Văn Thu", "343-005", "Trung tá", "Phó CAX CSTT", "DEPUTY_CHIEF", "0988.178.118", "Trung Hòa, Phú Cường", "ACTIVE", new Date().toISOString()],
      ["USR-007", "Doanh343006", "1", "Nguyễn Văn Doanh", "343-006", "Trung tá", "Phó CAX CSKV", "DEPUTY_CHIEF", "0987.668.867", "Đức An, Đức Trung, Phú Ninh, Nho Lâm, Hạnh Lâm, Vân Nghệ, Trung Hòa, Phú Cường, Quảng Lạc, Bắc Nam Phú, Tây Thịnh", "ACTIVE", new Date().toISOString()],
      ["USR-008", "Dung343007", "1", "Nguyễn Tiến Dũng", "343-007", "Đại úy", "Cảnh sát khu vực", "PROCESSING_OFFICER", "0977.123.456", "Nho Lâm, Đức An", "ACTIVE", new Date().toISOString()],
      ["USR-009", "Hoang343008", "1", "Trần Văn Hoàng", "343-008", "Thượng úy", "Cảnh sát khu vực", "PROCESSING_OFFICER", "0988.234.567", "Đức Trung, Phú Ninh", "ACTIVE", new Date().toISOString()],
      ["USR-010", "Minh343009", "1", "Lê Quang Minh", "343-009", "Đại úy", "Cảnh sát khu vực", "PROCESSING_OFFICER", "0912.345.678", "Vân Nghệ, Hạnh Lâm", "ACTIVE", new Date().toISOString()]
    ];
    defaultOfficers.forEach(function(row) { userSheet.appendRow(row); });
  }

  // 2. Sheet CASES
  var caseHeaders = [
    "caseId", "createdAt", "village", "incidentLocation", "category", "priority", 
    "description", "suspectDescription", "isAnonymous", "reporterName", "reporterPhone", 
    "status", "assignedDeputyId", "assignedDeputyName", "assignedOfficerIds", "assignedOfficerNames", 
    "directiveNote", "deadline", "resolutionNote", "updatedAt"
  ];
  var casesSheet = ss.getSheetByName(CONFIG.SHEETS.CASES) || ss.insertSheet(CONFIG.SHEETS.CASES);
  ensureSheetHeaders(casesSheet, caseHeaders);

  // 3. Các Sheet còn lại
  ensureSheetHeaders(ss.getSheetByName(CONFIG.SHEETS.ATTACHMENTS) || ss.insertSheet(CONFIG.SHEETS.ATTACHMENTS), ["attachmentId", "caseId", "fileName", "mimeType", "fileSize", "driveFileId", "driveUrl", "uploadedAt"]);
  ensureSheetHeaders(ss.getSheetByName(CONFIG.SHEETS.PROCESS_HISTORY) || ss.insertSheet(CONFIG.SHEETS.PROCESS_HISTORY), ["historyId", "caseId", "action", "toStatus", "performedBy", "performedAt", "note"]);
  ensureSheetHeaders(ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG) || ss.insertSheet(CONFIG.SHEETS.AUDIT_LOG), ["logId", "timestamp", "userId", "userName", "userRole", "action", "entity", "entityId", "details"]);
}

function getSheetData(sheetName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    var headers = values[0];
    var list = [];
    for (var r = 1; r < values.length; r++) {
      var item = {};
      for (var c = 0; c < headers.length; c++) item[headers[c]] = values[r][c];
      list.push(item);
    }
    return list;
  } catch (e) { return []; }
}

function appendSheetRow(sheetName, obj) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    
    for (var k in obj) {
      if (headers.indexOf(k) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(k);
        headers.push(k);
      }
    }
    var row = headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ""; });
    sheet.appendRow(row);
    return true;
  } catch(e) { return false; } finally { lock.releaseLock(); }
}

function updateSheetRow(sheetName, keyField, keyValue, updateObj) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return false;
    var headers = data[0];
    
    for (var k in updateObj) {
      if (headers.indexOf(k) === -1) {
        var nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(k);
        headers.push(k);
      }
    }
    
    var keyColIndex = headers.indexOf(keyField);
    if (keyColIndex === -1) return false;
    var cleanKeyVal = String(keyValue || "").trim().toUpperCase();

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][keyColIndex]).trim().toUpperCase() === cleanKeyVal) {
        for (var key in updateObj) {
          var cIndex = headers.indexOf(key);
          if (cIndex !== -1) {
            sheet.getRange(r + 1, cIndex + 1).setValue(updateObj[key]);
          }
        }
        return true;
      }
    }
    return false;
  } catch(e) { return false; } finally { lock.releaseLock(); }
}

function writeAuditLog(user, action, entity, entityId, details) {
  try {
    appendSheetRow(CONFIG.SHEETS.AUDIT_LOG, {
      logId: "LOG-" + new Date().getTime(),
      timestamp: new Date().toISOString(),
      userId: user ? user.userId : "ANONYMOUS",
      userName: user ? (user.rank + " " + user.fullName) : "Hệ thống / Công dân",
      userRole: user ? user.role : "GUEST",
      action: action,
      entity: entity,
      entityId: entityId,
      details: details || ""
    });
  } catch (e) {}
}

/* =========================================================================
   1. XÁC THỰC & ĐỔI MẬT KHẨU
   ========================================================================= */

function authenticateOfficer(username, password) {
  initSystemDatabase();
  var u = String(username || "").trim().toLowerCase();
  var p = String(password || "").trim();
  var users = getSheetData(CONFIG.SHEETS.USERS);
  var user = users.find(function(item) {
    return String(item.username || "").toLowerCase() === u || String(item.badgeNumber || "").toLowerCase() === u;
  });
  
  if (!user) return { success: false, message: "Tài khoản hoặc số hiệu CAND không chính xác." };
  
  var storedPass = String(user.passwordHash !== undefined && user.passwordHash !== "" ? user.passwordHash : "1").trim();
  if (p !== storedPass) {
    return { success: false, message: "Mật khẩu không đúng. Vui lòng kiểm tra lại." };
  }

  if (user.status === "INACTIVE") return { success: false, message: "Tài khoản đang bị tạm khóa." };
  
  var vList = user.assignedVillages ? String(user.assignedVillages).split(",").map(function(s) { return s.trim(); }) : [];
  var resData = {
    userId: user.userId, username: user.username, fullName: user.fullName,
    badgeNumber: user.badgeNumber, rank: user.rank, position: user.position,
    role: user.role, phone: user.phone, assignedVillages: vList
  };
  writeAuditLog(resData, "LOGIN", "AUTH", user.userId, "Đăng nhập hệ thống thành công");
  return { success: true, data: resData };
}

function changeOfficerPassword(oldPass, newPass, currentUser) {
  if (!currentUser) return { success: false, message: "Chưa đăng nhập." };
  var users = getSheetData(CONFIG.SHEETS.USERS);
  var user = users.find(function(u) { return u.userId === currentUser.userId; });
  
  var currentStoredPass = String(user.passwordHash !== undefined && user.passwordHash !== "" ? user.passwordHash : "1").trim();
  if (!user || currentStoredPass !== String(oldPass).trim()) {
    return { success: false, message: "Mật khẩu hiện tại không đúng." };
  }

  var updated = updateSheetRow(CONFIG.SHEETS.USERS, "userId", currentUser.userId, { passwordHash: String(newPass).trim() });
  if (!updated) {
    return { success: false, message: "Lỗi cập nhật mật khẩu trên hệ thống!" };
  }
  
  writeAuditLog(currentUser, "CHANGE_PASSWORD", "USERS", currentUser.userId, "Đổi mật khẩu tài khoản thành công");
  return { success: true, message: "Đổi mật khẩu thành công! Hãy ghi nhớ mật khẩu mới." };
}

/* =========================================================================
   2. TIẾP NHẬN TIN BÁO TỐ GIÁC
   ========================================================================= */

function submitCitizenCase(payload) {
  try {
    initSystemDatabase();
    var now = new Date();
    var caseId = "DH-" + now.getFullYear() + "-" + Math.floor(10000 + Math.random() * 90000);
    var isAnon = payload.isAnonymous === true || payload.isAnonymous === "true" || payload.isAnonymous === "YES";
    
    var caseRecord = {
      caseId: caseId, createdAt: now.toISOString(), village: payload.village || "Nho Lâm",
      incidentLocation: payload.incidentLocation || "", category: payload.category || "Tố giác tội phạm",
      priority: payload.priority || "MEDIUM", description: payload.description || "",
      suspectDescription: payload.suspectDescription || "", isAnonymous: isAnon ? "YES" : "NO",
      reporterName: isAnon ? "Công dân (Ẩn danh)" : (payload.reporterName || "Công dân"),
      reporterPhone: isAnon ? "" : (payload.reporterPhone || ""), status: "NEW",
      assignedDeputyId: "", assignedDeputyName: "", assignedOfficerIds: "", assignedOfficerNames: "",
      directiveNote: "", deadline: new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString(),
      resolutionNote: "", updatedAt: now.toISOString()
    };
    appendSheetRow(CONFIG.SHEETS.CASES, caseRecord);
    
    var attCount = 0;
    if (payload.files && payload.files.length > 0) {
      var folder = getOrCreateUploadFolder();
      payload.files.forEach(function(f, idx) {
        if (f.base64Content) {
          try {
            var blob = Utilities.newBlob(Utilities.base64Decode(f.base64Content), f.mimeType || "application/octet-stream", caseId + "_" + (f.fileName || ("Tep_" + idx)));
            var dFile = folder.createFile(blob);
            dFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            appendSheetRow(CONFIG.SHEETS.ATTACHMENTS, {
              attachmentId: "ATT-" + now.getTime() + "-" + idx, caseId: caseId,
              fileName: f.fileName || dFile.getName(), mimeType: f.mimeType || "",
              fileSize: f.size || "", driveFileId: dFile.getId(), driveUrl: dFile.getUrl(), uploadedAt: now.toISOString()
            });
            attCount++;
          } catch(e) {}
        }
      });
    }
    
    appendSheetRow(CONFIG.SHEETS.PROCESS_HISTORY, {
      historyId: "HIST-" + now.getTime(), caseId: caseId, action: "Tiếp nhận tin báo",
      toStatus: "NEW", performedBy: isAnon ? "Hệ thống (Ẩn danh)" : (payload.reporterName || "Người dân"),
      performedAt: now.toISOString(), note: "Tiếp nhận tại Thôn " + caseRecord.village
    });
    
    try {
      MailApp.sendEmail(
        CONFIG.EMAIL_RECEIVER,
        "🚨 [CAND ĐỨC HỢP] TIN BÁO MỚI - " + caseId + " (Thôn " + caseRecord.village + ")",
        "Mã hồ sơ: " + caseId + "\nĐịa bàn: Thôn " + caseRecord.village + "\nNội dung: " + caseRecord.description + "\nNgười báo: " + caseRecord.reporterName
      );
    } catch(mErr) {}
    
    return { success: true, caseId: caseId, message: "Tiếp nhận tin báo thành công!" };
  } catch (err) { return { success: false, message: err.message }; }
}

function trackCitizenCase(caseId) {
  var cleanId = String(caseId || "").trim().toUpperCase();
  var cases = getSheetData(CONFIG.SHEETS.CASES);
  var record = cases.find(function(c) { return String(c.caseId).trim().toUpperCase() === cleanId; });
  if (!record) return { success: false, message: "Không tìm thấy mã hồ sơ: " + cleanId };
  return { success: true, data: record };
}

/* =========================================================================
   3. PHÂN CÔNG CHỦ TRÌ & THỰC HIỆN (CHỈ CHỈ HUY & ADMIN ĐƯỢC PHÉP)
   ========================================================================= */

function assignCaseToLeadershipAndOfficers(payload, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };

  // Kiểm tra quyền hạn nghiêm ngặt
  var r = String(currentUser.role || "").toUpperCase();
  var isLeadership = (r === "SUPER_ADMIN" || r === "ADMIN" || r === "CHIEF" || r === "DEPUTY_CHIEF");
  if (!isLeadership) {
    return { success: false, message: "Từ chối quyền: Cán bộ không có quyền phân công hồ sơ." };
  }

  var caseId = String(payload.caseId || "").trim();
  var deputyId = String(payload.deputyId || "").trim();
  var deputyName = String(payload.deputyName || "").trim();
  var officerIds = payload.officerIds || [];
  var officerNames = payload.officerNames || [];
  var directive = payload.directive || "";
  var deadline = payload.deadline || "";
  
  var users = getSheetData(CONFIG.SHEETS.USERS);

  if (deputyId && (!deputyName || deputyName.indexOf("Chọn") !== -1)) {
    var depUser = users.find(function(u) { return u.userId === deputyId; });
    if (depUser) deputyName = depUser.rank + " " + depUser.fullName;
  }

  if (Array.isArray(officerIds) && officerIds.length > 0 && (!officerNames || officerNames.length === 0)) {
    officerNames = officerIds.map(function(oId) {
      var offUser = users.find(function(u) { return u.userId === oId; });
      return offUser ? (offUser.rank + " " + offUser.fullName) : oId;
    });
  }

  var deputyIdStr = deputyId;
  var deputyNameStr = deputyName;
  var officerIdsStr = Array.isArray(officerIds) ? officerIds.join(", ") : String(officerIds || "");
  var officerNamesStr = Array.isArray(officerNames) ? officerNames.join(", ") : String(officerNames || "");
  var now = new Date().toISOString();
  
  var updated = updateSheetRow(CONFIG.SHEETS.CASES, "caseId", caseId, {
    assignedDeputyId: deputyIdStr,
    assignedDeputyName: deputyNameStr,
    assignedOfficerIds: officerIdsStr,
    assignedOfficerNames: officerNamesStr,
    directiveNote: directive,
    deadline: deadline || new Date(Date.now() + 7*24*60*60*1000).toISOString(),
    status: "ASSIGNED",
    updatedAt: now
  });

  if (!updated) return { success: false, message: "Không tìm thấy hồ sơ để cập nhật." };
  
  var historySummary = "Phó Trưởng CAX chủ trì: " + (deputyNameStr || "Chưa chỉ định") +
                       " | Cán bộ: " + (officerNamesStr || "Chưa chọn") +
                       (directive ? (" | Chỉ đạo: " + directive) : "");

  appendSheetRow(CONFIG.SHEETS.PROCESS_HISTORY, {
    historyId: "HIST-" + new Date().getTime(), caseId: caseId,
    action: "Phân công Phó CAX chủ trì & Cán bộ thực hiện", toStatus: "ASSIGNED",
    performedBy: currentUser.rank + " " + currentUser.fullName, performedAt: now, note: historySummary
  });
  writeAuditLog(currentUser, "ASSIGN_CASE", "CASES", caseId, historySummary);
  
  return { 
    success: true, 
    message: "Đã phân công thành công!",
    data: {
      assignedDeputyId: deputyIdStr, assignedDeputyName: deputyNameStr,
      assignedOfficerIds: officerIdsStr, assignedOfficerNames: officerNamesStr
    }
  };
}

function getOfficerCases(filter, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var cases = getSheetData(CONFIG.SHEETS.CASES);
  var r = String(currentUser.role || "").toUpperCase();
  var isLeadership = (r === "SUPER_ADMIN" || r === "ADMIN" || r === "CHIEF" || r === "DEPUTY_CHIEF");

  var accessible = cases.filter(function(c) {
    if (isLeadership) return true;
    var vList = currentUser.assignedVillages || [];
    if (typeof vList === "string") vList = vList.split(",").map(function(s) { return s.trim(); });
    var inMyVillage = c.village && vList.indexOf(c.village) !== -1;
    var isAssignedToMe = String(c.assignedOfficerIds || "").indexOf(currentUser.userId) !== -1 ||
                         String(c.assignedDeputyId || "").indexOf(currentUser.userId) !== -1;
    return inMyVillage || isAssignedToMe;
  });

  if (!filter) return { success: true, data: accessible };
  var filtered = accessible.filter(function(c) {
    if (filter.village && filter.village !== "ALL" && c.village !== filter.village) return false;
    if (filter.status && filter.status !== "ALL" && c.status !== filter.status) return false;
    if (filter.keyword) {
      var kw = filter.keyword.toLowerCase();
      var text = (c.caseId + " " + c.description + " " + c.reporterName).toLowerCase();
      if (text.indexOf(kw) === -1) return false;
    }
    return true;
  });
  return { success: true, data: filtered };
}

function getCaseDetail(caseId, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var cases = getSheetData(CONFIG.SHEETS.CASES);
  var record = cases.find(function(c) { return String(c.caseId).trim().toUpperCase() === String(caseId).trim().toUpperCase(); });
  if (!record) return { success: false, message: "Không tìm thấy hồ sơ." };
  
  var histories = getSheetData(CONFIG.SHEETS.PROCESS_HISTORY).filter(function(h) { return String(h.caseId).trim().toUpperCase() === String(caseId).trim().toUpperCase(); });
  var attachments = getSheetData(CONFIG.SHEETS.ATTACHMENTS).filter(function(a) { return String(a.caseId).trim().toUpperCase() === String(caseId).trim().toUpperCase(); });
  return { success: true, data: { caseRecord: record, histories: histories, attachments: attachments } };
}

function updateCaseStatus(caseId, nextStatus, note, resolution, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var now = new Date().toISOString();
  var payload = { status: nextStatus, updatedAt: now };
  if (resolution) payload.resolutionNote = resolution;
  updateSheetRow(CONFIG.SHEETS.CASES, "caseId", caseId, payload);
  
  appendSheetRow(CONFIG.SHEETS.PROCESS_HISTORY, {
    historyId: "HIST-" + new Date().getTime(), caseId: caseId,
    action: "Cập nhật: " + nextStatus, toStatus: nextStatus,
    performedBy: currentUser.rank + " " + currentUser.fullName, performedAt: now, note: note || resolution || ""
  });
  return { success: true, message: "Cập nhật trạng thái thành công!" };
}

function deleteCaseRecordPermanent(caseId, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var r = String(currentUser.role || "").toUpperCase();
  if (r !== "SUPER_ADMIN" && r !== "CHIEF") {
    return { success: false, message: "Chỉ Trưởng CAX hoặc Quản trị viên mới có quyền xóa hồ sơ." };
  }
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CASES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(caseId)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Đã xóa hồ sơ " + caseId };
    }
  }
  return { success: false, message: "Không tìm thấy hồ sơ." };
}

/* =========================================================================
   4. QUẢN LÝ CÁN BỘ & ĐỊA BÀN PHỤ TRÁCH
   ========================================================================= */

function getAllOfficers(currentUser) {
  initSystemDatabase();
  var users = getSheetData(CONFIG.SHEETS.USERS);
  return { success: true, data: users };
}

function saveOfficerUser(officerData, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  
  var r = String(currentUser.role || "").toUpperCase();
  if (r !== "SUPER_ADMIN" && r !== "ADMIN" && r !== "CHIEF") {
    return { success: false, message: "Chỉ Quản trị viên hoặc Trưởng CAX mới có quyền cập nhật cán bộ." };
  }

  var vStr = "";
  if (Array.isArray(officerData.assignedVillages)) {
    vStr = officerData.assignedVillages.join(", ");
  } else {
    vStr = String(officerData.assignedVillages || "").trim();
  }

  if (officerData.userId) {
    updateSheetRow(CONFIG.SHEETS.USERS, "userId", officerData.userId, {
      fullName: officerData.fullName, badgeNumber: officerData.badgeNumber, rank: officerData.rank,
      position: officerData.position, role: officerData.role, phone: officerData.phone,
      assignedVillages: vStr
    });
    writeAuditLog(currentUser, "UPDATE_OFFICER", "USERS", officerData.userId, "Cập nhật cán bộ: " + officerData.fullName + " | Thôn: " + vStr);
    return { success: true, message: "Cập nhật thông tin cán bộ và địa bàn thành công!" };
  } else {
    var newId = "USR-" + (100 + Math.floor(Math.random() * 900));
    appendSheetRow(CONFIG.SHEETS.USERS, {
      userId: newId, username: officerData.username || ("cb_" + newId.toLowerCase()),
      passwordHash: "1", fullName: officerData.fullName, badgeNumber: officerData.badgeNumber,
      rank: officerData.rank, position: officerData.position, role: officerData.role || "PROCESSING_OFFICER",
      phone: officerData.phone, assignedVillages: vStr,
      status: "ACTIVE", createdAt: new Date().toISOString()
    });
    writeAuditLog(currentUser, "CREATE_OFFICER", "USERS", newId, "Thêm mới cán bộ: " + officerData.fullName + " | Thôn: " + vStr);
    return { success: true, message: "Thêm cán bộ mới thành công! (Mật khẩu khởi tạo: 1)" };
  }
}

function resetOfficerPassword(userId, currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var r = String(currentUser.role || "").toUpperCase();
  if (r !== "SUPER_ADMIN" && r !== "ADMIN" && r !== "CHIEF") {
    return { success: false, message: "Từ chối quyền reset mật khẩu." };
  }
  updateSheetRow(CONFIG.SHEETS.USERS, "userId", userId, { passwordHash: "1" });
  writeAuditLog(currentUser, "RESET_PASSWORD", "USERS", userId, "Đặt lại mật khẩu về mặc định (1)");
  return { success: true, message: "Đã reset mật khẩu cán bộ về '1'." };
}

function getAuditLogs(currentUser) {
  if (!currentUser) return { success: false, message: "Yêu cầu đăng nhập." };
  var logs = getSheetData(CONFIG.SHEETS.AUDIT_LOG);
  return { success: true, data: logs.reverse().slice(0, 100) };
}

function getDashboardStats(currentUser) {
  if (!currentUser) return { success: false, message: "Dữ liệu bảo mật nội bộ. Vui lòng đăng nhập." };
  var cases = getSheetData(CONFIG.SHEETS.CASES);
  var today = new Date().toISOString().slice(0, 10);
  var vStats = {};
  CONFIG.VILLAGES.forEach(function(v) { vStats[v] = 0; });
  var catStats = {};

  cases.forEach(function(c) {
    if (c.village && vStats[c.village] !== undefined) vStats[c.village]++;
    var cat = c.category || "Khác";
    catStats[cat] = (catStats[cat] || 0) + 1;
  });

  return {
    success: true,
    data: {
      total: cases.length,
      today: cases.filter(function(c) { return String(c.createdAt || '').slice(0, 10) === today; }).length,
      newCases: cases.filter(function(c) { return c.status === 'NEW'; }).length,
      processing: cases.filter(function(c) { return c.status === 'PROCESSING' || c.status === 'ASSIGNED'; }).length,
      resolved: cases.filter(function(c) { return c.status === 'RESOLVED' || c.status === 'CLOSED'; }).length,
      villageStats: vStats, categoryStats: catStats
    }
  };
}
