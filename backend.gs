// ========== backend.gs ==========
// Consolidated Backend for Warrant Tracker — Optimized v2
// ========== CONFIGURATION ==========

const WARRANT_DB_ID = "1sq8GyDDiqS2U989TpcTlMyJ5P9gtrRTgbJjcQHtEOB4";
const UPDATE_DB_ID = "1mP3tver14Tp0wYd9xkk4Lh5Mv_g39UPy06juu8t_Dgk";

const SHEET_USERS = "users";
const SHEET_LOGS = "log";
const SHEET_PROCESSING = "processing";
const SHEET_WARRANT_INDEX = "warrant_index";

const WARRANT_SHEET_START = 55;
const WARRANT_SHEET_END = 69;

// ========== OPTIMIZED CACHE TTLs ==========
// TTL นานขึ้นสำหรับข้อมูลที่เปลี่ยนแปลงน้อย เพื่อลด I/O
const USER_CACHE_TTL_SECONDS = 3600;           // 1 ชม. (สิทธิ์ผู้ใช้เปลี่ยนน้อย)
const PENDING_PROCESS_CACHE_TTL_SECONDS = 30;   // 30 วินาที (ข้อมูลเปลี่ยนบ่อย ต้องสด)
const WARRANT_SEARCH_CACHE_TTL_SECONDS = 3600;  // 1 ชม. (หมายจับค้างเก่าไม่ค่อยเปลี่ยน)
const WARRANT_INDEX_CACHE_TTL_SECONDS = 3600;   // 1 ชม. (index เปลี่ยนแปลงเมื่อมี sync)
const LOCK_WAIT_MS = 30000;

// ========== CONSTANTS ==========

const USER_CACHE_KEY = "authorizedUsers:v2";
const PENDING_PROCESS_CACHE_KEY = "pendingProcess:v2";
const WARRANT_SEARCH_CACHE_KEY = "warrantSearch:v2";
const WARRANT_INDEX_CACHE_KEY = "warrantIndex:v2";

const WARRANT_STATUS_WANTED = "ต้องการตัว";
const WARRANT_STATUS_PENDING_REVOCATION = "สิ้นผลรอเพิกถอน";
const WARRANT_STATUS_REVOKED = "เพิกถอน";

const CASE_STATUS_PENDING = "รอดำเนินการ";
const CASE_STATUS_SUBMITTED_TO_COURT = "เสนอศาล";
const CASE_STATUS_WAITING_FORWARD = "รอส่งต่อ";
const CASE_STATUS_FORWARDED = "ส่งต่อแล้ว";

const PROCESSING_WARRANT_STATUS_PENDING_REVOCATION = "รอเพิกถอน";

const SYNC_STATUS_PENDING = "pending";
const SYNC_STATUS_SYNCED = "synced";
const SYNC_STATUS_ERROR = "error";

const REASON_OPTIONS = [
  "จำเลยมาศาลตามนัด", "จำเลยมาศาลและศาลอนุญาตให้ปล่อยชั่วคราว",
  "จับตัวจำเลยได้", "อายัดตัวจำเลยได้", "จับผู้ต้องหาได้",
  "ถอนหมายจับ", "ถอนฟ้อง/ถอนคำร้องทุกข์",
  "ขาดอายุความ", "จำเลยเสียชีวิต", "คดีถึงที่สุด", "อื่น ๆ"
];

const PROCESSING_HEADERS = [
  "วัน-เวลา", "LINEUSERID", "เลขที่หมายจับ", "ชื่อสกุล", "ประกัน",
  "เสนอท่าน", "สถานะสำนวน", "เหตุ", "รายละเอียดเหตุ", "สถานะหมายจับ",
  "syncStatus", "syncedAt", "syncError", "ลำดับรายการแก้ไข"
];

const LOG_HEADERS = [
  "วัน-เวลา", "LINEUSERID", "ชื่อ LINE", "หน้าเว็บ", "Action", "สถานะสิทธิ์", "รายละเอียด"
];

const WARRANT_INDEX_HEADERS = [
  "warrantNo", "sheetName", "rowNumber", "fullName", "id13", "status"
];

// ========== UTILITIES ==========

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT_MS);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function nowText() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
}

function normalizeText_(value) {
  return String(value == null ? "" : value).trim();
}

function indexOfHeader(headers, possibleNames, fallbackIndex) {
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeText_(headers[i]);
    for (let name of possibleNames) {
      if (h === name || h.indexOf(name) !== -1) return i;
    }
  }
  return fallbackIndex;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeFlush_() {
  // flush หลังเขียนข้อมูลสำคัญเพื่อให้แน่ใจว่าข้อมูลถูกบันทึกก่อนปล่อย Lock
  try { SpreadsheetApp.flush(); } catch (e) { console.warn("flush failed", e); }
}

// ========== DATABASE ACCESS ==========

function getWarrantDB() { return SpreadsheetApp.openById(WARRANT_DB_ID); }
function getUpdateDB() { return SpreadsheetApp.openById(UPDATE_DB_ID); }

function getOrCreateUpdateSheet_(name, headers) {
  const db = getUpdateDB();
  let sheet = db.getSheetByName(name);
  if (!sheet) sheet = db.insertSheet(name);
  if (headers && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function ensureProcessingSheet_() {
  const sheet = getOrCreateUpdateSheet_(SHEET_PROCESSING, PROCESSING_HEADERS);
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PROCESSING_HEADERS.length)).getValues()[0];
  const needsHeader = PROCESSING_HEADERS.some((h, i) => normalizeText_(current[i]) !== h);
  if (needsHeader) sheet.getRange(1, 1, 1, PROCESSING_HEADERS.length).setValues([PROCESSING_HEADERS]);
  return sheet;
}

function getWarrantColumnMap_(headers) {
  return {
    seq: indexOfHeader(headers, ["ลำดับ"], 0),
    type: indexOfHeader(headers, ["ประเภทหมายจับ"], 1),
    warrantNo: indexOfHeader(headers, ["เลขที่หมายจับ"], 2),
    issuedDate: indexOfHeader(headers, ["วันที่ออก"], 3),
    fullName: indexOfHeader(headers, ["ชื่อสกุล", "ชื่อ-สกุล"], 4),
    id13: indexOfHeader(headers, ["13 หลัก", "เลขบัตรประชาชน", "เลขประจำตัวประชาชน"], 5),
    blackCaseNo: indexOfHeader(headers, ["เลขคดีดำ"], 6),
    redCaseNo: indexOfHeader(headers, ["เลขคดีแดง"], 7),
    charge: indexOfHeader(headers, ["ความผิด"], 8),
    addressNo: indexOfHeader(headers, ["บ้านเลขที่"], 9),
    moo: indexOfHeader(headers, ["หมู่"], 10),
    tambon: indexOfHeader(headers, ["ตำบล"], 11),
    amphoe: indexOfHeader(headers, ["อำเภอ"], 12),
    province: indexOfHeader(headers, ["จังหวัด"], 13),
    limitation: indexOfHeader(headers, ["อายุความ"], 14),
    bail: indexOfHeader(headers, ["ประกัน"], 15),
    submitTo: indexOfHeader(headers, ["ท่าน"], 16),
    status: indexOfHeader(headers, ["สถานะ"], 17),
    note: indexOfHeader(headers, ["หมายเหตุ"], 18)
  };
}

// ========== WARRANT INDEX SHEET (Optimization #1) ==========
// Index Sheet ช่วยลดการค้นหาด้วย TextFinder ซึ่งช้ามาก
// มี column: warrantNo, sheetName, rowNumber, fullName, id13, status

function ensureWarrantIndex_() {
  const db = getUpdateDB();
  let sheet = db.getSheetByName(SHEET_WARRANT_INDEX);
  if (!sheet) {
    sheet = db.insertSheet(SHEET_WARRANT_INDEX);
    sheet.getRange(1, 1, 1, WARRANT_INDEX_HEADERS.length).setValues([WARRANT_INDEX_HEADERS]);
    rebuildWarrantIndex_();
  }
  return sheet;
}

function rebuildWarrantIndex_() {
  const cache = CacheService.getScriptCache();
  cache.remove(WARRANT_INDEX_CACHE_KEY);
  
  const db = getUpdateDB();
  let sheet = db.getSheetByName(SHEET_WARRANT_INDEX);
  if (!sheet) {
    sheet = db.insertSheet(SHEET_WARRANT_INDEX);
    sheet.getRange(1, 1, 1, WARRANT_INDEX_HEADERS.length).setValues([WARRANT_INDEX_HEADERS]);
  }
  
  // Clear old data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  
  const warrantDb = getWarrantDB();
  const allRows = [];
  
  for (let year = WARRANT_SHEET_START; year <= WARRANT_SHEET_END; year++) {
    const sh = warrantDb.getSheetByName(String(year));
    if (!sh || sh.getLastRow() < 2) continue;
    
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const cols = getWarrantColumnMap_(headers);
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    
    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const warrantNo = normalizeText_(row[cols.warrantNo]);
      if (!warrantNo) continue;
      
      allRows.push([
        warrantNo,
        String(year),
        r + 2,  // rowNumber = index + 2 (header offset)
        normalizeText_(row[cols.fullName]),
        normalizeText_(row[cols.id13]),
        normalizeText_(row[cols.status]) || WARRANT_STATUS_WANTED
      ]);
    }
  }
  
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, WARRANT_INDEX_HEADERS.length).setValues(allRows);
    safeFlush_();
  }
}

function findWarrantByNo_(warrantNo) {
  const target = normalizeText_(warrantNo);
  const matches = [];
  if (!target) return matches;
  
  // Try using index sheet first
  try {
    const indexSheet = getUpdateDB().getSheetByName(SHEET_WARRANT_INDEX);
    if (indexSheet && indexSheet.getLastRow() > 1) {
      const values = indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 6).getValues();
      for (let r = 0; r < values.length; r++) {
        if (normalizeText_(values[r][0]) === target) {
          const sheetName = normalizeText_(values[r][1]);
          const rowNumber = Number(values[r][2]);
          const db = getWarrantDB();
          const sheet = db.getSheetByName(sheetName);
          if (sheet && rowNumber > 1) {
            const lastCol = sheet.getLastColumn();
            const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
            const columns = getWarrantColumnMap_(headers);
            const rowValues = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
            matches.push({ sheet, rowNumber, row: rowValues, columns });
          }
          break;
        }
      }
      if (matches.length > 0) return matches;
    }
  } catch (e) {
    console.warn("Index sheet lookup failed, falling back to TextFinder", e);
  }
  
  // Fallback to original TextFinder method
  const db = getWarrantDB();
  const cells = db.createTextFinder(target).findAll();
  cells.forEach(cell => {
    const sheet = cell.getSheet();
    const sheetName = sheet.getName();
    const year = Number(sheetName);
    if (year >= WARRANT_SHEET_START && year <= WARRANT_SHEET_END) {
      const rowNum = cell.getRow();
      const colNum = cell.getColumn();
      if (rowNum > 1) {
        const lastCol = sheet.getLastColumn();
        if (lastCol < 1) return;
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const columns = getWarrantColumnMap_(headers);
        if (colNum === columns.warrantNo + 1) {
          const rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
          if (normalizeText_(rowValues[columns.warrantNo]) === target) {
            matches.push({ sheet, rowNumber: rowNum, row: rowValues, columns });
          }
        }
      }
    }
  });
  return matches;
}

function loadWarrantSearchFromSheets_(searchType, term) {
  const results = [];
  if (!term) return results;
  const lowerTerm = term.toLowerCase();
  
  // Try index sheet first — much faster than TextFinder
  try {
    const indexSheet = getUpdateDB().getSheetByName(SHEET_WARRANT_INDEX);
    if (indexSheet && indexSheet.getLastRow() > 1) {
      const values = indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 6).getValues();
      const db = getWarrantDB();
      
      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        const warrantNo = normalizeText_(row[0]);
        const sheetName = normalizeText_(row[1]);
        const rowNum = Number(row[2]);
        const fullName = normalizeText_(row[3]);
        const id13 = normalizeText_(row[4]);
        const status = normalizeText_(row[5]) || WARRANT_STATUS_WANTED;
        
        const match = searchType === "id13"
          ? id13 === term
          : fullName.toLowerCase().indexOf(lowerTerm) !== -1;
        
        if (match && rowNum > 1) {
          const sheet = db.getSheetByName(sheetName);
          if (sheet) {
            const lastCol = sheet.getLastColumn();
            const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
            const columns = getWarrantColumnMap_(headers);
            const rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
            
            results.push({
              sheetName: sheetName,
              rowNumber: rowNum,
              warrantNumber: warrantNo,
              no: warrantNo,
              defendantName: fullName,
              id13: id13,
              bail: normalizeText_(rowValues[columns.bail]) || "",
              submitTo: normalizeText_(rowValues[columns.submitTo]) || "",
              status: status,
              charge: normalizeText_(rowValues[columns.charge]) || "",
              blackCaseNo: normalizeText_(rowValues[columns.blackCaseNo]) || "",
              redCaseNo: normalizeText_(rowValues[columns.redCaseNo]) || "",
              fullName: fullName,
              warrantNo: warrantNo
            });
          }
        }
      }
      return results;
    }
  } catch (e) {
    console.warn("Index search failed, falling back to TextFinder", e);
  }
  
  // Fallback to TextFinder
  const db = getWarrantDB();
  const cells = db.createTextFinder(term).findAll();
  cells.forEach(cell => {
    const sheet = cell.getSheet();
    const sheetName = sheet.getName();
    const year = Number(sheetName);
    if (year >= WARRANT_SHEET_START && year <= WARRANT_SHEET_END) {
      const rowNum = cell.getRow();
      if (rowNum <= 1) return;
      const lastCol = sheet.getLastColumn();
      if (lastCol < 1) return;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const columns = getWarrantColumnMap_(headers);
      const colNum = cell.getColumn();
      const isTargetColumn = searchType === "id13" ? colNum === columns.id13 + 1 : colNum === columns.fullName + 1;
      if (isTargetColumn) {
        const row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        const rowId13 = normalizeText_(row[columns.id13]);
        const rowName = normalizeText_(row[columns.fullName]);
        const confirmed = searchType === "id13" ? rowId13 === term : rowName.toLowerCase().indexOf(lowerTerm) !== -1;
        if (confirmed) {
          results.push({
            sheetName, rowNumber: rowNum,
            warrantNumber: normalizeText_(row[columns.warrantNo]),
            no: normalizeText_(row[columns.warrantNo]),
            defendantName: rowName, id13: rowId13,
            bail: normalizeText_(row[columns.bail]),
            submitTo: normalizeText_(row[columns.submitTo]),
            status: normalizeText_(row[columns.status]) || WARRANT_STATUS_WANTED,
            charge: normalizeText_(row[columns.charge]),
            blackCaseNo: normalizeText_(row[columns.blackCaseNo]),
            redCaseNo: normalizeText_(row[columns.redCaseNo]),
            fullName: rowName, warrantNo: normalizeText_(row[columns.warrantNo])
          });
        }
      }
    }
  });
  return results;
}

// ========== CACHE LAYER ==========

function clearPendingProcessCache_() {
  try { CacheService.getScriptCache().remove(PENDING_PROCESS_CACHE_KEY); } catch (e) {}
}

function clearWarrantCache_() {
  try {
    CacheService.getScriptCache().remove(WARRANT_SEARCH_CACHE_KEY);
    CacheService.getScriptCache().remove(WARRANT_INDEX_CACHE_KEY);
  } catch (e) {}
}

function getCachedWarrantSearch_(searchType, term) {
  if (!term) return loadWarrantSearchFromSheets_(searchType, term);
  const cache = CacheService.getScriptCache();
  const cacheKey = WARRANT_SEARCH_CACHE_KEY + ":" + searchType + ":" + term;
  const cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (err) { cache.remove(cacheKey); } }
  const results = loadWarrantSearchFromSheets_(searchType, term);
  try { cache.put(cacheKey, JSON.stringify(results), WARRANT_SEARCH_CACHE_TTL_SECONDS); } catch (e) {}
  return results;
}

function getCachedPendingProcess_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(PENDING_PROCESS_CACHE_KEY);
    if (cached) { try { return JSON.parse(cached); } catch (err) { cache.remove(PENDING_PROCESS_CACHE_KEY); } }
    const results = loadPendingProcessFromSheets_();
    try { cache.put(PENDING_PROCESS_CACHE_KEY, JSON.stringify(results), PENDING_PROCESS_CACHE_TTL_SECONDS); } catch (e) {}
    return results;
  } catch (err) {
    console.error("getCachedPendingProcess_ failed", err);
    throw new Error("ไม่สามารถโหลดรายการรอดำเนินการได้: " + err.message);
  }
}

function loadPendingProcessFromSheets_() {
  try {
    const sheet = ensureProcessingSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, PROCESSING_HEADERS.length).getValues();
    const list = [];
    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      if (!row || !row[0]) continue;
      try {
        list.push({
          rowId: r + 2, timestamp: row[0], userId: row[1],
          warrantNumber: row[2], defendantName: row[3], bail: row[4],
          proposedTo: row[5], caseStatus: row[6], endReason: row[7],
          reasonDetail: row[8], warrantStatus: row[9],
          syncStatus: row[10] || SYNC_STATUS_PENDING,
          syncedAt: row[11], syncError: row[12], processSeq: row[13]
        });
      } catch (rowErr) { console.warn("Skipped row " + (r + 2) + ": " + rowErr.message); }
    }
    list.reverse();
    return list;
  } catch (err) {
    console.error("loadPendingProcessFromSheets_ failed", err);
    throw new Error("ล้มเหลวในการโหลดรายการรอดำเนินการ: " + err.message);
  }
}

// ========== AUTHENTICATION ==========

function getAuthorizedUserIds_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(USER_CACHE_KEY);
  if (cached) { try { return JSON.parse(cached); } catch (e) { cache.remove(USER_CACHE_KEY); } }
  const db = getUpdateDB();
  let sheet = db.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = db.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 1, 3).setValues([["LINEUSERID", "ชื่อ", "สถานะ"]]);
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const ids = values.map(r => normalizeText_(r[0])).filter(id => id);
  try { cache.put(USER_CACHE_KEY, JSON.stringify(ids), USER_CACHE_TTL_SECONDS); } catch (e) {}
  return ids;
}

function getAuthorizationStatus(userId) {
  const cleanUserId = normalizeText_(userId);
  if (!cleanUserId) return "unauthorized";
  return getAuthorizedUserIds_().indexOf(cleanUserId) !== -1 ? "authorized" : "unauthorized";
}

function clearUserCache_() {
  try { CacheService.getScriptCache().remove(USER_CACHE_KEY); } catch (e) {}
}

// ========== LOGGING (Optimization #2: Log only write operations) ==========

const WRITE_ACTIONS = ["addProcess", "markRevoked", "markForwarded", "addUser", "removeUser", "updateUser"];

function logActivity(userId, displayName, page, action, authStatus, detail) {
  // ✅ Log เฉพาะ Action ที่เขียนข้อมูล เพื่อลด I/O ที่ไม่จำเป็น
  if (WRITE_ACTIONS.indexOf(action) === -1) return;
  try {
    const sheet = getOrCreateUpdateSheet_(SHEET_LOGS, LOG_HEADERS);
    sheet.appendRow([nowText(), userId || "", displayName || "", page || "unknown", action || "", authStatus || "", detail || ""]);
  } catch (err) { console.error("logActivity failed", err); }
}

// ========== SEARCH ==========

function searchWarrant(query) {
  const term = normalizeText_(query);
  if (!term) return { success: true, data: [] };
  const searchType = (/^\d{13}$/.test(term)) ? "id13" : "name";
  return { success: true, data: getCachedWarrantSearch_(searchType, term) };
}

// ========== PROCESS RECORDS ==========

function validateReason_(reason, reasonDetail) {
  const cleanReason = normalizeText_(reason);
  if (REASON_OPTIONS.indexOf(cleanReason) === -1) throw new Error("เหตุไม่ถูกต้อง");
  if (cleanReason === "อื่น ๆ" && !normalizeText_(reasonDetail)) throw new Error("กรุณาระบุรายละเอียดเหตุ");
}

function getNextProcessSeq_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const seqColumn = PROCESSING_HEADERS.length;
  const seqValues = sheet.getRange(2, seqColumn, lastRow - 1, 1).getValues();
  const maxSeq = seqValues.reduce((max, row) => {
    const seq = Number(row[0]);
    return Number.isFinite(seq) && seq > max ? seq : max;
  }, 0);
  return maxSeq + 1;
}

function addProcessRecord(userId, payload) {
  return withScriptLock_(() => {
    const selectedWarrants = payload.selectedWarrants || [];
    if (!selectedWarrants.length) throw new Error("กรุณาเลือกหมายจับอย่างน้อย 1 รายการ");
    validateReason_(payload.endReason, payload.reasonDetail);
    const sheet = ensureProcessingSheet_();
    let nextSeq = getNextProcessSeq_(sheet);
    const seenWarrantNos = {};
    const rows = [];
    selectedWarrants.forEach(item => {
      const warrantNo = normalizeText_(typeof item === "string" ? item : item.warrantNo);
      if (!warrantNo) throw new Error("เลขที่หมายจับไม่ถูกต้อง");
      if (seenWarrantNos[warrantNo]) throw new Error("เลือกเลขที่หมายจับซ้ำ: " + warrantNo);
      seenWarrantNos[warrantNo] = true;
      rows.push([
        nowText(), userId, warrantNo,
        normalizeText_(item.defendantName || payload.defendantName || ""),
        normalizeText_(item.bail) || "-",
        normalizeText_(payload.proposedTo || item.submitTo || ""),
        CASE_STATUS_SUBMITTED_TO_COURT,
        normalizeText_(payload.endReason),
        normalizeText_(payload.reasonDetail),
        PROCESSING_WARRANT_STATUS_PENDING_REVOCATION,
        SYNC_STATUS_PENDING, "", "", nextSeq++
      ]);
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PROCESSING_HEADERS.length).setValues(rows);
    safeFlush_();  // ✅ flush ก่อนปล่อย Lock
    clearPendingProcessCache_();
    return { success: true, added: rows.length, count: rows.length, sync: { success: true, pending: true } };
  });
}

function getPendingProcess() {
  try {
    return { success: true, data: getCachedPendingProcess_() };
  } catch (err) {
    console.error("getPendingProcess error", err);
    return { success: false, error: err.message || "ไม่สามารถโหลดรายการรอดำเนินการได้", data: [] };
  }
}

function markProcessRevoked(rowId, userId) {
  return withScriptLock_(() => {
    const processSheet = ensureProcessingSheet_();
    const rowNumber = Number(rowId);
    if (rowNumber < 2 || rowNumber > processSheet.getLastRow()) throw new Error("ไม่พบรายการดำเนินการ");
    const row = processSheet.getRange(rowNumber, 1, 1, PROCESSING_HEADERS.length).getValues()[0];
    const warrantNo = normalizeText_(row[2]);
    const processStatus = mapProcessingWarrantStatus_(row[9]);
    if (processStatus === WARRANT_STATUS_REVOKED) throw new Error("หมายจับนี้ถูกบันทึกว่าเพิกถอนแล้ว");
    if (processStatus !== WARRANT_STATUS_PENDING_REVOCATION) throw new Error("เพิกถอนได้เฉพาะรายการที่อยู่สถานะสิ้นผลรอเพิกถอน");
    const matches = findWarrantByNo_(warrantNo);
    if (matches.length === 0) throw new Error("ไม่พบหมายจับเลขที่ " + warrantNo);
    if (matches.length > 1) throw new Error("พบเลขที่หมายจับซ้ำในฐานข้อมูล: " + warrantNo);
    const found = matches[0];
    const currentStatus = normalizeText_(found.row[found.columns.status]);
    if (currentStatus === WARRANT_STATUS_REVOKED) throw new Error("หมายจับนี้เพิกถอนแล้ว");
    if (currentStatus !== WARRANT_STATUS_PENDING_REVOCATION) throw new Error("สถานะปัจจุบันคือ " + currentStatus + " จึงเพิกถอนไม่ได้");
    
    // Update warrant database first (critical path)
    try { found.sheet.getRange(found.rowNumber, found.columns.status + 1).setValue(WARRANT_STATUS_REVOKED); } catch (err) { throw new Error("ล้มเหลวในการอัปเดตฐานข้อมูลหมายจับ: " + err.message); }
    // Update processing sheet (secondary path)
    try {
      processSheet.getRange(rowNumber, 7).setValue(CASE_STATUS_WAITING_FORWARD);
      processSheet.getRange(rowNumber, 10).setValue(WARRANT_STATUS_REVOKED);
      processSheet.getRange(rowNumber, 11).setValue(SYNC_STATUS_SYNCED);
      processSheet.getRange(rowNumber, 12).setValue(nowText());
      processSheet.getRange(rowNumber, 13).setValue("");
      safeFlush_();  // ✅ flush ก่อนปล่อย Lock
    } catch (err) {
      console.error("Warning: Warrant updated but processing sheet update failed: " + err.message);
      throw new Error("อัปเดตฐานข้อมูลหมายจับสำเร็จ แต่อัปเดตรายการดำเนินการล้มเหลว: " + err.message);
    }
    clearWarrantCache_();
    clearPendingProcessCache_();
    return { success: true };
  });
}

function markProcessForwarded(rowId, userId) {
  return withScriptLock_(() => {
    const processSheet = ensureProcessingSheet_();
    const rowNumber = Number(rowId);
    if (rowNumber < 2 || rowNumber > processSheet.getLastRow()) throw new Error("ไม่พบรายการดำเนินการ");
    const row = processSheet.getRange(rowNumber, 1, 1, PROCESSING_HEADERS.length).getValues()[0];
    const caseStatus = normalizeText_(row[6]);
    const processStatus = mapProcessingWarrantStatus_(row[9]);
    if (caseStatus === CASE_STATUS_FORWARDED) throw new Error("สำนวนนี้ถูกส่งต่อแล้ว");
    if (caseStatus !== CASE_STATUS_WAITING_FORWARD) throw new Error("ส่งต่อได้เฉพาะรายการที่อยู่สถานะรอส่งต่อ");
    if (processStatus !== WARRANT_STATUS_REVOKED) throw new Error("ส่งต่อได้เฉพาะรายการที่เพิกถอนแล้ว");
    processSheet.getRange(rowNumber, 7).setValue(CASE_STATUS_FORWARDED);
    safeFlush_();
    clearPendingProcessCache_();
    return { success: true };
  });
}

// ========== AUTO SYNC (Optimization #3: Batch update with setValues) ==========

function mapProcessingWarrantStatus_(status) {
  const cleanStatus = normalizeText_(status);
  if (cleanStatus === PROCESSING_WARRANT_STATUS_PENDING_REVOCATION) return WARRANT_STATUS_PENDING_REVOCATION;
  if (cleanStatus === WARRANT_STATUS_REVOKED) return WARRANT_STATUS_REVOKED;
  if (!cleanStatus) return WARRANT_STATUS_PENDING_REVOCATION;
  return cleanStatus;
}

function autoUpdateWarrantDatabase() {
  return withScriptLock_(() => autoUpdateWarrantDatabase_());
}

function autoUpdateWarrantDatabase_() {
  const processSheet = ensureProcessingSheet_();
  const lastRow = processSheet.getLastRow();
  if (lastRow < 2) return { success: true, synced: 0, errors: 0 };
  
  const values = processSheet.getRange(2, 1, lastRow - 1, PROCESSING_HEADERS.length).getValues();
  
  // ✅ Batch: รวบรวมข้อมูล warrantNo ที่ต้อง sync ก่อน แล้วค่อยอัปเดตทีเดียว
  // ช่วยลดจำนวนครั้งที่เรียก findWarrantByNo_ (ซึ่งใช้ TextFinder)
  const batchSync = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const warrantNo = normalizeText_(row[2]);
    const targetStatus = mapProcessingWarrantStatus_(row[9]);
    const syncStatus = normalizeText_(row[10]);
    if (warrantNo && syncStatus !== SYNC_STATUS_SYNCED && targetStatus === WARRANT_STATUS_PENDING_REVOCATION) {
      batchSync.push({ index: i, rowNumber: i + 2, warrantNo: warrantNo, targetStatus: targetStatus });
    }
  }
  
  if (batchSync.length === 0) return { success: true, synced: 0, errors: 0 };
  
  // ✅ Batch lookup: รวบรวมรายการที่ต้องอัปเดตก่อนทั้งหมด (ลด TextFinder)
  const updates = [];
  const errors_list = [];
  
  for (const item of batchSync) {
    try {
      const matches = findWarrantByNo_(item.warrantNo);
      if (matches.length === 0) throw new Error("ไม่พบหมายจับเลขที่ " + item.warrantNo);
      if (matches.length > 1) throw new Error("พบเลขที่หมายจับซ้ำ: " + item.warrantNo);
      const found = matches[0];
      const currentStatus = normalizeText_(found.row[found.columns.status]) || WARRANT_STATUS_WANTED;
      
      if (currentStatus === WARRANT_STATUS_REVOKED) throw new Error("หมายจับเพิกถอนแล้ว");
      if (currentStatus !== WARRANT_STATUS_WANTED && currentStatus !== WARRANT_STATUS_PENDING_REVOCATION) {
        throw new Error("พบสถานะที่ไม่รู้จัก: " + currentStatus);
      }
      
      const note = currentStatus === WARRANT_STATUS_PENDING_REVOCATION ? "สถานะเป็นสิ้นผลรอเพิกถอนอยู่แล้ว" : "";
      
      updates.push({
        found: found,
        rowNumber: item.rowNumber,
        shouldUpdateWarrant: currentStatus === WARRANT_STATUS_WANTED,
        note: note
      });
    } catch (err) {
      errors_list.push({ rowNumber: item.rowNumber, error: err.message });
    }
  }
  
  // ✅ Batch apply: อัปเดตทีเดียว
  for (const upd of updates) {
    if (upd.shouldUpdateWarrant) {
      upd.found.sheet.getRange(upd.found.rowNumber, upd.found.columns.status + 1).setValue(WARRANT_STATUS_PENDING_REVOCATION);
    }
    processSheet.getRange(upd.rowNumber, 11, 1, 3).setValues([[SYNC_STATUS_SYNCED, nowText(), upd.note]]);
  }
  
  for (const err of errors_list) {
    processSheet.getRange(err.rowNumber, 11, 1, 3).setValues([[SYNC_STATUS_ERROR, nowText(), err.error]]);
  }
  
  safeFlush_();
  clearWarrantCache_();
  
  return {
    success: errors_list.length === 0,
    synced: updates.length,
    errors: errors_list.length
  };
}

// ========== USER MANAGEMENT ==========

function getAllAuthorizedUsers() {
  return withScriptLock_(() => {
    const db = getUpdateDB();
    let sheet = db.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, users: [] };
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    return { success: true, users: values.map(row => ({ userId: normalizeText_(row[0]), displayName: normalizeText_(row[1]), status: normalizeText_(row[2]) })).filter(u => u.userId) };
  });
}

function addAuthorizedUser(userId, displayName) {
  return withScriptLock_(() => {
    const cleanUserId = normalizeText_(userId);
    const cleanName = normalizeText_(displayName);
    if (!cleanUserId) throw new Error("ต้องระบุ LINEUSERID");
    if (!cleanName) throw new Error("ต้องระบุชื่อแสดงผล");
    const db = getUpdateDB();
    let sheet = db.getSheetByName(SHEET_USERS);
    if (!sheet) { sheet = db.insertSheet(SHEET_USERS); sheet.getRange(1, 1, 1, 3).setValues([["LINEUSERID", "ชื่อ", "สถานะ"]]); }
    if (sheet.getLastRow() >= 2) {
      const exists = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().some(r => normalizeText_(r[0]) === cleanUserId);
      if (exists) throw new Error("LINEUSERID " + cleanUserId + " มีอยู่แล้ว");
    }
    sheet.appendRow([cleanUserId, cleanName, "active"]);
    clearUserCache_();
    return { success: true, message: "เพิ่มเจ้าหน้าที่ " + cleanName + " สำเร็จ" };
  });
}

function updateAuthorizedUser(userId, displayName, status) {
  return withScriptLock_(() => {
    const cleanUserId = normalizeText_(userId);
    if (!cleanUserId) throw new Error("ต้องระบุ LINEUSERID");
    const db = getUpdateDB();
    let sheet = db.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error("ไม่พบชีตผู้ใช้");
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    let foundRow = -1;
    for (let i = 0; i < values.length; i++) { if (normalizeText_(values[i][0]) === cleanUserId) { foundRow = i + 2; break; } }
    if (foundRow === -1) throw new Error("ไม่พบ LINEUSERID " + cleanUserId);
    if (displayName) sheet.getRange(foundRow, 2).setValue(normalizeText_(displayName));
    if (status) sheet.getRange(foundRow, 3).setValue(normalizeText_(status));
    clearUserCache_();
    return { success: true, message: "แก้ไขข้อมูลเจ้าหน้าที่ " + cleanUserId + " สำเร็จ" };
  });
}

function removeAuthorizedUser(userId) {
  return withScriptLock_(() => {
    const cleanUserId = normalizeText_(userId);
    if (!cleanUserId) throw new Error("ต้องระบุ LINEUSERID");
    const db = getUpdateDB();
    let sheet = db.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) throw new Error("ไม่พบชีตผู้ใช้");
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    let foundRow = -1;
    for (let i = 0; i < values.length; i++) { if (normalizeText_(values[i][0]) === cleanUserId) { foundRow = i + 2; break; } }
    if (foundRow === -1) throw new Error("ไม่พบ LINEUSERID " + cleanUserId);
    sheet.deleteRow(foundRow);
    clearUserCache_();
    return { success: true, message: "ลบเจ้าหน้าที่ " + cleanUserId + " สำเร็จ" };
  });
}

function getUserStatus(userId) {
  const cleanUserId = normalizeText_(userId);
  const authStatus = getAuthorizationStatus(cleanUserId);
  return { success: true, userId: cleanUserId, authorized: authStatus === "authorized", status: authStatus };
}

// ========== DIAGNOSTICS ==========

function diagnoseSystem() {
  const report = { timestamp: nowText(), databases: {}, sheets: {}, indexStatus: "", errors: [] };
  try {
    const wdb = getWarrantDB();
    report.databases.warrant = { id: WARRANT_DB_ID, accessible: true, sheets: wdb.getSheets().map(s => ({ name: s.getName(), rows: s.getLastRow() })) };
  } catch (err) { report.errors.push("Warrant DB: " + err.message); report.databases.warrant = { accessible: false, error: err.message }; }
  try {
    const udb = getUpdateDB();
    report.databases.update = { id: UPDATE_DB_ID, accessible: true, sheets: udb.getSheets().map(s => ({ name: s.getName(), rows: s.getLastRow() })) };
  } catch (err) { report.errors.push("Update DB: " + err.message); report.databases.update = { accessible: false, error: err.message }; }
  try {
    const sheet = ensureProcessingSheet_();
    report.sheets.processing = { exists: true, lastRow: sheet.getLastRow(), headers: sheet.getRange(1, 1, 1, PROCESSING_HEADERS.length).getValues()[0] };
  } catch (err) { report.errors.push("Processing: " + err.message); }
  try {
    const idxSheet = getUpdateDB().getSheetByName(SHEET_WARRANT_INDEX);
    report.indexStatus = idxSheet ? "exists (" + (idxSheet.getLastRow() - 1) + " entries)" : "not found";
  } catch (err) { report.indexStatus = "error: " + err.message; }
  return report;
}

function rebuildIndex() {
  // เรียก API นี้เพื่อ rebuild index sheet เมื่อข้อมูลหมายจับเปลี่ยน
  return withScriptLock_(() => {
    rebuildWarrantIndex_();
    clearWarrantCache_();
    return { success: true, message: " rebuilt and caches cleared" };
  });
}

// ========== API GATEWAY ==========

function doPost(e) {
  let action = "unknown";
  let userId = "";
  let displayName = "";
  let page = "unknown";
  let authStatus = "unauthorized";

  try {
    if (!e || !e.postData || !e.postData.contents) return jsonResponse_({ success: false, error: "No data received" });
    const data = JSON.parse(e.postData.contents);
    action = data.action || "unknown";
    userId = data.userId || "";
    displayName = data.displayName || "";
    page = data.page || "unknown";
    const payload = data.payload || data;
    authStatus = getAuthorizationStatus(userId);
    logActivity(userId, displayName, page, action, authStatus, "request");
    if (authStatus !== "authorized") return jsonResponse_({ success: false, error: "Unauthorized user" });

    let result;
    switch (action) {
      case "search":         result = searchWarrant(payload.keyword); break;
      case "getPending":     result = getPendingProcess(); break;
      case "addProcess":     result = addProcessRecord(userId, payload); break;
      case "syncProcessing": result = autoUpdateWarrantDatabase(); break;
      case "markRevoked":    result = markProcessRevoked(payload.rowId, userId); break;
      case "markForwarded":  result = markProcessForwarded(payload.rowId, userId); break;
      case "addUser":        result = addAuthorizedUser(payload.userId, payload.displayName); break;
      case "removeUser":     result = removeAuthorizedUser(payload.userId); break;
      case "updateUser":     result = updateAuthorizedUser(payload.userId, payload.displayName, payload.status); break;
      case "getUsers":       result = getAllAuthorizedUsers(); break;
      case "getUserStatus":  result = getUserStatus(userId); break;
      case "diagnose":       result = diagnoseSystem(); break;
      case "rebuildIndex":   result = rebuildIndex(); break;
      default:               result = { success: false, error: "Invalid action: " + action };
    }
    logActivity(userId, displayName, page, action, authStatus, result.success === false ? (result.error || "failed") : "success");
    return jsonResponse_(result);
  } catch (err) {
    const message = err.message || String(err);
    logActivity(userId, displayName, page, action, authStatus, message);
    return jsonResponse_({ success: false, error: message });
  }
}

function doGet() {
  return jsonResponse_({
    success: true,
    message: "Warrant Tracker API v2 — Optimized",
    version: "2.0",
    endpoints: [
      "search, getPending, addProcess, syncProcessing",
      "markRevoked, markForwarded, addUser, removeUser",
      "updateUser, getUsers, getUserStatus, diagnose, rebuildIndex"
    ]
  });
}