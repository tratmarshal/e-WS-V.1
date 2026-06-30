// ========== gas.gs ==========
// Backend for Warrant Tracker — Google Sheets API v4
// ========== CONFIGURATION ==========

const WARRANT_DB_ID = "1gg0hX7IWMeyyYik9oFQphSJIJEzJaju2VT2DpLluc2E";
const UPDATE_DB_ID  = "18rpfC0MuEvb73ldDyuevqYECA9LClEY5wYCi7lZ87Go";

const SHEET_PROCESSING = "processing";

const PROCESSING_HEADERS = [
  "วัน-เวลา", "เลขที่หมายจับ", "ชื่อสกุล", "ประกัน",
  "เสนอท่าน", "สถานะสำนวน", "เหตุ", "รายละเอียดเหตุ",
  "สถานะหมายจับ", "ชุดจับ"
];

// ========== UTILITIES ==========

function nowText() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
}

function normalizeText_(value) {
  return String(value == null ? "" : value).trim();
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== COLUMN MAP for Warrant DB ==========

function getColumnMap(headers) {
  const map = {
    seq: 0, type: 1, warrantNo: 2, issuedDate: 3, fullName: 4,
    id13: 5, blackCaseNo: 6, redCaseNo: 7, charge: 8,
    addressNo: 9, moo: 10, tambon: 11, amphoe: 12, province: 13,
    limitation: 14, bail: 15, submitTo: 16, status: 17, note: 18
  };
  if (headers && headers.length > 0) {
    headers.forEach((h, i) => {
      const header = normalizeText_(h);
      if (header === "ลำดับ") map.seq = i;
      else if (header === "ประเภทหมายจับ" || header === "ประเภท") map.type = i;
      else if (header === "เลขที่หมายจับ") map.warrantNo = i;
      else if (header === "วันที่ออก") map.issuedDate = i;
      else if (header === "ชื่อสกุล" || header === "ชื่อ-สกุล") map.fullName = i;
      else if (header === "13 หลัก" || header === "เลขบัตรประชาชน" || header === "เลขประจำตัวประชาชน") map.id13 = i;
      else if (header === "เลขคดีดำ") map.blackCaseNo = i;
      else if (header === "เลขคดีแดง") map.redCaseNo = i;
      else if (header === "ความผิด") map.charge = i;
      else if (header === "บ้านเลขที่") map.addressNo = i;
      else if (header === "หมู่") map.moo = i;
      else if (header === "ตำบล") map.tambon = i;
      else if (header === "อำเภอ") map.amphoe = i;
      else if (header === "จังหวัด") map.province = i;
      else if (header === "อายุความ") map.limitation = i;
      else if (header === "ประกัน") map.bail = i;
      else if (header === "ท่าน" || header === "เสนอท่าน") map.submitTo = i;
      else if (header === "สถานะ") map.status = i;
      else if (header === "หมายเหตุ") map.note = i;
    });
  }
  return map;
}

// ========== SEARCH DB ==========

function searchDB(keyword, searchType) {
  const term = normalizeText_(keyword);
  if (!term) return jsonResponse_({ success: true, data: [] });

  try {
    const meta = Sheets.Spreadsheets.get(WARRANT_DB_ID);
    const allSheets = meta.sheets.map(s => s.properties.title);
    const ranges = allSheets.map(name => `${name}!A1:S`);
    const result = Sheets.Spreadsheets.Values.batchGet(WARRANT_DB_ID, { ranges: ranges });
    const allRows = [];

    result.valueRanges.forEach(vr => {
      if (!vr.values || vr.values.length < 2) return;
      const headers = vr.values[0];
      const colMap = getColumnMap(headers);
      const colD_Index = 5; // column F (13 หลัก) หรือดึงจาก colMap.id13 ด้านล่าง

      for (let r = 1; r < vr.values.length; r++) {
        const row = vr.values[r];
        const warrantNo = normalizeText_(row[colMap.warrantNo]);
        if (!warrantNo) continue;

        let match = false;
        if (searchType === "id13") {
          const colId13 = normalizeText_(row[colMap.id13] || "");
          match = (colId13 === term);
        } else {
          const name = normalizeText_(row[colMap.fullName] || "");
          match = name.toLowerCase().includes(term.toLowerCase());
        }

        if (match) {
          allRows.push({
            warrantNumber: normalizeText_(row[colMap.warrantNo] || ""),
            no: normalizeText_(row[colMap.warrantNo] || ""),
            defendantName: normalizeText_(row[colMap.fullName] || ""),
            fullName: normalizeText_(row[colMap.fullName] || ""),
            id13: normalizeText_(row[colMap.id13] || ""),
            charge: normalizeText_(row[colMap.charge] || ""),
            status: normalizeText_(row[colMap.status] || ""), // ดึงค่าจากคอลัมน์ O ล่าสุดส่งกลับให้ Frontend
            blackCaseNo: normalizeText_(row[colMap.blackCaseNo] || ""),
            redCaseNo: normalizeText_(row[colMap.redCaseNo] || ""),
            bail: normalizeText_(row[colMap.bail] || ""),
            submitTo: normalizeText_(row[colMap.submitTo] || "")
          });
        }
      }
    });

    return jsonResponse_({ success: true, data: allRows });

  } catch (err) {
    console.error("searchDB error:", err);
    return jsonResponse_({ success: false, error: err.message || "การค้นหาล้มเหลว" });
  }
}

// ========== SAVE DB ==========

function saveDB(action, payload) {
  try {
    if (action === "addProcess") {
      return addProcessRecord(payload);
    } else if (action === "markRevoked") {
      return markProcessRevoked(payload);
    } else if (action === "markForwarded") {
      return markProcessForwarded(payload);
    } else if (action === "report") {
      return markProcessReported(payload);
    } else {
      return { success: false, error: "ไม่พบ action: " + action };
    }
  } catch (err) {
    console.error("saveDB error:", err);
    return { success: false, error: err.message || "การบันทึกล้มเหลว" };
  }
}

// ฟังก์ชันดึงสถานะล่าสุดแบบเรียลไทม์จาก Database ต้นทาง (เพื่อความปลอดภัยในการเขียนข้อมูล)
function getLatestWarrantsStatusMap_(warrantNumbers) {
  const statusMap = {};
  if (!warrantNumbers || warrantNumbers.length === 0) return statusMap;

  const meta = Sheets.Spreadsheets.get(WARRANT_DB_ID);
  const allSheets = meta.sheets.map(s => s.properties.title);
  const ranges = allSheets.map(name => `${name}!A1:S`);
  const result = Sheets.Spreadsheets.Values.batchGet(WARRANT_DB_ID, { ranges: ranges });

  result.valueRanges.forEach(vr => {
    if (!vr.values || vr.values.length < 2) return;
    const headers = vr.values[0];
    const colMap = getColumnMap(headers);

    for (let r = 1; r < vr.values.length; r++) {
      const row = vr.values[r];
      const warrantNo = normalizeText_(row[colMap.warrantNo]);
      if (warrantNumbers.indexOf(warrantNo) !== -1) {
        statusMap[warrantNo] = normalizeText_(row[colMap.status] || "");
      }
    }
  });
  return statusMap;
}

function addProcessRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // รอคิวเพื่อเข้าบันทึกข้อมูล ป้องกันข้อมูลชนกัน
  try {
    const selectedWarrants = payload.selectedWarrants || [];
    if (!selectedWarrants.length) throw new Error("กรุณาเลือกหมายจับอย่างน้อย 1 รายการ");

    // 1. ดึงข้อมูลเลขที่หมายจับที่ถูกส่งมาบันทึก
    const warrantNumbersToCheck = selectedWarrants.map(w => normalizeText_(w.warrantNo || ""));

    // 2. ค้นหาสถานะปัจจุบัน (คอลัมน์ O) จากแหล่งข้อมูลหลักโดยตรง
    const latestStatusMap = getLatestWarrantsStatusMap_(warrantNumbersToCheck);

    // 3. ตรวจสอบสถานะ: หากมีหมายจับใดที่มีสถานะที่ไม่ใช่ "ว่าง" และไม่ใช่ "ต้องการตัว" ให้แจ้งข้อผิดพลาดกลับทันที
    for (let i = 0; i < warrantNumbersToCheck.length; i++) {
      const wNo = warrantNumbersToCheck[i];
      const currentStatus = latestStatusMap[wNo];
      
      // ตรวจสอบสถานะ หากไม่มีข้อมูลอยู่จริง หรือไม่ใช่สถานะที่อนุญาตให้บันทึกได้
      if (currentStatus === undefined) {
        throw new Error("ไม่พบเลขที่หมายจับ " + wNo + " ในฐานข้อมูลระบบ");
      }
      if (currentStatus !== "" && currentStatus !== "ต้องการตัว") {
        throw new Error("หมายจับเลขที่ " + wNo + " มีการเปลี่ยนแปลงสถานะเป็น '" + currentStatus + "' ไปแล้ว ไม่สามารถบันทึกซ้ำได้");
      }
    }

    // ตรวจสอบว่าชีต "processing" มีอยู่หรือไม่ ถ้าไม่มีให้สร้าง
    let nextRow = 2;
    try {
      const existing = Sheets.Spreadsheets.Values.get(UPDATE_DB_ID, `'${SHEET_PROCESSING}'!A:A`);
      nextRow = existing.values ? existing.values.length + 1 : 2;
    } catch (e) {
      // ชีตไม่มีอยู่ — สร้างชีตใหม่
      const spreadsheet = SpreadsheetApp.openById(UPDATE_DB_ID);
      spreadsheet.insertSheet(SHEET_PROCESSING);
      nextRow = 2;
    }

    const timestamp = nowText();
    const rows = selectedWarrants.map(w => [
      timestamp,
      normalizeText_(w.warrantNo || ""),
      normalizeText_(payload.defendantName || ""),
      "-",
      normalizeText_(payload.proposedTo || ""),
      "เสนอศาล",
      normalizeText_(payload.endReason || ""),
      normalizeText_(payload.reasonDetail || ""),
      "รอเพิกถอน",
      normalizeText_(payload.arrestTeam || "")
    ]);

    if (nextRow === 2) {
      try {
        Sheets.Spreadsheets.Values.update({
          values: [PROCESSING_HEADERS]
        }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!A1:J1`, {
          valueInputOption: "USER_ENTERED"
        });
      } catch (e) { /* ignore */ }
    }

    Sheets.Spreadsheets.Values.update({
      values: rows
    }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!A${nextRow}`, {
      valueInputOption: "USER_ENTERED"
    });

    return { success: true, added: rows.length, count: rows.length };
  } finally {
    lock.releaseLock(); // ปลดล็อกทุกกรณี
  }
}

function markProcessRevoked(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rowNum = parseInt(payload.rowId);
    if (!rowNum || rowNum < 2) throw new Error("ไม่พบรายการดำเนินการ");

    Sheets.Spreadsheets.Values.update({
      values: [["เพิกถอน"]]
    }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!I${rowNum}`, {
      valueInputOption: "USER_ENTERED"
    });

    Sheets.Spreadsheets.Values.update({
      values: [["รอส่งต่อสำนวน"]]
    }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!F${rowNum}`, {
      valueInputOption: "USER_ENTERED"
    });

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function markProcessForwarded(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rowNum = parseInt(payload.rowId);
    if (!rowNum || rowNum < 2) throw new Error("ไม่พบรายการดำเนินการ");

    Sheets.Spreadsheets.Values.update({
      values: [["ส่งต่อสำเร็จแล้ว"]]
    }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!F${rowNum}`, {
      valueInputOption: "USER_ENTERED"
    });

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function markProcessReported(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rowNum = parseInt(payload.rowId);
    if (!rowNum || rowNum < 2) throw new Error("ไม่พบรายการดำเนินการ");

    Sheets.Spreadsheets.Values.update({
      values: [["รายงานแล้ว"]]
    }, UPDATE_DB_ID, `'${SHEET_PROCESSING}'!F${rowNum}`, {
      valueInputOption: "USER_ENTERED"
    });

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ========== CALL LIST ==========

function callList() {
  try {
    const result = Sheets.Spreadsheets.Values.get(UPDATE_DB_ID, `'${SHEET_PROCESSING}'!A:J`);

    if (!result.values || result.values.length < 2) {
      return jsonResponse_({ success: true, data: [] });
    }

    const list = [];
    for (let r = 1; r < result.values.length; r++) {
      const row = result.values[r];
      if (!row[0]) continue;

      list.push({
        rowId: r + 1,
        timestamp: row[0] || "",
        warrantNumber: row[1] || "",
        defendantName: row[2] || "",
        bail: row[3] || "",
        proposedTo: row[4] || "",
        caseStatus: row[5] || "",
        endReason: row[6] || "",
        reasonDetail: row[7] || "",
        warrantStatus: row[8] || "",
        arrestTeam: row[9] || ""
      });
    }

    list.reverse();
    return jsonResponse_({ success: true, data: list });

  } catch (err) {
    console.error("callList error:", err);
    return jsonResponse_({ success: false, error: err.message || "ไม่สามารถโหลดข้อมูลได้", data: [] });
  }
}

// ========== API GATEWAY ==========

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ success: false, error: "No data received" });
    }

    const data = JSON.parse(e.postData.contents);
    const fn = data.fn || "";
    const payload = data.payload || {};

    switch (fn) {
      case "searchDB":
        return searchDB(payload.keyword, payload.searchType);
      case "saveDB":
        return jsonResponse_(saveDB(payload.action, payload.data || payload));
      case "callList":
        return callList();
      default:
        return jsonResponse_({ success: false, error: "Invalid function: " + fn });
    }
  } catch (err) {
    const message = err.message || String(err);
    console.error("doPost error:", message);
    return jsonResponse_({ success: false, error: message });
  }
}

function doGet() {
  return jsonResponse_({
    success: true,
    message: "Warrant Tracker API v3 — Sheets API v4",
    functions: ["searchDB", "saveDB", "callList"]
  });
}
