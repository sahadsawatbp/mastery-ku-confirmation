/**
 * API ตรวจสอบสถานะเอกสารจาก Google Sheet
 *
 * วิธีใช้: เปิด Google Sheet > ส่วนขยาย (Extensions) > Apps Script
 * แล้ววางโค้ดนี้ทั้งไฟล์ จากนั้นแก้ CONFIG ให้ตรงกับชีตของคุณ
 */

const CONFIG = {
  // ชื่อแท็บ (sheet) ที่เก็บข้อมูล
  SHEET_NAME: "Page_01",

  // แถวที่เป็นหัวคอลัมน์ (ปกติคือแถว 1)
  HEADER_ROW: 1,

  // คอลัมน์ที่ใช้ค้นหา (ชื่อต้องตรงกับหัวคอลัมน์ในชีตทุกตัวอักษร)
  SEARCH_COLUMNS: ["1.ชื่อ - นามสกุล", "2.อีเมล", "4.เลขประจำตัวประชาชน"],

  // คอลัมน์ที่อนุญาตให้ส่งไปแสดงบนเว็บ
  // คอลัมน์อื่นในชีต (เช่น เลขบัตรประชาชน ที่อยู่ หมายเหตุภายใน) จะไม่ถูกส่งออกไป
  // ชื่อเอกสารต้องตรงกับ DOCUMENTS[].field ใน index.html
  RETURN_COLUMNS: [
    "1.ชื่อ - นามสกุล",
    "TPQI",
    "สำเนาบัตรประชาชน",
    "สำเนาวุฒิการศึกษา",
    "หมายเหตุ"
  ],

  // ฟอร์มอาจถูกส่งซ้ำเพื่ออัปเดต ทำให้คนเดียวมีหลายแถว
  // ระบบจะนับจำนวนค่า TRUE (ผ่าน) ในคอลัมน์เอกสารเหล่านี้ แล้วแสดงแถวที่ได้คะแนนสูงสุด
  // ถ้าคะแนนเท่ากัน ใช้แถวล่าสุด (แถวล่างสุดในชีต)
  DOCUMENT_COLUMNS: ["TPQI", "สำเนาบัตรประชาชน", "สำเนาวุฒิการศึกษา"],

  // คอลัมน์ที่ใช้ระบุว่าแถวไหนเป็นคนเดียวกัน (ค่าเหมือนกัน = คนเดียวกัน)
  IDENTITY_COLUMN: "4.เลขประจำตัวประชาชน",

  // คอลัมน์ผลสรุปที่เจ้าหน้าที่เลือก (Dropdown: ผ่าน / ไม่ผ่าน / รอตรวจสอบ)
  // ชื่อหัวคอลัมน์ในชีตต้องตรงกับค่านี้ ถ้ายังไม่มีคอลัมน์นี้ ระบบจะข้ามไปโดยไม่ error
  STATUS_COLUMN: "สถานะ",

  // "exact"    = ต้องตรงทั้งค่า (แนะนำ ป้องกันคนเดาไล่ดูข้อมูลคนอื่น)
  // "contains" = มีคำที่พิมพ์อยู่ในค่าก็เจอ (เสี่ยงเปิดเผยข้อมูลมากกว่า)
  MATCH: "exact",

  MIN_QUERY_LENGTH: 3,
  MAX_RESULTS: 10
};

function doGet(e) {
  try {
    const q = String((e && e.parameter && e.parameter.q) || "").trim();
    if (q.length < CONFIG.MIN_QUERY_LENGTH) {
      return respond_({ ok: false, error: "คำค้นหาสั้นเกินไป" });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return respond_({ ok: false, error: "ไม่พบแท็บ " + CONFIG.SHEET_NAME });

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= CONFIG.HEADER_ROW) return respond_({ ok: true, results: [], updated: now_() });

    // getDisplayValues = ได้ค่าตามที่เห็นในชีต (วันที่/เบอร์โทรไม่เพี้ยน)
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0]
      .map(function (h) { return String(h).trim(); });
    const body = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastCol).getDisplayValues();

    const searchIdx = CONFIG.SEARCH_COLUMNS.map(function (c) { return headers.indexOf(c); }).filter(function (i) { return i >= 0; });
    const returnIdx = CONFIG.RETURN_COLUMNS.map(function (c) { return headers.indexOf(c); });
    const docIdx = CONFIG.DOCUMENT_COLUMNS.map(function (c) { return headers.indexOf(c); });
    const idIdx = headers.indexOf(CONFIG.IDENTITY_COLUMN);
    const statusIdx = headers.indexOf(CONFIG.STATUS_COLUMN);   // -1 = ยังไม่มีคอลัมน์ ไม่ถือว่าผิด

    if (searchIdx.length === 0) {
      return respond_({ ok: false, error: "ไม่พบคอลัมน์สำหรับค้นหา ตรวจสอบ SEARCH_COLUMNS" });
    }

    // ถ้าชื่อคอลัมน์ไม่ตรง (พิมพ์ผิด/เว้นวรรคเกิน) ต้องแจ้ง error
    // ไม่เช่นนั้นเว็บจะเห็นเอกสารเป็น "ยังไม่ได้รับ" ทั้งที่จริงแค่ชื่อไม่ตรง
    const missing = CONFIG.RETURN_COLUMNS.filter(function (c, k) { return returnIdx[k] < 0; })
      .concat(CONFIG.DOCUMENT_COLUMNS.filter(function (c, k) { return docIdx[k] < 0; }))
      .concat(idIdx < 0 ? [CONFIG.IDENTITY_COLUMN] : [])
      .filter(function (c, i, all) { return all.indexOf(c) === i; });
    if (missing.length) {
      return respond_({ ok: false, error: "ไม่พบคอลัมน์ในชีต: " + missing.join(", ") });
    }

    const nq = norm_(q);

    // 1) แถวที่ตรงกับคำค้น
    const matched = [];
    for (var r = 0; r < body.length; r++) {
      const row = body[r];
      const hit = searchIdx.some(function (i) {
        const v = norm_(row[i]);
        if (!v) return false;
        return CONFIG.MATCH === "contains" ? v.indexOf(nq) !== -1 : v === nq;
      });
      if (hit) matched.push(r);
    }

    // 2) รวมแถวของคนเดียวกัน (เลขบัตรประชาชนเดียวกัน)
    //    ค้นด้วยอีเมลเก่าก็ยังเห็นแถวที่ส่งอัปเดตใหม่ของคนเดียวกัน
    //    แถวที่ไม่มีเลขบัตรจะนับเป็นคนละคนกับแถวอื่น
    function keyOf(r) {
      const id = norm_(body[r][idIdx]);
      return id ? "id:" + id : "row:" + r;
    }
    const wanted = {};
    matched.forEach(function (r) { wanted[keyOf(r)] = true; });

    const groups = {};
    const keyOrder = [];
    for (var j = 0; j < body.length; j++) {
      const k = keyOf(j);
      if (!wanted[k]) continue;
      if (!groups[k]) { groups[k] = []; keyOrder.push(k); }
      groups[k].push(j);
    }

    // 3) แต่ละคนเลือกแถวที่มี TRUE มากที่สุด (เสมอกันเอาแถวล่าสุด)
    function score(r) {
      return docIdx.reduce(function (n, i) { return n + (isPass_(body[r][i]) ? 1 : 0); }, 0);
    }
    function best(rows) {
      var b = rows[0];
      rows.forEach(function (r) { if (score(r) >= score(b)) b = r; });  // rows เรียงบนลงล่าง จึงเอาแถวหลังสุดเมื่อเสมอ
      return b;
    }

    const results = keyOrder.slice(0, CONFIG.MAX_RESULTS).map(function (k) {
      const row = body[best(groups[k])];
      const item = {};
      CONFIG.RETURN_COLUMNS.forEach(function (name, n) {
        item[name] = row[returnIdx[n]];
      });
      if (statusIdx >= 0) item[CONFIG.STATUS_COLUMN] = row[statusIdx];
      return item;
    });

    return respond_({ ok: true, results: results, updated: now_() });
  } catch (err) {
    console.error(err);
    return respond_({ ok: false, error: "เกิดข้อผิดพลาดภายในระบบ" });
  }
}

/** รันฟังก์ชันนี้ใน Apps Script editor เพื่อทดสอบ แล้วดูผลที่ Execution log */
function testSearch() {
  Logger.log(doGet({ parameter: { q: "A001" } }).getContent());
}

/** นับเป็น "ผ่าน": TRUE / ✓ / ✔ หรือข้อความที่มีคำว่า ผ่าน (ไม่รวม ไม่ผ่าน) */
function isPass_(v) {
  const s = String(v).trim();
  if (!s) return false;
  if (s.toUpperCase() === "TRUE" || s === "✓" || s === "✔") return true;
  return s.indexOf("ผ่าน") !== -1 && s.indexOf("ไม่") === -1;
}

function norm_(s) {
  return String(s).toLowerCase().replace(/[\s\-]/g, "");
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d/M/yyyy HH:mm");
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}