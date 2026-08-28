/**
 * Backend สำหรับ "ใบรายงานการทำงานของเครื่องจักรกลช่วย"
 * - รับข้อมูลจากหน้าเว็บ (index.html) มาเก็บลง Google Sheet
 * - รองรับสถานะ "ร่าง" (บันทึกไว้ก่อน แก้ต่อข้ามเครื่องได้) และ "ส่งแล้ว"
 * - บันทึกซ้ำที่แถวเดิมด้วย id (ไม่สร้างข้อมูลซ้ำ)
 * - ส่งรายการกลับให้พนักงาน (โหลดร่าง) และเสมียน (พิมพ์เฉพาะที่ส่งแล้ว)
 *
 * วิธีติดตั้ง: ดูไฟล์ README-setup.md
 * หมายเหตุ: ถ้าเคยสร้างชีต "Reports" ด้วยเวอร์ชันเก่าไว้ ให้ลบแท็บนั้นทิ้งก่อน
 *           แล้ว Deploy ใหม่ ระบบจะสร้างหัวตารางชุดใหม่ให้อัตโนมัติ
 */

var SHEET_NAME = 'Reports';
var HEADERS = ['id', 'สถานะ', 'เวลาบันทึกร่าง', 'เวลาส่ง', 'วันที่',
  'เครื่องจักร No.', 'Code', 'เลขประจำตัวผู้ขับ', 'กะ', 'ประเภทงาน', 'ข้อมูลทั้งหมด (JSON)'];

/** POST: บันทึก/อัปเดตรายงาน (upsert ด้วย id) */
function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'save') return saveRecord_(p);
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function saveRecord_(p) {
  var sh = getSheet_();
  var d = p.data || {};
  var id = p.id || ('R' + Date.now());
  var status = p.status || 'ร่าง';
  var now = new Date();

  function rowVals(createdTs, submittedTs) {
    return [id, status, createdTs, submittedTs, d.date || '', d.mc_no || '', d.mc_code || '',
      d.drv_id || '', d.shift || '', d.wtype || '', JSON.stringify(d)];
  }

  var idx = findRowById_(sh, id);
  if (idx > 0) {
    var ex = sh.getRange(idx, 1, 1, HEADERS.length).getValues()[0];
    var createdTs = ex[2] || now;
    var submittedTs = ex[3];
    if (status === 'ส่งแล้ว' && !submittedTs) submittedTs = now;
    sh.getRange(idx, 1, 1, HEADERS.length).setValues([rowVals(createdTs, submittedTs)]);
  } else {
    sh.appendRow(rowVals(now, status === 'ส่งแล้ว' ? now : ''));
  }
  return json_({ ok: true, id: id, status: status });
}

function findRowById_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** GET: ?action=list (ส่งแล้ว) | drafts (ร่าง) | get&id=... */
function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'list') return json_(records_('ส่งแล้ว'));
  if (a === 'drafts') return json_(records_('ร่าง'));
  if (a === 'get') return json_(getById_(e.parameter.id) || { ok: false });
  return json_({ ok: true, msg: 'ITD machine report backend is running.' });
}

function records_(statusFilter) {
  var sh = getSheet_(), last = sh.getLastRow(), out = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (statusFilter && String(rows[i][1]) !== statusFilter) continue;
      var data = {};
      try { data = JSON.parse(rows[i][10]); } catch (x) {}
      out.push({ id: rows[i][0], status: rows[i][1], ts: rows[i][2], data: data });
    }
  }
  out.reverse(); // ใหม่สุดอยู่บน
  return out;
}

function getById_(id) {
  var sh = getSheet_(), idx = findRowById_(sh, id);
  if (idx < 0) return null;
  var r = sh.getRange(idx, 1, 1, HEADERS.length).getValues()[0];
  var data = {};
  try { data = JSON.parse(r[10]); } catch (x) {}
  return { id: r[0], status: r[1], data: data };
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
