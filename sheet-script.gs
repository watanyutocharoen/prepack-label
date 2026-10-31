/* Prepack Label — ตัวรับ-ส่งข้อมูลกับ Google Sheet
   วางใน Extensions > Apps Script แล้ว Deploy > New deployment > Web app
   (Execute as: Me / Who has access: Anyone) */
const SECRET = '';   // ถ้าอยากตั้งรหัสลับ ใส่ตรงนี้ แล้วใส่รหัสเดียวกันในช่อง "รหัสลับ" ของแอป

const H_DRUG = ['รหัสยา', 'ชื่อยา'];
const H_LIST = ['ประเภท', 'ค่า'];
const H_HIST = ['id', 'upd', 'วันแบ่งบรรจุ', 'รหัสยา', 'ชื่อยา', 'Lot', 'หมดอายุ', 'จำนวนต่อแพ็ค', 'ผู้แบ่งบรรจุ', 'ผู้ทำสติ๊กเกอร์', 'ผู้ติดสติ๊กเกอร์', 'เภสัชกรตรวจสอบ', 'หน่วยที่แพ็ค', 'json'];
const KINDS = ['units', 'assts', 'pharms'];

function sheet_(name, hdr) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(hdr); s.setFrozenRows(1); }
  return s;
}
function rows_(s, n) {
  const last = s.getLastRow();
  return last < 2 ? [] : s.getRange(2, 1, last - 1, n).getValues();
}
function put_(s, hdr, data) {
  const last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, s.getMaxColumns()).clearContent();
  if (data.length) s.getRange(2, 1, data.length, hdr.length).setValues(data);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function readAll_() {
  const drugs = rows_(sheet_('ยา', H_DRUG), 2).map(r => ({ code: String(r[0]), name: String(r[1]) }));
  const lists = { units: [], assts: [], pharms: [] };
  rows_(sheet_('รายชื่อ', H_LIST), 2).forEach(r => { if (lists[r[0]]) lists[r[0]].push(String(r[1])); });
  const history = rows_(sheet_('ประวัติ', H_HIST), H_HIST.length).map(r => {
    try { const o = JSON.parse(r[13]); o.upd = Number(r[1]) || 0; return o; } catch (e) { return null; }
  }).filter(Boolean);
  return { ok: true, drugs: drugs, units: lists.units, assts: lists.assts, pharms: lists.pharms, history: history, hasData: drugs.length > 0 || history.length > 0 };
}
function histRow_(h) {
  return [h.id, h.upd || 0, h.pack || '', h.code || '', h.name || '', h.lot || '', h.exp || '',
    (h.qty || '') + ' ' + (h.unit || ''), h.maker || '', h.asst || '', h.sticker || '', h.verBy || '', h.site || '', JSON.stringify(h)];
}

function doGet(e) {
  if (SECRET && (!e || !e.parameter || e.parameter.k !== SECRET)) return out_({ ok: false, err: 'key' });
  return out_({ ok: true, msg: 'Prepack Label sheet พร้อมใช้งาน' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    if (SECRET && b.k !== SECRET) return out_({ ok: false, err: 'key' });

    if (b.drugs) put_(sheet_('ยา', H_DRUG), H_DRUG, b.drugs.map(d => [String(d.code || ''), String(d.name || '')]));

    const L = b.lists || {};
    if (KINDS.some(k => L[k])) {
      const cur = readAll_();
      const data = [];
      KINDS.forEach(k => ((L[k] && L[k].length) ? L[k] : cur[k]).forEach(v => data.push([k, v])));
      put_(sheet_('รายชื่อ', H_LIST), H_LIST, data);
    }

    if ((b.up && b.up.length) || (b.del && b.del.length)) {
      const s = sheet_('ประวัติ', H_HIST);
      const map = {};
      rows_(s, H_HIST.length).forEach(r => { map[r[0]] = r; });
      (b.up || []).forEach(h => {
        const old = map[h.id];
        if (!old || (Number(old[1]) || 0) <= (h.upd || 0)) map[h.id] = histRow_(h);
      });
      (b.del || []).forEach(id => { delete map[id]; });
      const all = Object.keys(map).map(k => map[k]);
      all.sort((a, c) => (JSON.parse(a[13]).createdAt || 0) - (JSON.parse(c[13]).createdAt || 0));
      put_(s, H_HIST, all);
    }
    return out_(readAll_());
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally {
    lock.releaseLock();
  }
}
