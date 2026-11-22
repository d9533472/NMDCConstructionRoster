/**
 * 施工組班表 — Apps Script 後端（綁定在「施工組班表」試算表）
 * 工作表：
 *   目前班表：每月一列（月份｜版本｜存檔時間｜存檔人｜備註｜JSON）
 *   歷史版本：每次存檔追加一列，可載回任一舊版
 *   YYYY-MM：每次存檔同步寫入一份可直接閱讀的班表
 */
const CUR = '目前班表';
const HIST = '歷史版本';
const HEAD = ['月份', '版本', '存檔時間', '存檔人', '備註', 'JSON'];

function setup() {
  sheet_(CUR); sheet_(HIST);
  const def = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('工作表1') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (def && SpreadsheetApp.getActiveSpreadsheet().getSheets().length > 1) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(def);
}

const TZ = 'Asia/Taipei';

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // 月份、存檔時間一律存成文字，避免試算表自動轉成日期
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('C:C').setNumberFormat('@');
  return sh;
}

function monthStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM') : String(v);
}

function isoStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString();
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
  return m ? new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':' + m[6] + '+08:00').toISOString() : String(v);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rowsFor_(sh, month) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getValues()
    .map((r, i) => ({ row: i + 2, month: monthStr_(r[0]), version: Number(r[1]), savedAt: isoStr_(r[2]), savedBy: r[3], note: r[4], json: r[5] }))
    .filter(r => r.month === month);
}

function current_(month) {
  return rowsFor_(sheet_(CUR), month)[0] || null;
}

function doGet(e) {
  const p = e.parameter || {};
  try {
    if (p.action === 'get') {
      const c = current_(p.month);
      return json_({ ok: true, found: !!c, version: c ? c.version : 0, savedAt: c ? c.savedAt : '', savedBy: c ? c.savedBy : '', note: c ? c.note : '', data: c ? JSON.parse(c.json) : null });
    }
    if (p.action === 'history') {
      const list = rowsFor_(sheet_(HIST), p.month).reverse().slice(0, 50)
        .map(r => ({ version: r.version, savedAt: r.savedAt, savedBy: r.savedBy, note: r.note }));
      return json_({ ok: true, list });
    }
    if (p.action === 'version') {
      const r = rowsFor_(sheet_(HIST), p.month).find(x => x.version === Number(p.version));
      return json_({ ok: !!r, data: r ? JSON.parse(r.json) : null, version: r ? r.version : 0, savedAt: r ? r.savedAt : '', savedBy: r ? r.savedBy : '' });
    }
    return json_({ ok: true, msg: '施工組班表 API' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'save') return json_({ ok: false, error: 'unknown action' });
    const month = String(b.month);
    if (!/^\d{4}-\d{2}$/.test(month)) return json_({ ok: false, error: 'bad month' });
    const cur = current_(month);
    const curV = cur ? cur.version : 0;
    // 有人在你打開之後先存過 → 不覆蓋，除非前端確認 force
    if (!b.force && Number(b.baseVersion || 0) !== curV) {
      return json_({ ok: false, conflict: true, version: curV, savedAt: cur.savedAt, savedBy: cur.savedBy });
    }
    const v = curV + 1;
    const now = new Date();
    const by = String(b.savedBy || '').slice(0, 40);
    const note = String(b.note || '').slice(0, 200);
    const js = JSON.stringify(b.data);
    const row = [month, v, Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), by, note, js];
    const cs = sheet_(CUR);
    if (cur) cs.getRange(cur.row, 1, 1, HEAD.length).setValues([row]);
    else cs.appendRow(row);
    sheet_(HIST).appendRow(row);
    if (Array.isArray(b.grid)) writeGrid_(month, b.grid, v, now, by);
    return json_({ ok: true, version: v, savedAt: now.toISOString(), savedBy: by });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 寫一份人看得懂的班表到「YYYY-MM」工作表 */
function writeGrid_(month, grid, v, now, by) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(month);
  if (!sh) sh = ss.insertSheet(month);
  sh.clear();
  const w = Math.max.apply(null, grid.map(r => r.length));
  const g = grid.map(r => r.concat(Array(w - r.length).fill('')));
  sh.getRange(1, 1).setValue('施工組班表 ' + month + '（v' + v + '，' + Utilities.formatDate(now, TZ, 'yyyy/MM/dd HH:mm') + (by ? '，' + by : '') + '）').setFontWeight('bold');
  sh.getRange(3, 1, g.length, w).setValues(g);
  sh.getRange(3, 1, 1, w).setFontWeight('bold');
  sh.setFrozenRows(3); sh.setFrozenColumns(1);
}
