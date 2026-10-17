/**
 * Crystal.shop — الخلفية (Google Apps Script) — النسخة 2
 * جديد: حفظ المنتجات والإعدادات + رفع صور المنتجات إلى Google Drive
 * ملاحظة: بعد أي تعديل على الكود اعمل Deploy > Manage deployments > Edit > New version > Deploy
 */
const ADMIN_KEY = 'ضع-كلمة-السر-هنا';

const ORDER_HEAD = ['id', 'time', 'status', 'model', 'qty', 'total', 'name', 'phone', 'city', 'address', 'source', 'note'];
const VISIT_HEAD = ['time', 'vid', 'source', 'device', 'page'];
const IMG_FOLDER = 'Crystal Shop Images';

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* شغّل هذه الدالة مرة واحدة من المحرر لمنح صلاحية Google Drive (لرفع الصور) */
function authorize_() {
  DriveApp.getRootFolder();
  return 'ok';
}

/* ===== الإعدادات والمنتجات ===== */
function getConfig_() {
  const sh = sheet_('Config', ['json']);
  if (sh.getLastRow() < 2) return null;
  try { return JSON.parse(sh.getRange(2, 1).getValue()); } catch (e) { return null; }
}

function saveConfig_(cfg) {
  if (!cfg || !Array.isArray(cfg.products)) throw new Error('bad config');
  const s = JSON.stringify(cfg);
  if (s.length > 45000) throw new Error('config too large');
  const sh = sheet_('Config', ['json']);
  sh.getRange(2, 1).setValue(s);
}

function uploadImage_(b) {
  const folders = DriveApp.getFoldersByName(IMG_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(IMG_FOLDER);
  const mime = /^image\/(jpeg|png|webp)$/.test(b.mime) ? b.mime : 'image/jpeg';
  const ext = mime === 'image/png' ? 'png' : mime === 'image/webp' ? 'webp' : 'jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b.data), mime, 'p-' + Date.now() + '.' + ext);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

/* طلبات المسؤول عبر POST (نص عادي JSON) */
function adminPost_(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });
    if (b.action === 'saveConfig') { saveConfig_(b.config); return json_({ ok: true }); }
    if (b.action === 'uploadImage') { return json_({ ok: true, url: uploadImage_(b) }); }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* استقبال البيانات من المتجر (POST) */
function doPost(e) {
  if (e.postData && /^text\/plain/.test(e.postData.type || '')) return adminPost_(e);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = e.parameter || {};
    if (p.type === 'visit') {
      sheet_('Visits', VISIT_HEAD).appendRow([new Date(), p.vid || '', p.source || 'direct', p.device || '', p.page || '']);
      return json_({ ok: true });
    }
    const sh = sheet_('Orders', ORDER_HEAD);
    const last = sh.getLastRow();
    let row = -1;
    if (last > 1 && p.id) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat();
      const i = ids.indexOf(p.id);
      if (i > -1) row = i + 2;
    }
    const vals = [p.id || '', new Date(), p.status || '', p.model || '', Number(p.qty) || 1, Number(p.total) || 0,
                  p.name || '', "'" + (p.phone || ''), p.city || '', p.address || '', p.source || '', ''];
    if (row > -1) {
      const old = sh.getRange(row, 1, 1, ORDER_HEAD.length).getValues()[0];
      vals[1] = old[1];
      vals[11] = old[11];
      sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    } else {
      sh.appendRow(vals);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = e.parameter || {};

  // عام: المتجر يقرأ منه المنتجات (لا يحتوي بيانات حساسة)
  if (p.action === 'products') return json_({ ok: true, config: getConfig_() });

  if (p.key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });

  if (p.action === 'setStatus' || p.action === 'setNote') {
    const sh = sheet_('Orders', ORDER_HEAD);
    const last = sh.getLastRow();
    if (last < 2) return json_({ ok: false, error: 'empty' });
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat();
    const i = ids.indexOf(p.id);
    if (i < 0) return json_({ ok: false, error: 'not found' });
    if (p.action === 'setStatus') sh.getRange(i + 2, 3).setValue(p.status);
    else sh.getRange(i + 2, 12).setValue(p.note || '');
    return json_({ ok: true });
  }

  const osh = sheet_('Orders', ORDER_HEAD), vsh = sheet_('Visits', VISIT_HEAD);
  const ol = osh.getLastRow(), vl = vsh.getLastRow();
  const orders = ol > 1 ? osh.getRange(2, 1, ol - 1, ORDER_HEAD.length).getValues().map(r => ({
    id: r[0], time: new Date(r[1]).getTime(), status: r[2], model: r[3], qty: r[4], total: r[5],
    name: r[6], phone: String(r[7]).replace(/^'/, ''), city: r[8], address: r[9], source: r[10], note: r[11]
  })) : [];
  const from = Math.max(2, vl - 30000 + 1);
  const visits = vl > 1 ? vsh.getRange(from, 1, vl - from + 1, 4).getValues().map(r => [new Date(r[0]).getTime(), r[1], r[2], r[3]]) : [];
  return json_({ ok: true, orders: orders, visits: visits });
}
