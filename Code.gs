/* =====================================================================
   Crystal.shop — الخلفية: الطلبات + الزوار + المنتجات + مساحة البائع
   الصق هذا الكود كما هو في Apps Script ثم شغّل الدالة setup مرة واحدة.
   لا تحتاج لتعديل أي شيء هنا.
   ===================================================================== */

/* ★★ غيّر كلمة السر هذه قبل أي شيء — بها تدخل إلى «مساحة البائع» (admin.html) ★★ */
const ADMIN_PASSWORD = 'CHANGE_ME';

const SH = { DASH: 'لوحة التحكم', PRODUCTS: 'المنتجات', ORDERS: 'الطلبات', VISITS: 'الزيارات', EVENTS: 'الأحداث' };
const PRODUCT_HEADERS = ['المعرّف (اختياري)', 'اسم المنتج', 'القسم', 'السعر', 'الصور', 'وصف قصير', 'المميزات', 'عنوان الأنواع', 'الأنواع', 'شارة', 'متوفر'];
const ORDER_HEADERS = ['التاريخ', 'رقم الطلب', 'الحالة', 'الاسم', 'الهاتف', 'المدينة', 'العنوان', 'المنتجات', 'المجموع', 'المصدر', 'ملاحظاتي'];
const VISIT_HEADERS = ['التاريخ', 'معرّف الزائر', 'المصدر', 'الجهاز', 'زائر جديد'];
const EVENT_HEADERS = ['التاريخ', 'الحدث', 'المنتج', 'معرّف الزائر'];
const STATUSES = ['لم يؤكد بعد', 'جديد', 'تم التأكيد', 'تم الشحن', 'تم التسليم', 'ملغي'];

/* ---------- استقبال البيانات من المتجر ---------- */

function doPost(e) {
  const pd = e && e.postData;
  if (pd && String(pd.type || '').indexOf('text/plain') === 0) return adminPost_(pd.contents);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = (e && e.parameter) || {};
    const type = p.type || 'order';
    if (type === 'visit') logVisit_(p);
    else if (type === 'event') logEvent_(p);
    else saveOrder_(p);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const a = p.action;
  try {
    if (a === 'content') return publicContent_();
    if (a === 'products') return json_({ ok: true, products: getProducts_() });
    if (a === 'admin' || a === 'setStatus' || a === 'setNote') {
      if (ADMIN_PASSWORD === 'CHANGE_ME') return json_({ ok: false, error: 'setpw' });
      if (String(p.pw || '') !== ADMIN_PASSWORD) return json_({ ok: false, error: 'auth' });
      if (a === 'setStatus') updateOrder_(p.id, 3, STATUSES.indexOf(p.status) >= 0 ? p.status : null);
      if (a === 'setNote') updateOrder_(p.id, 11, clean_(p.note, 500));
      return json_({ ok: true, data: adminData_() });
    }
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  return json_({ ok: true, name: 'Crystal.shop' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

/* منع تنفيذ الصيغ من نصوص الزبائن */
function clean_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/* ---------- الطلبات ---------- */

function saveOrder_(p) {
  const id = clean_(p.id, 40);
  if (!id) return;
  const sh = book_().getSheetByName(SH.ORDERS);
  const status = STATUSES.indexOf(p.status) >= 0 ? p.status : 'جديد';
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const idx = ids.indexOf(id);
  const total = Number(p.total) || 0;
  let notify = false;
  const rowVals = [new Date(), id, status, clean_(p.name, 100), clean_(p.phone, 20), clean_(p.city, 80),
    clean_(p.address, 400), clean_(p.items || ((p.model || '') + (p.qty ? ' × ' + p.qty : '')), 1500), total, clean_(p.source, 80)];
  if (idx === -1) {
    notify = status === 'جديد';
    sh.getRange(last + 1, 1, 1, rowVals.length).setValues([rowVals]);
  } else {
    const row = idx + 2;
    const old = String(sh.getRange(row, 3).getValue());
    rowVals[0] = sh.getRange(row, 1).getValue();                 // نحتفظ بتاريخ أول تسجيل
    if (old !== 'لم يؤكد بعد') rowVals[2] = old; else notify = status === 'جديد';   // لا نغيّر حالة وضعتها بنفسك
    sh.getRange(row, 1, 1, rowVals.length).setValues([rowVals]);
  }
  if (notify) notifyOrder_(rowVals);
}

function notifyOrder_(r) {
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'طلب جديد #' + r[1] + ' — ' + r[8] + ' درهم',
      'الاسم: ' + r[3] + '\nالهاتف: ' + r[4] + '\nالمدينة: ' + r[5] + '\nالعنوان: ' + r[6] + '\nالطلب: ' + r[7] + '\nالمجموع: ' + r[8] + ' درهم\n\nافتح مساحة البائع لتأكيده.');
  } catch (e) {}
}

/* ---------- الزيارات والأحداث ---------- */

function logVisit_(p) {
  const sh = book_().getSheetByName(SH.VISITS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues([[new Date(), clean_(p.vid, 60), clean_(p.src, 60) || 'مباشر',
    clean_(p.dev, 20), p.isNew === '1' ? 'نعم' : 'لا']]);
}

function logEvent_(p) {
  const allowed = ['view_product', 'add_to_cart', 'start_checkout'];
  if (allowed.indexOf(p.ev) === -1) return;
  const sh = book_().getSheetByName(SH.EVENTS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[new Date(), p.ev, clean_(p.item, 120), clean_(p.vid, 60)]]);
}

/* ---------- المنتجات ---------- */

function getProducts_(all) {
  const sh = book_().getSheetByName(SH.PRODUCTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
  return rows.map(function (r, i) { return toProduct_(r, i + 2, all); }).filter(Boolean);
}

function toProduct_(r, rowNum, all) {
  const str = v => String(v == null ? '' : v).trim();
  const list = (v, sep) => str(v).split(sep).map(s => s.trim()).filter(Boolean);
  const name = str(r[1]);
  const price = Number(r[3]);
  if (!name || !(price > 0)) return null;
  const off = str(r[10]) === 'لا';
  if (off && !all) return null;
  const variants = list(r[8], '|').map(function (t) {
    const parts = t.split('=');
    const o = { name: parts[0].trim() };
    const pr = Number(parts[1]);
    if (parts.length > 1 && pr > 0) o.price = pr;
    return o;
  }).filter(v => v.name);
  const id = str(r[0]).replace(/[^\w؀-ۿ-]+/g, '-') || ('p' + rowNum);
  const p = { id: id, name: name, category: str(r[2]), price: price,
    images: list(r[4], /[,\n]/), summary: str(r[5]), details: list(r[6], '|') };
  if (variants.length) { p.variantLabel = str(r[7]) || 'النوع'; p.variants = variants; }
  if (str(r[9])) p.badge = str(r[9]);
  if (all) p.off = off ? 1 : 0;
  return p;
}

/* ---------- إعدادات المتجر ---------- */

function publicContent_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('content');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const txt = JSON.stringify({ ok: true, products: getProducts_(), settings: getSettings_() });
  try { if (txt.length < 90000) cache.put('content', txt, 300); } catch (x) {}
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
function bust_() { try { CacheService.getScriptCache().remove('content'); } catch (x) {} }

function getSettings_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SETTINGS') || '{}'); } catch (x) { return {}; }
}

function saveSettings_(c) {
  c = c || {};
  const str = (v, m) => String(v == null ? '' : v).trim().slice(0, m);
  const num = v => { v = Number(v); return v >= 0 && v < 100000 ? Math.round(v * 100) / 100 : 0; };
  const img = v => { v = str(v, 400); return /^https:\/\//.test(v) || /^[\w.\-\/]+\.(jpe?g|png|webp)$/i.test(v) ? v : ''; };
  const url = v => { v = str(v, 300); return /^https:\/\//.test(v) ? v : ''; };
  const out = { topbar: str(c.topbar, 160), heroTitle: str(c.heroTitle, 120), heroSub: str(c.heroSub, 300), heroImage: img(c.heroImage),
    whatsapp: str(c.whatsapp, 20).replace(/\D/g, ''), deliveryFee: num(c.deliveryFee), freeDeliveryFrom: num(c.freeDeliveryFrom),
    instagram: url(c.instagram), facebook: url(c.facebook) };
  out.reviews = (Array.isArray(c.reviews) ? c.reviews : []).slice(0, 12).map(r => ({ name: str(r.name, 60),
    stars: Math.min(5, Math.max(1, Number(r.stars) || 5)), text: str(r.text, 300) })).filter(r => r.name && r.text);
  out.faq = (Array.isArray(c.faq) ? c.faq : []).slice(0, 12).map(x => ({ q: str(x.q, 160), a: str(x.a, 500) })).filter(x => x.q && x.a);
  const txt = JSON.stringify(out);
  if (txt.length > 8500) throw new Error('الإعدادات كبيرة جداً، اختصر الآراء أو الأسئلة');
  PropertiesService.getScriptProperties().setProperty('SETTINGS', txt);
  bust_();
}

/* ---------- حفظ المنتجات من مساحة البائع (تُكتب في ورقة «المنتجات») ---------- */

function saveProducts_(list) {
  list = (Array.isArray(list) ? list : []).slice(0, 300);
  const str = (v, m) => String(v == null ? '' : v).trim().slice(0, m);
  const seen = {};
  const rows = [];
  list.forEach(function (p, i) {
    const name = str(p.name, 120), price = Number(p.price);
    if (!name || !(price > 0)) return;
    let id = str(p.id, 60).replace(/[^\w\-]+/g, '-').replace(/^-+|-+$/g, '');
    if (!id || seen[id]) id = 'p' + Date.now().toString(36) + i;
    seen[id] = 1;
    const imgs = (Array.isArray(p.images) ? p.images : []).map(u => str(u, 400)).filter(u => /^https:\/\//.test(u) || /^[\w.\-\/]+\.(jpe?g|png|webp)$/i.test(u)).slice(0, 8);
    const det = (Array.isArray(p.details) ? p.details : []).map(d => str(d, 120).replace(/\|/g, '/')).filter(Boolean).slice(0, 12);
    const vars = (Array.isArray(p.variants) ? p.variants : []).map(function (v) {
      const n = str(v.name, 50).replace(/[|=]/g, ' '), vp = Number(v.price);
      return n ? (vp > 0 ? n + '=' + vp : n) : '';
    }).filter(Boolean).slice(0, 20);
    rows.push([id, name, str(p.category, 40), price, imgs.join(', '), str(p.summary, 400), det.join(' | '),
      str(p.variantLabel, 30), vars.join(' | '), str(p.badge, 24), p.off ? 'لا' : 'نعم']);
  });
  const sh = book_().getSheetByName(SH.PRODUCTS);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, PRODUCT_HEADERS.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@'); sh.getRange(2, 5, rows.length, 7).setNumberFormat('@'); sh.getRange(2, 2, rows.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, PRODUCT_HEADERS.length).setValues(rows);
  }
  bust_();
}

function adminProducts_() {
  return getProducts_(true);
}

function uploadImage_(name, b64) {
  b64 = String(b64 || '').replace(/^data:[^,]+,/, '');
  if (!b64 || b64.length > 3000000) throw new Error('الصورة كبيرة جداً');
  const it = DriveApp.getFoldersByName('Crystal.shop images');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Crystal.shop images');
  const safe = String(name || 'img').replace(/[^\w\-]+/g, '-').slice(0, 40) + '-' + Date.now() + '.jpg';
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', safe));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w1000';
}

function adminPost_(raw) {
  try {
    const p = JSON.parse(raw || '{}');
    if (ADMIN_PASSWORD === 'CHANGE_ME') return json_({ ok: false, error: 'setpw' });
    if (String(p.pw || '') !== ADMIN_PASSWORD) return json_({ ok: false, error: 'auth' });
    if (p.action === 'saveProducts') { saveProducts_(p.products); return json_({ ok: true, data: adminData_() }); }
    if (p.action === 'saveSettings') { saveSettings_(p.settings); return json_({ ok: true, data: adminData_() }); }
    if (p.action === 'upload') return json_({ ok: true, url: uploadImage_(p.name, p.data) });
    return json_({ ok: false, error: 'unknown' });
  } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------- مساحة البائع ---------- */

function updateOrder_(id, col, value) {
  if (value == null) return;
  const sh = book_().getSheetByName(SH.ORDERS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.getRange(i + 2, col).setValues([[value]]);
}

function dayKey_(d) { return Utilities.formatDate(new Date(d), 'Africa/Casablanca', 'yyyy-MM-dd'); }

function rows_(name, ncols) {
  const sh = book_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ncols).getValues();
}

function adminData_() {
  const orders = rows_(SH.ORDERS, ORDER_HEADERS.length).map(function (r) {
    return { date: r[0] ? new Date(r[0]).toISOString() : '', id: String(r[1]), status: String(r[2]), name: String(r[3]),
      phone: String(r[4]), city: String(r[5]), address: String(r[6]), items: String(r[7]), total: Number(r[8]) || 0,
      source: String(r[9]), note: String(r[10] || '') };
  }).filter(o => o.id).reverse();
  const visits = rows_(SH.VISITS, 5), events = rows_(SH.EVENTS, 4);

  const today = dayKey_(new Date()), days = {}, order = [];
  for (let i = 29; i >= 0; i--) {
    const k = dayKey_(new Date(Date.now() - i * 86400000));
    days[k] = { d: k, visitors: {}, orders: 0, sales: 0 }; order.push(k);
  }
  const allV = {}, srcs = {}, devs = {};
  visits.forEach(function (r) {
    const k = dayKey_(r[0]), vid = String(r[1]);
    allV[vid] = 1;
    if (days[k]) days[k].visitors[vid] = 1;
    const sk = String(r[2] || 'مباشر'); srcs[sk] = (srcs[sk] || 0) + 1;
    const dk = String(r[3] || '—'); devs[dk] = (devs[dk] || 0) + 1;
  });
  const real = o => o.status !== 'ملغي' && o.status !== 'لم يؤكد بعد';
  let sales = 0, delivered = 0, cnt = 0, fresh = 0, abandoned = 0, todayOrders = 0, todaySales = 0;
  orders.forEach(function (o) {
    if (o.status === 'لم يؤكد بعد') abandoned++;
    if (o.status === 'جديد') fresh++;
    if (!real(o)) return;
    cnt++; sales += o.total;
    if (o.status === 'تم التسليم') delivered += o.total;
    const k = o.date ? dayKey_(o.date) : '';
    if (k === today) { todayOrders++; todaySales += o.total; }
    if (days[k]) { days[k].orders++; days[k].sales += o.total; }
  });
  const evc = {}, started = {};
  events.forEach(function (r) { evc[r[1]] = (evc[r[1]] || 0) + 1; if (r[1] === 'start_checkout') started[String(r[3])] = 1; });
  const top = function (m) { return Object.keys(m).map(k => ({ name: k, n: m[k] })).sort((a, b) => b.n - a.n).slice(0, 8); };
  return {
    now: new Date().toISOString(),
    totals: { orders: cnt, sales: sales, delivered: delivered, fresh: fresh, abandoned: abandoned, todayOrders: todayOrders,
      todaySales: todaySales, visitors: Object.keys(allV).length, visits: visits.length,
      todayVisitors: Object.keys(days[today] ? days[today].visitors : {}).length,
      views: evc.view_product || 0, started: Object.keys(started).length },
    days: order.map(k => ({ d: k, visitors: Object.keys(days[k].visitors).length, orders: days[k].orders, sales: days[k].sales })),
    sources: top(srcs), devices: top(devs),
    orders: orders.slice(0, 500), products: adminProducts_(), settings: getSettings_()
  };
}

/* ---------- التجهيز الأول (شغّله مرة واحدة) ---------- */

function setup() {
  MailApp.getRemainingDailyQuota(); DriveApp.getRootFolder();   // لطلب الأذونات مرة واحدة
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SSID', ss.getId());
  ss.setSpreadsheetTimeZone('Africa/Casablanca');

  const dash = sheet_(ss, SH.DASH), prod = sheet_(ss, SH.PRODUCTS), ord = sheet_(ss, SH.ORDERS),
    vis = sheet_(ss, SH.VISITS), ev = sheet_(ss, SH.EVENTS);
  ['Sheet1', 'ورقة1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });

  header_(prod, PRODUCT_HEADERS); header_(ord, ORDER_HEADERS); header_(vis, VISIT_HEADERS); header_(ev, EVENT_HEADERS);

  if (prod.getLastRow() < 2) {
    prod.getRange(2, 1, 1, PRODUCT_HEADERS.length).setValues([[
      'crystal-3d-led', 'بلورة كريستال 3D مضيئة بإضاءة LED', 'هدايا وديكور', 99, '14184.jpg',
      'نقش ثلاثي الأبعاد داخل الكريستال وقاعدة خشبية بإضاءة LED دافئة. هدية مميزة وديكور أنيق.',
      'زجاج كريستالي | نقش ثلاثي الأبعاد داخل الزجاج | قاعدة خشبية بإضاءة LED دافئة | مناسبة للهدايا والديكور',
      'الموديل', 'العرسان | القمر | الغزال', '', 'نعم']]);
  }
  prod.getRange('K2:K1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).build());
  prod.getRange('A:A').setNumberFormat('@'); prod.getRange('E:K').setNumberFormat('@'); prod.getRange('B:C').setNumberFormat('@');
  prod.setColumnWidths(1, 1, 130); prod.setColumnWidth(2, 240); prod.setColumnWidth(5, 200); prod.setColumnWidth(6, 260);
  prod.setColumnWidth(7, 260); prod.setColumnWidth(9, 220);
  prod.getRange('A1').setNote('اختياري: كلمة لاتينية قصيرة تظهر في رابط المنتج، مثل lamp-moon. اتركها فارغة وسيُنشأ لها رقم تلقائي.');
  prod.getRange('E1').setNote('أسماء الصور المرفوعة على GitHub أو روابط كاملة. افصل بين الصور بفاصلة.');
  prod.getRange('G1').setNote('اكتب كل ميزة ثم افصل بينها بالرمز |');
  prod.getRange('I1').setNote('اكتب الأنواع مفصولة بالرمز | ويمكن إضافة سعر خاص بعد علامة =\nمثال: أبيض | دافئ=159');
  prod.getRange('K1').setNote('اختر "لا" لإخفاء المنتج من المتجر عند نفاد المخزون.');

  ord.getRange('B:B').setNumberFormat('@'); ord.getRange('D:H').setNumberFormat('@'); ord.getRange('J:K').setNumberFormat('@');
  ord.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm'); ord.getRange('I:I').setNumberFormat('#,##0');
  ord.getRange('C2:C5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(true).build());
  ord.setColumnWidth(1, 130); ord.setColumnWidth(2, 110); ord.setColumnWidth(3, 110); ord.setColumnWidth(4, 160); ord.setColumnWidth(5, 120);
  ord.setColumnWidth(7, 260); ord.setColumnWidth(8, 320);
  vis.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm'); ev.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');

  buildDashboard_(dash);

  ss.setActiveSheet(dash); ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

function sheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(sh, headers) {
  sh.setRightToLeft(true);
  const r = sh.getRange(1, 1, 1, headers.length);
  r.setValues([headers]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#064e3b').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
}

function buildDashboard_(sh) {
  sh.clear(); sh.setRightToLeft(true);
  const O = "'" + SH.ORDERS + "'!", V = "'" + SH.VISITS + "'!", E = "'" + SH.EVENTS + "'!";
  const real = O + 'C2:C,"<>ملغي",' + O + 'C2:C,"<>لم يؤكد بعد",' + O + 'C2:C,"<>"';
  const rows = [
    ['الطلبات', ''],
    ['كل الطلبات (بدون الملغاة)', '=COUNTIFS(' + real + ')'],
    ['طلبات جديدة تنتظر التأكيد', '=COUNTIF(' + O + 'C2:C,"جديد")'],
    ['طلبات اليوم', '=COUNTIFS(' + O + 'A2:A,">="&TODAY(),' + real + ')'],
    ['إجمالي المبيعات', '=SUMIFS(' + O + 'I2:I,' + real + ')'],
    ['مبيعات تم تسليمها', '=SUMIFS(' + O + 'I2:I,' + O + 'C2:C,"تم التسليم")'],
    ['تركوا بياناتهم ولم يرسلوا الطلب', '=COUNTIF(' + O + 'C2:C,"لم يؤكد بعد")'],
    ['', ''],
    ['الزوار', ''],
    ['زوار المتجر (أشخاص مختلفون)', '=IFERROR(COUNTUNIQUE(' + V + 'B2:B),0)'],
    ['زوار اليوم', '=IFERROR(COUNTUNIQUE(FILTER(' + V + 'B2:B,' + V + 'A2:A>=TODAY())),0)'],
    ['عدد الزيارات', '=COUNTA(' + V + 'B2:B)'],
    ['', ''],
    ['مسار الشراء', ''],
    ['مشاهدات المنتجات', '=COUNTIF(' + E + 'B2:B,"view_product")'],
    ['إضافات إلى السلة', '=COUNTIF(' + E + 'B2:B,"add_to_cart")'],
    ['بدء إتمام الطلب', '=COUNTIF(' + E + 'B2:B,"start_checkout")'],
    ['نسبة التحويل (طلبات ÷ زوار)', '=IFERROR(B2/B10,0)']
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows.map(function (r) { return [r[0], r[1].charAt(0) === '=' ? '' : r[1]]; }));
  rows.forEach(function (r, i) { if (r[1].charAt(0) === '=') sh.getRange(i + 1, 2).setFormula(r[1]); });
  [1, 9, 14].forEach(function (n) { sh.getRange(n, 1, 1, 2).setFontWeight('bold').setFontColor('#ffffff').setBackground('#064e3b'); });
  sh.getRange('B5:B6').setNumberFormat('#,##0 "درهم"');
  sh.getRange('B18').setNumberFormat('0.0%');
  sh.getRange('B2:B18').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 140); sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 110); sh.setColumnWidth(7, 220); sh.setColumnWidth(8, 130);

  sh.getRange('D1').setValue('مصادر الزيارات').setFontWeight('bold').setFontColor('#ffffff').setBackground('#064e3b');
  sh.getRange('E1').setBackground('#064e3b');
  sh.getRange('D2').setFormula('=IFERROR(QUERY(' + V + 'C2:C,"select C, count(C) where C is not null group by C order by count(C) desc limit 8 label C \'المصدر\', count(C) \'الزيارات\'",0),"")');
  sh.getRange('G1').setValue('الأكثر إضافة للسلة').setFontWeight('bold').setFontColor('#ffffff').setBackground('#064e3b');
  sh.getRange('H1').setBackground('#064e3b');
  sh.getRange('G2').setFormula('=IFERROR(QUERY(' + E + 'B2:C,"select C, count(C) where B=\'add_to_cart\' group by C order by count(C) desc limit 8 label C \'المنتج\', count(C) \'المرات\'",0),"")');
  sh.getRange('A20').setValue('تتحدث الأرقام تلقائيًا. غيّر حالة كل طلب من ورقة «الطلبات».').setFontColor('#64748b');
}
