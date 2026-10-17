/* Crystal.shop — مزامنة المنتجات من لوحة البائع
   يُحمَّل بعد سكربت المتجر الرئيسي، ويعيد رسم الموديلات والأثمان حسب ما حفظته في اللوحة.
   إذا فشل الاتصال يبقى المتجر يعمل بالموديلات الافتراضية الموجودة في index.html. */
(function () {
  'use strict';
  if (typeof CONFIG === 'undefined' || !CONFIG.sheetsUrl) return;

  var CACHE = 'cs_cfg';
  var P = [];                         // المنتجات الظاهرة حالياً
  var tier = { 1: 0, 2: 0, 3: 0 };    // نسب التخفيض حسب الكمية
  var lastApplied = '';

  var st = document.createElement('style');
  st.textContent = '.model .p s{color:#94a3b8;font-size:.95rem;font-weight:500;margin-inline-start:8px}';
  document.head.appendChild(st);

  function selected() {
    var name = (typeof getModel === 'function') ? getModel() : '';
    for (var i = 0; i < P.length; i++) if (P[i].name === name) return P[i];
    return P[0];
  }

  // الثمن يُحسب حسب الموديل المختار والكمية
  function price1() { var p = selected(); return p ? Number(p.price) : 99; }
  Object.defineProperty(CONFIG, 'unitPrice', { configurable: true, get: price1 });
  [1, 2, 3].forEach(function (q) {
    Object.defineProperty(CONFIG.prices, q, {
      configurable: true,
      get: function () { return Math.round(price1() * q * (1 - (tier[q] || 0) / 100)); }
    });
  });

  function safeUrl(u) {
    u = String(u || '').trim();
    if (!u) return '';
    if (/^https:\/\//i.test(u) || /^[\w\-./]+$/.test(u)) return u;
    return '';
  }

  function buildCard(p, checked) {
    var logo = (document.querySelector('img[data-logo]') || {}).src || '';
    var label = document.createElement('label'); label.className = 'model';
    var input = document.createElement('input');
    input.type = 'radio'; input.name = 'model'; input.value = p.name; input.setAttribute('form', 'orderForm');
    if (checked) input.checked = true;
    var card = document.createElement('div'); card.className = 'card';
    var thumb = document.createElement('div'); thumb.className = 'thumb';
    var ph = document.createElement('img'); ph.className = 'ph'; ph.src = logo; ph.alt = ''; ph.width = 96; ph.height = 96;
    thumb.appendChild(ph);
    var src = safeUrl(p.img);
    if (src) {
      var im = document.createElement('img'); im.className = 'prod'; im.src = src; im.alt = p.name;
      im.width = 480; im.height = 480; im.loading = 'lazy';
      im.onerror = function () { im.remove(); };
      thumb.appendChild(im);
    }
    var h3 = document.createElement('h3'); h3.textContent = p.name;
    var tag = document.createElement('p'); tag.className = 'tag'; tag.textContent = p.tag || '';
    var pr = document.createElement('div'); pr.className = 'p';
    pr.appendChild(document.createTextNode(Number(p.price) + ' ' + CONFIG.currency));
    if (Number(p.old) > Number(p.price)) { var s = document.createElement('s'); s.textContent = Number(p.old) + ' ' + CONFIG.currency; pr.appendChild(s); }
    var pick = document.createElement('span'); pick.className = 'pick';
    var c1 = document.createElement('span'); c1.className = 'check'; c1.textContent = '✔ تم الاختيار';
    var c2 = document.createElement('span'); c2.className = 'nocheck'; c2.textContent = 'اختر هذا الموديل';
    pick.appendChild(c1); pick.appendChild(c2);
    card.appendChild(thumb); card.appendChild(h3); card.appendChild(tag); card.appendChild(pr); card.appendChild(pick);
    label.appendChild(input); label.appendChild(card);
    return label;
  }

  function apply(cfg) {
    var key = JSON.stringify(cfg);
    if (key === lastApplied) return;
    var list = (cfg.products || []).filter(function (p) { return p && p.on !== false && p.name && Number(p.price) > 0; });
    if (!list.length) return;
    lastApplied = key;

    var keep = (typeof getModel === 'function') ? getModel() : '';
    P = list;
    tier = { 1: 0, 2: Number(cfg.d2) || 0, 3: Number(cfg.d3) || 0 };

    var box = document.querySelector('.models');
    if (box) {
      var has = list.some(function (p) { return p.name === keep; });
      box.innerHTML = '';
      list.forEach(function (p, i) { box.appendChild(buildCard(p, has ? p.name === keep : i === 0)); });
      box.querySelectorAll('input[name=model]').forEach(function (r) {
        r.addEventListener('change', function () {
          refresh();
          track('view', { model: getModel(), value: CONFIG.unitPrice });
        });
      });
    }

    var min = Math.min.apply(null, list.map(function (p) { return Number(p.price); }));
    document.querySelectorAll('[data-unit-price]').forEach(function (el) { el.textContent = min; });
    if (cfg.topbar) {
      var tb = document.querySelector('.topbar');
      if (tb) tb.textContent = String(cfg.topbar).replace(/\{price\}/g, min);
    }
    var wa = String(cfg.whatsapp || '').replace(/\D/g, '');
    if (wa.length >= 10) {
      CONFIG.whatsapp = wa;
      var f = document.getElementById('waFloat');
      if (f) f.href = waLink('سلام عليكم، بغيت نستفسر على بلورة الكريستال من متجر Crystal.shop');
      var fw = document.getElementById('footWa');
      if (fw) { fw.href = waLink(''); var s = fw.querySelector('span'); if (s) s.textContent = '0' + wa.replace(/^212/, ''); }
    }
    refresh();
  }

  try {
    var cached = JSON.parse(localStorage.getItem(CACHE) || 'null');
    if (cached && cached.products) apply(cached);
  } catch (e) {}

  fetch(CONFIG.sheetsUrl + (CONFIG.sheetsUrl.indexOf('?') > -1 ? '&' : '?') + 'action=products&_=' + Date.now())
    .then(function (r) { return r.json(); })
    .then(function (j) {
      if (j && j.ok && j.config && j.config.products) {
        try { localStorage.setItem(CACHE, JSON.stringify(j.config)); } catch (e) {}
        apply(j.config);
      }
    })
    .catch(function () {});
})();
