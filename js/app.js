/* ITR Generator – Annual V2.0 — main application */
(function () {
  'use strict';

  const CFG = window.ITR_CONFIG || {};
  const LIBS = {
    xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    exceljs: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
    docx: 'https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js',
    pdfmake: 'https://cdn.jsdelivr.net/npm/pdfmake@0.2.20/build/pdfmake.min.js',
    pdfmakeFonts: 'https://cdn.jsdelivr.net/npm/pdfmake@0.2.20/build/vfs_fonts.js',
    pdfjs: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    pdfjsWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js'
  };
  // separate keys from the quarterly app (both may run on the same github.io domain)
  const LS = { token: 'itra_token', user: 'itra_user', api: 'itra_api_url', banner: 'itra_install_banner_seen', landing: 'itra_landing' };

  /* ======================= utilities ======================= */
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => {
    const v = r2(n);
    const s = Math.abs(v).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return v < 0 ? '(₱' + s + ')' : '₱' + s;
  };
  const plain = (n) => r2(n).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const parseNum = (s) => {
    const t = String(s || '').replace(/[₱,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
    const n = parseFloat(t);
    return isFinite(n) ? r2(n) : 0;
  };
  const sizeFmt = (b) => b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const uid = () => Math.random().toString(36).slice(2, 10);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function toast(msg, type = '', ms = 4200) {
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => { el.style.transition = 'opacity .3s'; el.style.opacity = '0'; setTimeout(() => el.remove(), 320); }, ms);
  }

  const loaded = {};
  function loadScript(src) {
    if (!loaded[src]) {
      loaded[src] = new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = src; s.async = true;
        s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error('Could not load ' + src.split('/').pop() + '. Check your connection.')); };
        document.head.appendChild(s);
      });
    }
    return loaded[src];
  }

  function openModal(id) { $(id).hidden = false; }
  function closeModal(id) { $(id).hidden = true; }
  document.addEventListener('click', (e) => {
    const c = e.target.closest('[data-close]');
    if (c) { const m = c.closest('.modal'); if (m) { m.hidden = true; m.dispatchEvent(new Event('closed')); } }
  });

  /* ======================= state ======================= */
  const thisYear = new Date().getFullYear();
  // Annual returns: taxable year 2025 and upwards; default = the year just ended (filed this year)
  const FIRST_YEAR = CFG.FIRST_YEAR || 2025;
  const LAST_YEAR = Math.max(thisYear, FIRST_YEAR) + (CFG.YEARS_AHEAD || 4);
  const DEFAULT_YEAR = Math.min(Math.max(thisYear - 1, FIRST_YEAR), LAST_YEAR);
  const yearOptions = () => { let h = ''; for (let y = FIRST_YEAR; y <= LAST_YEAR; y++) h += '<option value="' + y + '">' + y + '</option>'; return h; };
  const validYear = (y) => (Number(y) >= FIRST_YEAR && Number(y) <= LAST_YEAR ? Number(y) : DEFAULT_YEAR);
  const S = {
    token: localStorage.getItem(LS.token),
    user: JSON.parse(localStorage.getItem(LS.user) || 'null'),
    clients: [], client: null,
    year: DEFAULT_YEAR, form: null,
    opts: { method: 'itemized', taxType: 'rcit25', mcitExempt: false, mcitRate: 0.02, sellerOf: 'Goods' },
    salesType: 'VAT',
    preset: JSON.parse(localStorage.getItem(LS.landing) || 'null'),
    files: {},
    X: null, adj: {}, M: null, tab: 'summary', busy: false
  };

  /* ======================= API ======================= */
  function apiUrl() {
    const saved = localStorage.getItem(LS.api);
    if (saved) return saved;
    return /^https:\/\//.test(CFG.API_URL || '') ? CFG.API_URL : '';
  }

  async function api(action, payload = {}) {
    const url = apiUrl();
    if (!url) { openSettings(); throw new Error('Set the backend Web App URL first.'); }
    let res;
    try {
      // text/plain avoids a CORS preflight, which Apps Script does not answer
      res = await fetch(url, { method: 'POST', body: JSON.stringify(Object.assign({ action, token: S.token }, payload)), redirect: 'follow' });
    } catch (e) {
      throw new Error('Cannot reach the backend. Check your internet connection and the Web App URL.');
    }
    let data;
    try { data = await res.json(); } catch (e) { throw new Error('Unexpected response from the backend (HTTP ' + res.status + '). Is the Web App deployed for "Anyone"?'); }
    if (!data.ok) {
      if (data.code === 'AUTH' && action !== 'login') { logout(true); }
      throw new Error(data.error || 'Request failed.');
    }
    return data;
  }

  /* ======================= intro & install ======================= */
  let introDone = false, deferredPrompt = null, bannerPending = false;
  const isStandalone = () => matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

  function runIntro() {
    const intro = $('#intro');
    const seen = sessionStorage.getItem('itr_intro');
    const finish = () => {
      if (introDone) return;
      introDone = true;
      intro.classList.add('done');
      setTimeout(() => intro.remove(), 700);
      sessionStorage.setItem('itr_intro', '1');
      if (bannerPending || !localStorage.getItem(LS.banner)) setTimeout(maybeShowBanner, 900);
    };
    intro.addEventListener('click', finish);
    setTimeout(finish, seen ? 900 : 2800);
  }

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    $('#menuInstall').hidden = false;
    if (!localStorage.getItem(LS.banner)) maybeShowBanner();
  });
  window.addEventListener('appinstalled', () => {
    localStorage.setItem(LS.banner, '1');
    hideBanner();
    toast('ITR Generator installed.', 'ok');
  });

  function maybeShowBanner() {
    if (isStandalone() || localStorage.getItem(LS.banner)) return;
    if (!introDone) { bannerPending = true; return; }
    const b = $('#installBanner');
    if (!b.hidden) return;
    if (deferredPrompt) {
      $('#installHint').textContent = 'Add it to your home screen or desktop for a faster, full-screen experience.';
      $('#installNow').textContent = 'Install';
    } else if (isIOS()) {
      $('#installHint').innerHTML = 'Tap <b>Share</b> <span aria-hidden="true">⎋</span> then <b>Add to Home Screen</b>.';
      $('#installNow').textContent = 'Got it';
    } else {
      $('#installHint').textContent = 'Use your browser menu → "Install app" or "Add to Home screen".';
      $('#installNow').textContent = 'Got it';
    }
    b.hidden = false;
    requestAnimationFrame(() => requestAnimationFrame(() => b.classList.add('show')));
  }
  function hideBanner() {
    const b = $('#installBanner');
    b.classList.remove('show');
    setTimeout(() => { b.hidden = true; }, 500);
  }
  async function doInstall() {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      const choice = await deferredPrompt.userChoice.catch(() => ({}));
      deferredPrompt = null;
      $('#menuInstall').hidden = true;
      if (choice && choice.outcome === 'accepted') localStorage.setItem(LS.banner, '1');
    }
    localStorage.setItem(LS.banner, '1');
    hideBanner();
  }
  $('#installNow').addEventListener('click', doInstall);
  $('#installLater').addEventListener('click', () => { localStorage.setItem(LS.banner, '1'); hideBanner(); });
  $('#menuInstall').addEventListener('click', () => { closeUserMenu(); doInstall(); });

  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }

  /* ======================= settings ======================= */
  function openSettings() {
    $('#apiUrlInput').value = apiUrl();
    $('#settingsMsg').textContent = '';
    $('#settingsMsg').className = 'form-note';
    openModal('#settingsModal');
  }
  $('#openSettings').addEventListener('click', openSettings);
  $('#menuSettings').addEventListener('click', () => { closeUserMenu(); openSettings(); });
  async function testUrl(url) {
    const msg = $('#settingsMsg');
    msg.className = 'form-note'; msg.textContent = 'Testing…';
    try {
      const res = await fetch(url, { method: 'POST', body: JSON.stringify({ action: 'ping' }) });
      const d = await res.json();
      if (!d.ok) throw new Error(d.error || 'Failed');
      msg.className = 'form-note ok'; msg.textContent = 'Connected ✓ ' + d.app + ' v' + d.version + ' · model ' + d.model;
      return true;
    } catch (e) {
      msg.className = 'form-note bad'; msg.textContent = 'Could not connect: ' + e.message;
      return false;
    }
  }
  $('#testApi').addEventListener('click', () => testUrl($('#apiUrlInput').value.trim()));
  $('#saveApi').addEventListener('click', async () => {
    const url = $('#apiUrlInput').value.trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) {
      $('#settingsMsg').className = 'form-note bad';
      $('#settingsMsg').textContent = 'That does not look like an Apps Script Web App URL.';
      return;
    }
    localStorage.setItem(LS.api, url);
    if (await testUrl(url)) setTimeout(() => closeModal('#settingsModal'), 700);
  });

  /* ======================= auth ======================= */
  function showLanding() {
    $('#appView').hidden = true;
    $('#landingView').hidden = false;
    syncLanding();
    window.scrollTo(0, 0);
  }
  function showApp() {
    $('#landingView').hidden = true;
    $('#appView').hidden = false;
    window.scrollTo(0, 0);
    const name = (S.user && S.user.name) || 'User';
    $('#userName').textContent = name;
    $('#userAvatar').textContent = name.trim().charAt(0).toUpperCase();
  }
  function openSignin() {
    const p = landingValues();
    $('#siChips').innerHTML = '<span>Tax Year ' + p.year + '</span><span>BIR Form ' + p.form + '</span><span>' + p.vat + '</span>';
    $('#loginError').textContent = '';
    openModal('#signinModal');
    setTimeout(() => $('#loginUser').focus(), 80);
  }

  /* ---------- landing ---------- */
  (function initLanding() {
    const sel = $('#ldYear');
    sel.innerHTML = yearOptions();
    const p = S.preset || {};
    sel.value = String(validYear(p.year));
    $('#ldFormSel').value = p.form === '1702RT' ? '1702RT' : '1701';
    setLdVat(p.vat === 'Non-VAT' ? 'Non-VAT' : 'VAT');
    $('#ldVat').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setLdVat(b.dataset.v); });
    $('#ldForm').addEventListener('submit', (e) => {
      e.preventDefault();
      saveLanding();
      if (S.token) enterApp(); else openSignin();
    });
    $('#ldSigninBtn').addEventListener('click', () => {
      if (S.token) logout(false); else { saveLanding(); openSignin(); }
    });
  })();
  function setLdVat(v) {
    $('#ldVat').dataset.v = v;
    $$('#ldVat button').forEach((b) => { const on = b.dataset.v === v; b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on)); });
  }
  function landingValues() {
    return { year: Number($('#ldYear').value), form: $('#ldFormSel').value, vat: $('#ldVat').dataset.v || 'VAT' };
  }
  function saveLanding() {
    S.preset = landingValues();
    localStorage.setItem(LS.landing, JSON.stringify(S.preset));
  }
  function syncLanding() {
    const p = S.preset;
    if (p) { p.year = validYear(p.year); $('#ldYear').value = String(p.year); $('#ldFormSel').value = p.form; setLdVat(p.vat); }
    const signed = !!S.token;
    $('#ldHello').hidden = !signed;
    $('#ldHello').textContent = signed ? 'Signed in as ' + ((S.user && S.user.name) || '') : '';
    $('span', $('#ldSigninBtn')).textContent = signed ? 'Sign out' : 'Sign in';
  }
  async function enterApp() {
    const p = S.preset || landingValues();
    p.year = validYear(p.year);
    const changed = S.year !== p.year || S.form !== p.form || S.salesType !== p.vat;
    S.year = p.year; S.form = p.form; S.salesType = p.vat;
    $('#yearSelect').value = String(S.year);
    if (S.form === '1702RT' || S.salesType === 'VAT') { if (S.opts.method === 'eight') S.opts.method = 'itemized'; }
    S.opts.mcitRate = defaultMcitRate();
    if (changed && S.X) clearResults();
    showApp();
    refresh();
    if (!S.clients.length) await loadClients();
    if (!S.client && matchMedia('(pointer: fine)').matches) setTimeout(() => $('#clientInput').focus(), 250);
  }
  function setPreset(k, v) {
    S.preset = Object.assign({ year: S.year, form: S.form, vat: S.salesType }, S.preset || {}, { [k]: v });
    localStorage.setItem(LS.landing, JSON.stringify(S.preset));
  }

  $('#pwToggle').addEventListener('click', () => {
    const i = $('#loginPass');
    i.type = i.type === 'password' ? 'text' : 'password';
    $('#pwToggle').textContent = i.type === 'password' ? 'Show' : 'Hide';
  });

  $('#loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = $('#loginBtn');
    $('#loginError').textContent = '';
    btn.classList.add('loading');
    try {
      const d = await api('login', { username: $('#loginUser').value.trim(), password: $('#loginPass').value });
      S.token = d.token; S.user = d.user;
      localStorage.setItem(LS.token, d.token);
      localStorage.setItem(LS.user, JSON.stringify(d.user));
      $('#loginPass').value = '';
      closeModal('#signinModal');
      await enterApp();
      toast('Welcome, ' + d.user.name + '!', 'ok');
      if (d.mustChangePassword) openPassword('You signed in with a temporary password. Please set your own now.');
    } catch (err) {
      $('#loginError').textContent = err.message;
    } finally {
      btn.classList.remove('loading');
    }
  });

  /* ---------- sign-up ---------- */
  let staffCache = null;
  async function openSignup() {
    closeModal('#signinModal');
    $('#suError').textContent = '';
    openModal('#signupModal');
    const sel = $('#suStaff');
    sel.innerHTML = '<option value="">Loading staff…</option>';
    sel.disabled = true;
    try {
      const d = await api('staffList');
      staffCache = d;
      const open = d.staff.filter((x) => !x.registered).length;
      sel.innerHTML = '<option value="">— Select your name —</option>' + d.staff.map((x) =>
        '<option value="' + esc(x.name) + '"' + (x.registered ? ' disabled' : '') + '>' + esc(x.name) +
        (x.registered ? ' (already registered)' : ' · ' + x.clients + ' client' + (x.clients === 1 ? '' : 's')) + '</option>').join('');
      sel.disabled = false;
      $('#suStaffHint').className = 'field-hint';
      $('#suStaffHint').textContent = open ? open + ' of ' + d.staff.length + ' staff can still enroll.' : 'All assigned staff already have accounts.';
    } catch (e) {
      sel.innerHTML = '<option value="">Could not load staff</option>';
      $('#suError').textContent = e.message;
    }
  }
  $('#suStaff').addEventListener('change', (e) => {
    const x = staffCache && staffCache.staff.find((s) => s.name === e.target.value);
    const h = $('#suStaffHint');
    if (x) { h.className = 'field-hint ok'; h.textContent = 'Username: ' + x.name + ' · ' + x.clients + ' assigned client' + (x.clients === 1 ? '' : 's'); }
  });
  $('#toSignup').addEventListener('click', openSignup);
  $('#toSignin').addEventListener('click', () => { closeModal('#signupModal'); openSignin(); });
  $('#signupForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const err = $('#suError');
    err.textContent = '';
    const staff = $('#suStaff').value, name = $('#suName').value.trim(), p1 = $('#suPass').value, p2 = $('#suPass2').value;
    if (!staff) { err.textContent = 'Select your name from the assigned staff list.'; return; }
    if (name.length < 2) { err.textContent = 'Enter your full name.'; return; }
    if (p1.length < 8) { err.textContent = 'Password must be at least 8 characters.'; return; }
    if (p1 !== p2) { err.textContent = 'Passwords do not match.'; return; }
    const btn = $('#suBtn');
    btn.classList.add('loading');
    try {
      const d = await api('signup', { staff, fullName: name, password: p1 });
      S.token = d.token; S.user = d.user;
      localStorage.setItem(LS.token, d.token);
      localStorage.setItem(LS.user, JSON.stringify(d.user));
      $('#signupForm').reset();
      closeModal('#signupModal');
      if (!S.preset) saveLanding();
      await enterApp();
      toast('Welcome, ' + d.user.name + '! Your account is ready.', 'ok', 5000);
    } catch (ex) {
      err.textContent = ex.message;
    } finally {
      btn.classList.remove('loading');
    }
  });

  function logout(expired) {
    if (!expired && S.token) api('logout').catch(() => {});
    S.token = null; S.user = null;
    localStorage.removeItem(LS.token); localStorage.removeItem(LS.user);
    resetAll(true);
    showLanding();
    if (expired) toast('Your session has expired. Please sign in again.', 'err');
  }

  // user menu
  function closeUserMenu() { $('#userMenu').hidden = true; $('#userMenuBtn').setAttribute('aria-expanded', 'false'); }
  $('#userMenuBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    const m = $('#userMenu'); m.hidden = !m.hidden;
    $('#userMenuBtn').setAttribute('aria-expanded', String(!m.hidden));
  });
  document.addEventListener('click', (e) => { if (!e.target.closest('.tb-user')) closeUserMenu(); });
  $('#menuLogout').addEventListener('click', () => { closeUserMenu(); logout(false); });
  $('#menuHome').addEventListener('click', () => { closeUserMenu(); showLanding(); });
  $('#homeBtn').addEventListener('click', showLanding);

  function openPassword(note) {
    $('#pwNote').textContent = note || '';
    $('#pwMsg').textContent = '';
    $('#passwordForm').reset();
    openModal('#passwordModal');
  }
  $('#menuPassword').addEventListener('click', () => { closeUserMenu(); openPassword(''); });
  $('#passwordForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = e.submitter || $('#passwordForm .btn-primary');
    btn.classList.add('loading');
    try {
      await api('changePassword', { oldPassword: $('#pwOld').value, newPassword: $('#pwNew').value });
      closeModal('#passwordModal');
      toast('Password updated.', 'ok');
    } catch (err) { $('#pwMsg').textContent = err.message; }
    finally { btn.classList.remove('loading'); }
  });

  /* ======================= step 1: clients combo ======================= */
  async function loadClients() {
    $('#clientInput').placeholder = 'Loading your clients…';
    try {
      const d = await api('clients');
      S.clients = d.clients || [];
      $('#clientInput').placeholder = S.clients.length ? 'Type a name or TIN (' + S.clients.length + ' assigned)…' : 'No clients assigned to you yet';
    } catch (e) {
      $('#clientInput').placeholder = 'Could not load clients';
      toast(e.message, 'err');
    }
  }

  let comboIdx = -1, comboItems = [];
  function highlight(text, q) {
    if (!q) return esc(text);
    const i = text.toLowerCase().indexOf(q.toLowerCase());
    if (i < 0) return esc(text);
    return esc(text.slice(0, i)) + '<mark>' + esc(text.slice(i, i + q.length)) + '</mark>' + esc(text.slice(i + q.length));
  }
  function renderCombo(q) {
    const list = $('#clientList');
    const qq = (q || '').trim().toLowerCase();
    const qDigits = qq.replace(/\D/g, '');
    comboItems = S.clients.filter((c) => !qq || c.name.toLowerCase().includes(qq) || (qDigits.length >= 3 && c.tin.replace(/\D/g, '').includes(qDigits)) || c.id.toLowerCase() === qq).slice(0, 60);
    comboIdx = comboItems.length ? 0 : -1;
    list.innerHTML = comboItems.length
      ? comboItems.map((c, i) => '<li role="option" data-i="' + i + '" class="' + (i === 0 ? 'active' : '') + '"><div><div class="cn">' + highlight(c.name, q) + '</div><div class="ct">' + esc(c.tin) + '</div></div><div><span class="tag ' + (c.type === 'Corporation' ? 'corp' : 'ind') + '">' + c.type + '</span>' + (c.vat ? ' <span class="tag ' + (c.vat === 'VAT' ? 'vat' : 'nonvat') + '">' + c.vat + '</span>' : '') + '</div></li>').join('')
      : '<li class="empty">' + (S.clients.length ? 'No matching client' : 'No clients assigned to you') + '</li>';
    list.hidden = false;
    $('#clientInput').setAttribute('aria-expanded', 'true');
  }
  function closeCombo() { $('#clientList').hidden = true; $('#clientInput').setAttribute('aria-expanded', 'false'); }
  function moveCombo(d) {
    if (!comboItems.length) return;
    comboIdx = (comboIdx + d + comboItems.length) % comboItems.length;
    $$('#clientList li').forEach((li, i) => li.classList.toggle('active', i === comboIdx));
    const a = $('#clientList li.active'); if (a) a.scrollIntoView({ block: 'nearest' });
  }
  $('#clientInput').addEventListener('input', (e) => {
    $('#clientClear').hidden = !e.target.value;
    if (S.client && e.target.value !== S.client.name) selectClient(null, true);
    renderCombo(e.target.value);
  });
  $('#clientInput').addEventListener('focus', (e) => { if (!S.client) renderCombo(e.target.value); else e.target.select(); });
  $('#clientInput').addEventListener('keydown', (e) => {
    const open = !$('#clientList').hidden;
    if (e.key === 'ArrowDown') { e.preventDefault(); if (!open) renderCombo(e.target.value); else moveCombo(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); moveCombo(-1); }
    else if (e.key === 'Enter') { e.preventDefault(); if (open && comboIdx >= 0) selectClient(comboItems[comboIdx]); }
    else if (e.key === 'Escape') closeCombo();
  });
  $('#clientList').addEventListener('mousedown', (e) => {
    const li = e.target.closest('li[data-i]');
    if (li) { e.preventDefault(); selectClient(comboItems[Number(li.dataset.i)]); }
  });
  $('#clientCaret').addEventListener('click', () => {
    if ($('#clientList').hidden) { renderCombo(''); $('#clientInput').focus(); } else closeCombo();
  });
  $('#clientClear').addEventListener('click', () => { $('#clientInput').value = ''; selectClient(null); $('#clientClear').hidden = true; $('#clientInput').focus(); renderCombo(''); });
  document.addEventListener('click', (e) => { if (!e.target.closest('#clientCombo')) closeCombo(); });

  function selectClient(c, silent) {
    const changed = (S.client && S.client.id) !== (c && c.id);
    S.client = c;
    if (c) {
      $('#clientInput').value = c.name;
      $('#clientClear').hidden = false;
      closeCombo();
      $('#clientCard').innerHTML =
        '<div><small>TIN</small><b>' + esc(c.tin || '—') + '</b></div>' +
        '<div><small>Taxpayer type</small><b>' + esc(c.type) + '</b></div>' +
        '<div><small>VAT status</small><b>' + esc(c.vat || 'Not set — choose in step 2') + '</b></div>' +
        '<div><small>RDO</small><b>' + esc(c.rdo || '—') + '</b></div>' +
        (c.lob ? '<div style="grid-column:1/-1"><small>Line of business</small><b>' + esc(c.lob) + '</b></div>' : '') +
        (c.forms ? '<div style="grid-column:1/-1"><small>Enrolled BIR forms</small><b class="forms">' + esc(c.forms) + '</b></div>' : '');
      $('#clientCard').hidden = false;
      S.opts.sellerOf = c.sellerOf || 'Goods';
      if (changed) {
        // BIR Form and VAT status come from the landing page selection; fall back to the client record
        S.form = (S.preset && S.preset.form) || (c.type === 'Corporation' ? '1702RT' : '1701');
        S.salesType = (S.preset && S.preset.vat) || c.vat || 'VAT';
        S.opts.method = 'itemized'; clearResults(); clearFiles();
      }
    } else {
      $('#clientCard').hidden = true;
      if (!silent) $('#clientInput').value = '';
      clearResults();
    }
    refresh();
  }

  /* ======================= step 2: year & form ======================= */
  (function initYears() {
    const sel = $('#yearSelect');
    sel.innerHTML = yearOptions();
    sel.value = String(S.year);
    sel.addEventListener('change', () => { S.year = Number(sel.value); setPreset('year', S.year); clearResults(); refresh(); });
    $('#yearPrev').addEventListener('click', () => { if (sel.selectedIndex > 0) { sel.selectedIndex--; sel.dispatchEvent(new Event('change')); } });
    $('#yearNext').addEventListener('click', () => { if (sel.selectedIndex < sel.options.length - 1) { sel.selectedIndex++; sel.dispatchEvent(new Event('change')); } });
  })();
  function defaultMcitRate() { return 0.02; }

  $('#formPick').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (S.form !== b.dataset.v) { S.form = b.dataset.v; setPreset('form', S.form); S.opts.method = 'itemized'; clearResults(); }
    refresh();
  });
  $('#vatSeg').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (S.salesType !== b.dataset.v) {
      S.salesType = b.dataset.v; setPreset('vat', S.salesType);
      if (S.salesType === 'VAT' && S.opts.method === 'eight') S.opts.method = 'itemized';
      clearResults();
    }
    refresh();
  });

  function renderOptions() {
    const box = $('#taxOptions');
    if (!S.form || !S.client) { box.innerHTML = ''; return; }
    const o = S.opts;
    const opt = (v, l, cur, dis) => '<option value="' + v + '"' + (String(cur) === String(v) ? ' selected' : '') + (dis ? ' disabled' : '') + '>' + l + '</option>';
    let html = '';
    if (S.form === '1701') {
      const vat = S.salesType === 'VAT';
      if (vat && o.method === 'eight') o.method = 'itemized';
      html += '<label class="field"><span>Tax method</span><select data-opt="method">' +
        opt('itemized', 'Graduated rates — Itemized deductions', o.method) +
        opt('osd', 'Graduated rates — OSD 40% of gross sales (BIR rule)', o.method) +
        opt('osd_gi', 'Graduated rates — OSD 40% of gross income (Tax Calc)', o.method) +
        opt('eight', '8% income tax option' + (vat ? ' (non-VAT only)' : ''), o.method, vat) + '</select></label>';
    } else {
      html += '<label class="field"><span>Deduction method</span><select data-opt="method">' + opt('itemized', 'Itemized deductions', o.method) + opt('osd', 'OSD (40% of gross income)', o.method) + '</select></label>';
      html += '<label class="field"><span>Income tax rate</span><select data-opt="taxType">' +
        opt('rcit25', 'RCIT 25% (or MCIT, whichever is higher)', o.taxType) + opt('rcit20', 'RCIT 20% (or MCIT) — NTI ≤ ₱5M & assets ≤ ₱100M', o.taxType) +
        opt('iaet10', 'Special rate 10% (proprietary educational / hospital)', o.taxType) + opt('nonres25', 'RCIT — Non-Resident 25%', o.taxType) + '</select></label>';
      if (window.AnnualEngine.TAX_TYPES[o.taxType].mcit) {
        html += '<label class="field"><span>MCIT rate</span><select data-opt="mcitRate">' + opt(0.02, '2% of gross income', o.mcitRate) + opt(0.01, '1% (Jul 2020 – Jun 2023)', o.mcitRate) + '</select></label>';
        html += '<label class="check"><input type="checkbox" data-opt="mcitExempt"' + (o.mcitExempt ? ' checked' : '') + ' /><span>New corporation<small>First 3 years — not yet subject to MCIT</small></span></label>';
      }
    }
    html += '<label class="field"><span>Seller of (entertainment limit)</span><select data-opt="sellerOf">' + opt('Goods', 'Goods — 0.5% of net sales', o.sellerOf) + opt('Services', 'Services — 1% of net revenue', o.sellerOf) + '</select></label>';
    box.innerHTML = html;
  }
  $('#taxOptions').addEventListener('change', (e) => {
    const k = e.target.dataset.opt; if (!k) return;
    const v = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    S.opts[k] = k === 'mcitRate' ? Number(v) : v;
    refresh();
    if (S.X) renderResults();
  });

  /* ======================= step 3: document slots ======================= */
  const ACCEPT_ANY = 'image/*,application/pdf,.pdf,.xlsx,.xls,.xlsm,.csv';
  function groups() {
    const vatF = S.salesType === 'VAT' ? '2550Q' : '2551Q';
    const qF = S.form === '1702RT' ? '1702Q' : '1701Q';
    const py = S.year - 1;
    const g = [
      { id: 'sales', title: 'Quarterly sales returns — ' + vatF, sub: 'One return per quarter, Q1 to Q4 ' + S.year + '. Photo, PDF or Excel.', badge: 'Required', cols: 'c4',
        slots: [1, 2, 3, 4].map((q) => ({ key: 'sales:' + q, t: 'Q' + q + ' ' + vatF, k: ['Jan–Mar', 'Apr–Jun', 'Jul–Sep', 'Oct–Dec'][q - 1] })) },
      { id: 'qitr', title: 'Quarterly income tax returns — ' + qF, sub: 'Q1 to Q3 ' + S.year + ' — for the tax payments and credits already made.', badge: S.form === '1702RT' ? 'Required' : 'Recommended', cols: 'c3',
        slots: [1, 2, 3].map((q) => ({ key: 'qitr:' + q, t: 'Q' + q + ' ' + qF, k: ['Jan–Mar', 'Apr–Jun', 'Jul–Sep'][q - 1] })) },
      { id: 'prior', title: 'Prior year — ' + py, sub: 'Basis for cost of sales, operating expenses and the comparative figures. One combined PDF (ITR + AFS) is fine.', badge: 'Required', cols: 'c2',
        slots: [{ key: 'priorItr', t: 'Annual ITR ' + py, k: S.form === '1702RT' ? 'BIR 1702-RT' : 'BIR 1701' }, { key: 'priorAfs', t: 'Audited FS ' + py, k: 'PDF' }] },
      { id: 'tpl', title: 'ITR Template', sub: 'The office Excel template for this client and year — it will be checked against the returns and corrected.', badge: 'Required', cols: 'c1',
        slots: [{ key: 'template', t: 'ITR Template ' + S.year, k: 'Excel (.xlsx)', accept: '.xlsx,.xlsm,.xls' }] },
      { id: 'wh', title: 'Withholding certificates', sub: 'BIR 2307 for the whole year — Excel summary, PDF or photos' + (S.form === '1701' ? '; BIR 2316 if the client also earns compensation' : '') + '.', badge: 'Credits', cols: S.form === '1701' ? 'c2' : 'c1',
        slots: [{ key: 'cwt', t: 'BIR 2307', k: 'all quarters' }].concat(S.form === '1701' ? [{ key: 'c2316', t: 'BIR 2316', k: 'compensation' }] : []) },
      { id: 'exp', title: 'Expenses', sub: 'Optional — receipts, ledgers or spreadsheets. If none, expenses come from the ITR Template or from the prior-year ITR/AFS.', badge: 'Optional', cols: 'c1',
        slots: [{ key: 'expenses', t: 'Expense documents', k: 'any layout' }] }
    ];
    return g;
  }
  const filesOf = (k) => S.files[k] || (S.files[k] = []);
  function kindOf(f) {
    const n = f.name.toLowerCase();
    if (/\.(xlsx|xls|xlsm|csv)$/.test(n) || /sheet|excel|csv/.test(f.type)) return 'xls';
    if (/\.pdf$/.test(n) || f.type === 'application/pdf') return 'pdf';
    return 'img';
  }
  function renderDocs() {
    if (!S.form) { $('#docGroups').innerHTML = ''; return; }
    let n = 0;
    $('#docGroups').innerHTML = groups().map((g) => {
      const filled = g.slots.some((s) => filesOf(s.key).length);
      return '<div class="dg' + (filled ? ' ok' : '') + '"><div class="dg-h"><div class="t"><span class="n">' + (++n) + '</span><div><b>' + esc(g.title) + '</b><small>' + esc(g.sub) + '</small></div></div><span class="' + (g.badge === 'Required' ? 'req' : 'opt') + '">' + g.badge + '</span></div>' +
        '<div class="slots ' + g.cols + '">' + g.slots.map((s) => {
          const fl = filesOf(s.key);
          return '<label class="slot' + (fl.length ? ' filled' : '') + '" data-slot="' + s.key + '"><input type="file" multiple accept="' + (s.accept || ACCEPT_ANY) + '" />' +
            '<div class="sl-h"><span class="sl-t">' + esc(s.t) + '</span><span class="sl-k">' + esc(s.k) + '</span></div>' +
            (fl.length ? '<ul class="sl-files">' + fl.map((x) => '<li><span class="fi s ' + x.kind + '">' + ({ pdf: 'PDF', img: 'IMG', xls: 'XLS' }[x.kind]) + '</span><span class="fn" title="' + esc(x.file.name) + '">' + esc(x.file.name) + '</span><button class="rm" data-rm="' + s.key + '|' + x.id + '" aria-label="Remove">&times;</button></li>').join('') + '</ul>' : '') +
            '<span class="sl-add">＋ <span>Drop or <u>browse</u></span></span></label>';
        }).join('') + '</div></div>';
    }).join('');
    const count = Object.keys(S.files).reduce((t, k) => t + S.files[k].length, 0);
    $('#docCount').textContent = count ? plural(count, 'file') : '';
  }
  function addFiles(key, list) {
    for (const f of Array.from(list || [])) {
      if (!/\.(pdf|png|jpe?g|webp|gif|heic|heif|xlsx|xls|xlsm|csv)$/i.test(f.name) && !/^image\//.test(f.type)) { toast('Skipped "' + f.name + '": use a photo, PDF or Excel file.', 'err'); continue; }
      if (f.size > 45 * 1048576) { toast('"' + f.name + '" is larger than 45 MB.', 'err'); continue; }
      if (key === 'template' && kindOf(f) !== 'xls') { toast('The ITR Template must be an Excel file.', 'err'); continue; }
      const item = { id: uid(), file: f, kind: kindOf(f) };
      if (key === 'template') S.files.template = [item]; else filesOf(key).push(item);
    }
    clearResults();
    refresh();
  }
  const dg = $('#docGroups');
  dg.addEventListener('click', (e) => {
    const rm = e.target.closest('[data-rm]');
    if (rm) {
      e.preventDefault(); e.stopPropagation();
      const [k, id] = rm.dataset.rm.split('|');
      S.files[k] = filesOf(k).filter((x) => x.id !== id);
      clearResults(); refresh();
    }
  });
  dg.addEventListener('change', (e) => {
    const slot = e.target.closest('[data-slot]');
    if (slot && e.target.type === 'file') { addFiles(slot.dataset.slot, e.target.files); }
  });
  ['dragenter', 'dragover'].forEach((ev) => dg.addEventListener(ev, (e) => { const s = e.target.closest('.slot'); if (s) { e.preventDefault(); s.classList.add('over'); } }));
  ['dragleave', 'drop'].forEach((ev) => dg.addEventListener(ev, (e) => { const s = e.target.closest('.slot'); if (s) { e.preventDefault(); s.classList.remove('over'); } }));
  dg.addEventListener('drop', (e) => { const s = e.target.closest('.slot'); if (s) addFiles(s.dataset.slot, e.dataTransfer.files); });
  function clearFiles() { S.files = {}; }

  function requirements() {
    const hard = [], soft = [];
    const salesQ = [1, 2, 3, 4].filter((q) => filesOf('sales:' + q).length);
    if (!salesQ.length) hard.push('the quarterly ' + (S.salesType === 'VAT' ? '2550Q' : '2551Q') + ' returns');
    else if (salesQ.length < 4) soft.push('missing ' + [1, 2, 3, 4].filter((q) => salesQ.indexOf(q) < 0).map((q) => 'Q' + q).join(', ') + ' sales return');
    if (!filesOf('priorItr').length && !filesOf('priorAfs').length) hard.push('the prior-year ITR / AFS');
    else if (!filesOf('priorAfs').length && !filesOf('priorItr').length) soft.push('prior-year AFS');
    if (!filesOf('template').length) hard.push('the ITR Template');
    const qi = [1, 2, 3].filter((q) => filesOf('qitr:' + q).length);
    if (S.form === '1702RT' && qi.length < 3) soft.push('missing ' + [1, 2, 3].filter((q) => qi.indexOf(q) < 0).map((q) => 'Q' + q).join(', ') + ' 1702Q');
    return { hard, soft };
  }

  /* ======================= refresh UI state ======================= */
  function refresh() {
    const s1 = !!S.client, s2 = s1 && !!S.form;
    const req = s2 ? requirements() : { hard: ['documents'], soft: [] };
    const s3 = s2 && !req.hard.length;
    $('#step2').classList.toggle('locked', !s1);
    $('#step3').classList.toggle('locked', !s2);
    $$('#formPick button').forEach((b) => {
      b.classList.toggle('on', b.dataset.v === S.form);
      const rec = S.client && ((S.client.type === 'Corporation') === (b.dataset.v === '1702RT'));
      let tag = $('.rec', b);
      if (rec && !tag) b.insertAdjacentHTML('beforeend', '<span class="rec">Matches taxpayer type</span>');
      if (!rec && tag) tag.remove();
    });
    $$('#vatSeg button').forEach((b) => {
      const on = b.dataset.v === S.salesType;
      b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on));
      let tag = $('.rec', b);
      const rec = S.client && S.client.vat === b.dataset.v;
      if (rec && !tag) b.insertAdjacentHTML('beforeend', '<span class="rec">Client record</span>');
      if (!rec && tag) tag.remove();
    });
    renderOptions();
    renderDocs();
    const hint = $('#analyzeHint');
    if (s2) {
      hint.className = 'hint foot-hint' + (req.hard.length ? ' bad' : req.soft.length ? ' bad' : '');
      hint.textContent = req.hard.length ? 'Upload ' + req.hard.join(', ') + ' to continue.' : req.soft.length ? 'You can continue, but: ' + req.soft.join('; ') + '.' : 'All set.';
    } else hint.textContent = '';
    $('#analyzeBtn').disabled = !s3 || S.busy;
    const st = [s1, s2, s3, !!S.X];
    const firstOpen = st.indexOf(false);
    $$('#stepper li').forEach((li, i) => { li.classList.toggle('done', st[i]); li.classList.toggle('active', i === firstOpen); });
  }

  /* ======================= file preparation ======================= */
  function abToB64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  async function compressImage(file) {
    const max = CFG.IMAGE_MAX_DIMENSION || 2000;
    let src;
    try { src = await createImageBitmap(file); }
    catch (e) {
      src = await new Promise((res, rej) => {
        const img = new Image();
        img.onload = () => res(img);
        img.onerror = () => rej(new Error('Cannot read image "' + file.name + '". Try JPG or PNG.'));
        img.src = URL.createObjectURL(file);
      });
    }
    const w = src.width, h = src.height;
    const scale = Math.min(1, max / Math.max(w, h));
    const c = document.createElement('canvas');
    c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(src, 0, 0, c.width, c.height);
    const url = c.toDataURL('image/jpeg', CFG.IMAGE_QUALITY || 0.85);
    return { mime: 'image/jpeg', data: url.split(',')[1] };
  }
  /** Large or long scanned PDFs are turned into page images in the browser (smaller upload, faster reading) */
  async function pdfToPages(file, onProgress) {
    await loadScript(LIBS.pdfjs);
    const pdfjs = window.pdfjsLib;
    pdfjs.GlobalWorkerOptions.workerSrc = LIBS.pdfjsWorker;
    const doc = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
    const n = Math.min(doc.numPages, CFG.PDF_MAX_PAGES || 40);
    const pages = [];
    for (let i = 1; i <= n; i++) {
      const page = await doc.getPage(i);
      const vp1 = page.getViewport({ scale: 1 });
      const scale = Math.min(2, (CFG.PDF_PAGE_WIDTH || 1400) / vp1.width);
      const vp = page.getViewport({ scale });
      const c = document.createElement('canvas');
      c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      pages.push(c.toDataURL('image/jpeg', 0.72).split(',')[1]);
      if (onProgress) onProgress(i, n);
    }
    return { pages, total: doc.numPages };
  }
  /** ITR Template: list every non-empty cell with its address (and formula), so Muse Spark can point to cells */
  function templateCells(wb) {
    const lines = [];
    wb.SheetNames.forEach((sn) => {
      const ws = wb.Sheets[sn];
      Object.keys(ws).filter((k) => k[0] !== '!').forEach((addr) => {
        const c = ws[addr];
        if (c == null || (c.v === '' && !c.f)) return;
        let v = c.v instanceof Date ? c.v.toISOString().slice(0, 10) : c.v;
        lines.push(sn + '!' + addr + (c.f ? ' [formula =' + c.f + ']' : '') + ' = ' + v);
      });
    });
    return lines.join('\n');
  }
  async function prepareFile(item, docType, onProgress) {
    const f = item.file;
    if (item.kind === 'xls') {
      await loadScript(LIBS.xlsx);
      const buf = await f.arrayBuffer();
      const wb = window.XLSX.read(buf, { type: 'array', cellDates: true, cellFormula: true });
      const text = docType === 'itr_template' ? templateCells(wb)
        : wb.SheetNames.map((n) => '### Sheet: ' + n + '\n' + window.XLSX.utils.sheet_to_csv(wb.Sheets[n], { blankrows: false })).join('\n\n');
      if (!text.replace(/[#\s,]|Sheet:/g, '')) throw new Error('The spreadsheet appears to be empty.');
      return { name: f.name, mime: f.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', text, data: f.size < 8 * 1048576 ? abToB64(buf) : '' };
    }
    if (item.kind === 'pdf') {
      if (f.size > (CFG.PDF_RASTER_OVER_MB || 6) * 1048576 || docType === 'prior_year') {
        try {
          const r = await pdfToPages(f, onProgress);
          if (r.pages.length) return { name: f.name, mime: 'application/pdf', pages: r.pages, pageCount: r.total };
        } catch (e) { /* fall back to the original PDF */ }
      }
      return { name: f.name, mime: 'application/pdf', data: abToB64(await f.arrayBuffer()) };
    }
    const img = await compressImage(f);
    return { name: f.name.replace(/\.(heic|heif|png|webp|gif)$/i, '.jpg'), mime: img.mime, data: img.data };
  }
  async function pool(items, limit, fn) {
    const out = new Array(items.length);
    let next = 0;
    const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) { const i = next++; out[i] = await fn(items[i], i); }
    });
    await Promise.all(workers);
    return out;
  }
  function progItem(label, sub) {
    const li = document.createElement('li');
    li.className = 'wait';
    li.innerHTML = '<span class="st"></span><div class="pl"><b>' + esc(label) + '</b><small>' + esc(sub) + '</small></div>';
    $('#progList').appendChild(li);
    return {
      run: (s) => { li.className = ''; if (s) $('small', li).textContent = s; },
      ok: (s) => { li.className = 'ok'; if (s) $('small', li).textContent = s; },
      err: (s) => { li.className = 'err'; $('small', li).textContent = s; }
    };
  }

  /* ======================= analyze ======================= */
  $('#analyzeBtn').addEventListener('click', analyze);
  $('#pmClose').addEventListener('click', () => closeModal('#progressModal'));

  const DOC_LABEL = { sales_vat: '2550Q', sales_nonvat: '2551Q', quarterly_itr: 'Quarterly ITR', prior_year: 'Prior-year ITR/AFS', itr_template: 'ITR Template', form_2307: 'BIR 2307', form_2316: 'BIR 2316', expenses: 'Expenses' };

  async function analyze() {
    if (S.busy) return;
    S.busy = true; refresh();
    $('#progList').innerHTML = '';
    $('#pmClose').hidden = true;
    $('.spark-orb').classList.remove('done');
    $('#pmTitle').textContent = 'Muse Spark is reading the year’s documents';
    $('#pmSub').textContent = 'Revenue, quarterly payments, prior-year statements, the ITR Template and withholding certificates…';
    openModal('#progressModal');

    const salesDoc = S.salesType === 'VAT' ? 'sales_vat' : 'sales_nonvat';
    const jobs = [];
    const push = (key, docType, label, extra) => filesOf(key).forEach((x) => jobs.push(Object.assign({ x, docType, label, key }, extra || {})));
    [1, 2, 3, 4].forEach((q) => push('sales:' + q, salesDoc, 'Q' + q + ' ' + DOC_LABEL[salesDoc], { quarter: q }));
    [1, 2, 3].forEach((q) => push('qitr:' + q, 'quarterly_itr', 'Q' + q + ' ' + (S.form === '1702RT' ? '1702Q' : '1701Q'), { quarter: q }));
    push('priorItr', 'prior_year', 'Prior-year annual ITR');
    push('priorAfs', 'prior_year', 'Prior-year AFS');
    push('template', 'itr_template', 'ITR Template');
    push('cwt', 'form_2307', 'BIR 2307');
    push('c2316', 'form_2316', 'BIR 2316');
    push('expenses', 'expenses', 'Expenses');
    jobs.forEach((j) => { j.p = progItem(j.x.file.name, j.label + ' · waiting'); });
    const compP = progItem('Compute ' + (S.form === '1702RT' ? '1702-RT' : '1701') + ' and financial statements', 'waiting');

    const common = { clientId: S.client.id, year: S.year, form: S.form, vat: S.salesType };
    const results = await pool(jobs, CFG.PARALLEL_UPLOADS || 3, async (j) => {
      try {
        j.p.run(j.label + ' · preparing file…');
        const file = await prepareFile(j.x, j.docType, (i, n) => j.p.run(j.label + ' · preparing page ' + i + ' of ' + n + '…'));
        j.p.run(j.label + ' · Muse Spark is reading' + (file.pages ? ' ' + file.pages.length + ' pages' : '') + '…');
        const d = await api('extract', Object.assign({ docType: j.docType, file, quarter: j.quarter || 0 }, common));
        j.p.ok(j.label + ' · ' + summarize(j.docType, d.data));
        return { ok: true, job: j, d };
      } catch (e) {
        j.p.err(e.message);
        return { ok: false, job: j, error: e.message };
      }
    });
    const ok = results.filter((r) => r.ok);
    const failed = results.filter((r) => !r.ok);
    const sales = ok.filter((r) => r.job.docType === salesDoc);
    if (!sales.length) {
      $('.spark-orb').classList.add('done');
      $('#pmTitle').textContent = 'Analysis stopped';
      $('#pmSub').textContent = 'No revenue could be read from the quarterly sales returns. Check the files and try again.';
      $('#pmClose').hidden = false; compP.err('Skipped');
      S.busy = false; refresh();
      return;
    }
    S.X = collect(ok, failed);
    S.adj = { lineOverrides: {} };
    S.tab = 'summary';
    try { renderResults(true); compP.ok('Done'); }
    catch (e) { compP.err(e.message); console.error(e); }
    $('.spark-orb').classList.add('done');
    await sleep(failed.length ? 1800 : 700);
    closeModal('#progressModal');
    S.busy = false; refresh();
    $('#step4').scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (failed.length) toast(plural(failed.length, 'file') + ' could not be read. See the Documents tab.', 'err', 6000);
  }

  function summarize(t, d) {
    if (t === 'sales_vat' || t === 'sales_nonvat') return 'Revenue ' + money(d.revenue);
    if (t === 'quarterly_itr') return 'Tax due ' + money(d.tax_due) + ' · paid ' + money(Math.max(0, d.tax_payable));
    if (t === 'prior_year') return 'Revenue ' + money(d.sci.revenue || d.itr.gross_sales) + ' · net income ' + money(d.sci.net_income);
    if (t === 'itr_template') return plural(d.checkpoints.length, 'checkpoint') + ' · ' + plural(d.operating_expenses.length, 'expense line');
    if (t === 'form_2307') return plural(d.certificates.length, 'certificate') + ' · CWT ' + money(d.total_tax_withheld);
    if (t === 'form_2316') return plural(d.certificates.length, 'certificate') + ' · withheld ' + money(d.total_withheld);
    if (t === 'expenses') return plural(d.items.length, 'item') + ' · ' + money(d.document_total);
    return 'done';
  }

  /** merge several prior-year extractions (e.g. ITR PDF + AFS PDF): keep the first non-empty value for each field */
  function mergePrior(a, b) {
    if (a == null) return b;
    if (b == null) return a;
    if (Array.isArray(a) || Array.isArray(b)) return (Array.isArray(a) && a.length >= (b || []).length) ? a : b;
    if (typeof a === 'object' && typeof b === 'object') {
      const out = Object.assign({}, a);
      Object.keys(b).forEach((k) => { out[k] = mergePrior(a[k], b[k]); });
      return out;
    }
    if (typeof a === 'number') return a !== 0 ? a : b;
    if (typeof a === 'string') return a !== '' ? a : b;
    return a;
  }

  function collect(ok, failed) {
    const by = (t) => ok.filter((r) => r.job.docType === t);
    const W = [];
    const salesDoc = S.salesType === 'VAT' ? 'sales_vat' : 'sales_nonvat';
    const sales = by(salesDoc).map((r) => {
      const d = Object.assign({}, r.d.data, { file: r.d.fileName });
      if (d.quarter && Number(d.quarter) !== r.job.quarter) W.push(r.d.fileName + ' looks like Q' + d.quarter + ' but was uploaded under Q' + r.job.quarter + ' (Q' + r.job.quarter + ' was used).');
      if (d.year && Number(d.year) !== S.year) W.push(r.d.fileName + ' is for ' + d.year + ', not ' + S.year + '.');
      if (d.taxpayer_tin && S.client.tin && tinDigits(d.taxpayer_tin) !== tinDigits(S.client.tin)) W.push('TIN on ' + r.d.fileName + ' (' + d.taxpayer_tin + ') does not match the client.');
      d.quarter = r.job.quarter;
      return d;
    });
    const qitrs = by('quarterly_itr').map((r) => Object.assign({}, r.d.data, { quarter: r.job.quarter, file: r.d.fileName }));
    let prior = null;
    by('prior_year').forEach((r) => { prior = mergePrior(prior, r.d.data); });
    const tplRes = by('itr_template')[0];
    const certs2307 = [], certs2316 = [], items = [];
    by('form_2307').forEach((r) => r.d.data.certificates.forEach((c) => certs2307.push(Object.assign({ file: r.d.fileName }, c))));
    by('form_2316').forEach((r) => r.d.data.certificates.forEach((c) => certs2316.push(Object.assign({ file: r.d.fileName }, c))));
    by('expenses').forEach((r) => r.d.data.items.forEach((i) => items.push(Object.assign({ file: r.d.fileName }, i))));
    ok.forEach((r) => { if (r.d.data && r.d.data.confidence === 'low') W.push('Muse Spark has low confidence reading ' + r.d.fileName + (r.d.data.notes ? ': ' + r.d.data.notes : '.')); });
    if (prior && prior.year && Number(prior.year) !== S.year - 1) W.push('The prior-year documents appear to be for ' + prior.year + ', not ' + (S.year - 1) + '.');
    failed.forEach((f) => W.push('Could not read ' + f.job.x.file.name + ': ' + f.error));
    return {
      sales, qitrs, prior, template: tplRes ? tplRes.d.data : null, templateFile: filesOf('template')[0] ? filesOf('template')[0].file : null,
      certs2307, certs2316, expenseItems: items, warnings: W,
      docs: ok.map((r) => ({ type: r.job.docType, label: r.job.label, file: r.d.fileName, data: r.d.data, url: r.d.archivedUrl })).concat(failed.map((f) => ({ type: f.job.docType, label: f.job.label, file: f.job.x.file.name, error: f.error })))
    };
  }
  const tinDigits = (t) => String(t || '').replace(/\D/g, '').slice(0, 9);

  /* ======================= computation ======================= */
  function engineInput() {
    const X = S.X;
    return {
      client: { name: S.client.name, tin: S.client.tin, address: S.client.address, lob: S.client.lob, rdo: S.client.rdo, type: S.form === '1702RT' ? 'Corporation' : 'Individual', vat: S.salesType, sellerOf: S.opts.sellerOf },
      year: S.year, form: S.form, vat: S.salesType, opts: Object.assign({}, S.opts),
      sales: X.sales, quarterlyItrs: X.qitrs, prior: X.prior || {}, template: X.template,
      expenseItems: X.expenseItems, certs2307: X.certs2307, certs2316: X.certs2316, adj: S.adj
    };
  }
  function model() {
    const M = window.AnnualEngine.build(engineInput());
    M.warnings = (S.X.warnings || []).concat(M.warnings);
    M.user = S.user; M.docs = S.X.docs; M.templateFile = S.X.templateFile;
    M.client = Object.assign({}, M.client, { rdo: S.client.rdo, address: S.client.address || (S.X.prior && S.X.prior.entity && S.X.prior.entity.address) || '' });
    M.entity = (S.X.prior && S.X.prior.entity) || {};
    M.methodLabel = methodLabel();
    M.templateLoans = (S.X.template && S.X.template.loans) || [];
    M.templateTaxesLicenses = (S.X.template && S.X.template.taxes_licenses) || [];
    return M;
  }
  function methodLabel() {
    const o = S.opts;
    if (S.form === '1701') return { itemized: 'Graduated rates — Itemized deductions', osd: 'Graduated rates — OSD 40% of gross sales', osd_gi: 'Graduated rates — OSD 40% of gross income', eight: '8% income tax option' }[o.method];
    const tt = window.AnnualEngine.TAX_TYPES[o.taxType];
    return (o.method === 'osd' ? 'OSD' : 'Itemized deductions') + ' · ' + tt.label + (tt.mcit && !o.mcitExempt ? ' vs MCIT ' + (o.mcitRate * 100) + '%' : '');
  }

  /* ======================= results ======================= */
  $('#newBtn').addEventListener('click', () => resetAll(false));
  function clearResults() { S.X = null; $('#step4').hidden = true; $('#results').innerHTML = ''; }
  function resetAll(full) {
    clearResults(); clearFiles();
    if (full) S.clients = [];
    selectClient(null);
    refresh();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  $('#tabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-tab]'); if (!b) return;
    S.tab = b.dataset.tab; renderResults();
  });

  const inp = (attr, val, edited, w) => '<input class="num-input' + (edited ? ' edited' : '') + '" inputmode="decimal" ' + attr + ' value="' + plain(val) + '"' + (w ? ' style="width:' + w + 'px"' : '') + ' />';
  const row = (cls, label, cur, prior, basis, ind) => '<tr class="' + (cls || '') + '"><td' + (ind ? ' class="i"' : '') + '>' + label + '</td><td class="a">' + cur + '</td><td class="a">' + (prior == null ? '' : prior) + '</td><td class="b">' + (basis || '') + '</td></tr>';
  const mny = (n) => (n == null || n === '' ? '—' : money(n));

  function renderResults(animate) {
    if (!S.X) return;
    const M = model();
    S.M = M;
    const focusKey = document.activeElement && document.activeElement.getAttribute && (document.activeElement.getAttribute('data-k'));
    $('#step4').hidden = false;
    $('#resultTitle').textContent = (S.form === '1702RT' ? 'BIR 1702-RT' : 'BIR 1701') + ' · ' + S.year + ' · ' + S.client.name;
    const nBad = M.corrections.filter((c) => !c.ok).length;
    $$('#tabs [data-tab]').forEach((b) => {
      b.classList.toggle('on', b.dataset.tab === (S.tab || 'summary'));
      const bd = $('.badge', b); if (bd) bd.remove();
      if (b.dataset.tab === 'tpl' && nBad) b.insertAdjacentHTML('beforeend', '<span class="badge">' + nBad + '</span>');
      if (b.dataset.tab === 'summary' && M.warnings.length) b.insertAdjacentHTML('beforeend', '<span class="badge">' + M.warnings.length + '</span>');
    });
    const T = M.tax, FS = M.fs, E = M.expenses;
    const tab = S.tab || 'summary';
    let h = '';
    const yr = S.year, py = S.year - 1;
    const head = (title) => '<div class="fs-head"><b>' + esc(S.client.name) + '</b><small>' + title + '</small></div>';
    const thead = '<tr><th></th><th>' + yr + '</th><th>' + py + '</th><th style="text-align:left">Basis / notes</th></tr>';

    if (tab === 'summary') {
      const kpi = (label, val, sub, hl) => '<div class="kpi' + (hl ? ' hl' : '') + '"><small>' + label + '</small><b data-count="' + val + '">' + money(animate ? 0 : val) + '</b>' + (sub ? '<em>' + sub + '</em>' : '') + '</div>';
      h += '<div class="kpis k5">' + kpi('Revenue', M.revenueTotal, 'Q1–Q4 ' + yr) + kpi('Cost of sales', E.cosTotal, E.cos.length + ' lines') + kpi('Operating expenses', E.opexTotal, E.opex.length + ' lines') +
        kpi('Net income', FS.sci.netIncome, 'after tax') + kpi(T.overpayment ? 'Overpayment' : 'Tax still due', T.overpayment || T.payable, 'Tax due ' + money(T.taxDue), true) + '</div>';
      if (M.warnings.length) h += '<div class="alerts">' + M.warnings.map((w) => '<div class="alert">' + esc(w) + '</div>').join('') + '</div>';
      // expense basis
      const av = E.available;
      h += '<div class="adj-grid"><div class="h"><b>Basis of cost of sales &amp; operating expenses</b><small>' +
        ({ documents: 'From the uploaded expense documents.', template: 'From the operating expenses listed in the ITR Template.', prior: 'Prior-year (' + py + ') ITR/AFS line items re-scaled to this year’s revenue.', none: 'No basis available.' }[E.basis]) + '</small></div>' +
        '<div class="basis-pick" style="grid-column:1/-1">' + [['documents', 'Expense documents'], ['template', 'ITR Template'], ['prior', 'Prior-year ' + py + ' rates']].map(([k, l]) => '<button data-basis="' + k + '"' + (E.basis === k ? ' class="on"' : '') + (av[k] ? '' : ' disabled') + '>' + l + '</button>').join('') + '</div>';
      if (E.basis === 'prior') {
        h += '<label class="adj"><span>Cost of sales — % of revenue</span><input data-k="cosPct" data-pct value="' + E.cosPct + '"' + (Math.abs(E.cosPct - E.priorCosPct) > 0.004 ? ' class="changed"' : '') + ' /><small>' + py + ': ' + E.priorCosPct + '% · ' + money(E.cosTotal) + '</small></label>' +
          '<label class="adj"><span>Operating expenses — % of revenue</span><input data-k="opexPct" data-pct value="' + E.opexPct + '"' + (Math.abs(E.opexPct - E.priorOpexPct) > 0.004 ? ' class="changed"' : '') + ' /><small>' + py + ': ' + E.priorOpexPct + '% · ' + money(E.opexTotal) + '</small></label>' +
          '<div class="adj"><span>&nbsp;</span><button class="btn btn-outline" data-reset-pct>Use ' + py + ' rates</button><small>Depreciation, interest and taxes &amp; licenses use actual figures.</small></div>';
      }
      h += '</div>';
      h += '<div class="panel"><div class="panel-h"><h4>At a glance</h4><small>' + esc(M.methodLabel) + '</small></div><table class="tbl">' +
        [['Revenue (sum of 4 quarterly returns)', money(M.revenueTotal)], ['Gross profit', money(FS.sci.grossProfit)], ['Income before income tax', money(FS.sci.ibt)], ['Taxable income', money(T.taxable)],
          ['Income tax due', money(T.taxDue)], ['Less: credits (quarterly payments, CWT, excess)', money(T.totalCredits)], [T.overpayment ? 'Overpayment' : 'Tax still due', money(T.overpayment || T.payable)],
          ['Total assets (Dec 31, ' + yr + ')', money(FS.sfp.totalAssets)], ['Cash — balancing figure', money(FS.sfp.cur.cash)],
          ['ITR Template check', M.corrections.length ? (M.corrections.length - nBad) + ' of ' + M.corrections.length + ' figures match' + (nBad ? ' · ' + nBad + ' corrected' : ' ✓') : 'No checkpoints found']]
          .map(([a, b]) => '<tr><td>' + a + '</td><td class="amt">' + b + '</td></tr>').join('') + '</table></div>';
    }

    if (tab === 'tax') {
      h += '<div class="panel"><div class="panel-h"><h4>' + esc(T.formLabel) + ' — Annual Income Tax · ' + yr + '</h4><small>' + esc(T.schedule) + '</small></div><div class="scroll-x"><table class="tbl">';
      T.lines.forEach((l) => {
        const cls = l.type === 'total' ? 'tot' : l.type === 'sub' ? 'sub' : '';
        const amt = l.type === 'rate' ? l.amount + '%' : l.edit ? inp('data-k="' + l.edit + '" data-adj', l.amount) : money(l.amount);
        h += '<tr class="' + cls + '"><td>' + esc(l.label) + (l.note ? '<span class="note">' + esc(l.note) + '</span>' : '') + '</td><td class="amt">' + amt + '</td></tr>';
      });
      h += '<tr class="sec"><td colspan="2">Less: Tax credits / payments</td></tr>';
      T.credits.forEach((c) => { h += '<tr><td class="ind">' + esc(c.label) + (c.note ? '<span class="note">' + esc(c.note) + '</span>' : '') + '</td><td class="amt">' + inp('data-k="' + c.key + '" data-adj', c.amount) + '</td></tr>'; });
      h += '<tr class="tot"><td>Total tax credits / payments</td><td class="amt">' + money(T.totalCredits) + '</td></tr>';
      h += '<tr class="grand"><td>' + (T.overpayment ? 'Overpayment' : 'Tax Still Due') + '</td><td class="amt">' + money(T.overpayment || T.payable) + '</td></tr>';
      h += '</table></div></div>';
      if (S.X.certs2307.length) h += '<div class="panel"><div class="panel-h"><h4>BIR 2307 — creditable tax withheld</h4><small>' + plural(S.X.certs2307.length, 'certificate') + '</small></div><div class="scroll-x"><table class="tbl">' +
        S.X.certs2307.map((c) => '<tr><td><b>' + esc(c.payor_name || '—') + '</b><span class="note">TIN ' + esc(c.payor_tin || '—') + ' · ' + (c.quarter ? 'Q' + c.quarter + ' ' : '') + esc([c.period_from, c.period_to].filter(Boolean).join(' – ')) + ' · income ' + money(c.total_income) + '</span></td><td class="amt">' + money(c.total_tax_withheld) + '</td></tr>').join('') + '</table></div></div>';
      if (S.X.certs2316.length) h += '<div class="panel"><div class="panel-h"><h4>BIR 2316 — compensation</h4></div><table class="tbl">' +
        S.X.certs2316.map((c) => '<tr><td><b>' + esc(c.employer_name) + '</b><span class="note">Taxable compensation ' + money(c.taxable_compensation) + '</span></td><td class="amt">' + money(c.tax_withheld) + '</td></tr>').join('') + '</table></div>';
    }

    if (tab === 'is') {
      const s = FS.sci;
      h += head(s.title + ' · For the years ended December 31, ' + yr + ' and ' + py);
      h += '<div class="scroll-x"><table class="fst">' + thead;
      h += row('sub', FS.isProp ? 'REVENUES' : 'SALES', money(s.revenue), mny(s.priorRevenue), 'Sum of Q1–Q4 ' + (S.salesType === 'VAT' ? '2550Q (net of VAT)' : '2551Q'));
      const lineRow = (cat, l) => {
        const k = cat + '|' + window.AnnualEngine.norm(l.label);
        const el = E[cat === 'cos' ? 'cos' : 'opex'].find((x) => x.key === k);
        const editable = !!el;
        return '<tr class="' + (l.amount ? '' : 'zero') + '"><td class="i">' + esc(l.label) + '</td><td class="a">' + (editable ? inp('data-k="' + esc(k) + '" data-line', l.amount, el.edited, 130) : money(l.amount)) + '</td><td class="a">' + mny(l.prior) + '</td><td class="b">' + esc(l.basis || '') + '</td></tr>';
      };
      if (s.cos.length) {
        h += '<tr class="sec"><td colspan="4">Cost of sales</td></tr>' + s.cos.map((l) => lineRow('cos', l)).join('');
        h += row('tot', 'Total cost of sales', money(s.cosTotal), mny(s.priorCos));
        h += row('sub', 'GROSS PROFIT', money(s.grossProfit), mny(s.priorGross));
      }
      h += '<tr class="sec"><td colspan="4">Operating expenses</td></tr>' + s.opex.map((l) => lineRow('opex', l)).join('');
      h += row('tot', 'Total operating expenses', money(s.opexTotal), mny(s.priorOpex));
      if (!FS.isProp || s.otherIncome) { h += row('sub', 'OPERATING INCOME', money(s.operatingIncome), null); h += row('', 'Other income', inp('data-k="otherIncome" data-adj', s.otherIncome, false, 130), mny(s.priorOtherIncome)); }
      h += row('sub', 'INCOME BEFORE INCOME TAX', money(s.ibt), mny(s.priorIbt));
      h += row('', 'Provision for income tax', money(s.provision), mny(s.priorProvision), esc(T.formLabel + ' — ' + (T.schedule || '')));
      h += row('grand', 'NET INCOME', money(s.netIncome), mny(s.priorNetIncome));
      h += '</table></div><p class="fs-note">Edit any expense amount to override it. Prior-year figures are from the ' + py + ' AFS / ITR.</p>';
    }

    if (tab === 'bs') {
      const B = FS.sfp, c = B.cur, p = B.prior, L = B.labels || {};
      const lab = (k, d) => esc(L[k] || d);
      const a = (k, label) => '<label class="adj"><span>' + label + '</span><input data-k="' + k + '" data-adj value="' + plain(S.adj[k] != null ? S.adj[k] : defaultAdj(k, M)) + '"' + (S.adj[k] != null ? ' class="changed"' : '') + ' /><small>' + py + ': ' + money(priorAdj(k, M)) + '</small></label>';
      h += '<div class="adj-grid"><div class="h"><b>Year-end balances (Dec 31, ' + yr + ')</b><small>Carried over from ' + py + ' unless you change them. Cash is the balancing figure.</small></div>' +
        a('receivables', 'Receivables') + a('inventoryEnd', 'Inventory (ending)') + a('prepaid', 'Prepaid expenses') + a('otherCA', 'Other current assets') + a('otherNCA', 'Other non-current assets') +
        a('payables', 'Accounts payable') + a('accrued', 'Accrued expenses') + a('loans', 'Loans payable') + a('ppeAdditions', 'PPE additions in ' + yr) + a('depreciation', 'Depreciation for ' + yr) +
        (FS.isProp ? a('drawings', 'Drawings') : a('dividends', 'Dividends declared') + a('capitalAdditions', 'Additional capital')) + '</div>';
      h += head(B.title + ' · As of December 31, ' + yr + ' and ' + py);
      h += '<div class="scroll-x"><table class="fst">' + thead;
      h += '<tr class="sec"><td colspan="4">Assets</td></tr>';
      const r = (k, d, basis) => (c[k] || p[k]) ? row('', lab(k, d), money(c[k]), mny(p[k]), basis, true) : '';
      h += r('cash', 'Cash and cash equivalents', 'Balancing figure (total liabilities + equity − other assets)') + r('receivables', 'Accounts receivable') + r('inventory', 'Inventory') + r('prepaid', 'Prepaid expenses') + r('other_ca', 'Other current assets');
      h += row('tot', 'Total current assets', money(B.totalCA), mny(B.priorCA));
      if (B.hasGrossPPE) { h += r('ppe_cost', 'Property and equipment — cost', 'Prior cost + additions') + row('', 'Less: Accumulated depreciation', '(' + plain(c.accumulated_depreciation) + ')', '(' + plain(p.accumulated_depreciation) + ')', 'Prior + depreciation ' + money(FS.dep), true); }
      h += row('', lab('ppe_net', 'Property and equipment — net'), money(c.ppe_net), mny(p.ppe_net), B.hasGrossPPE ? '' : 'Prior net + additions − depreciation', true);
      h += r('other_nca', 'Other non-current assets');
      h += row('tot', 'Total non-current assets', money(B.totalNCA), mny(B.priorNCA));
      h += row('grand', 'TOTAL ASSETS', money(B.totalAssets), mny(B.priorAssets));
      h += '<tr class="sec"><td colspan="4">Liabilities</td></tr>';
      h += r('payables', 'Accounts payable', S.adj.payables == null ? 'Prior × revenue growth' : 'Entered') + r('accrued', 'Accrued expenses payable', S.adj.accrued == null ? 'Prior × revenue growth' : 'Entered') + r('vat_payable', 'VAT payable', S.salesType === 'VAT' ? 'Q4 2550Q' : '') +
        row('', lab('income_tax_payable', 'Income tax payable'), money(c.income_tax_payable), mny(p.income_tax_payable), 'Tax still due per ' + T.formLabel, true) + r('other_cl', 'Other current liabilities') + r('loans', 'Loans payable', S.X.template && (S.X.template.loans || []).length ? 'Outstanding per ITR Template' : '') + r('other_ncl', 'Other non-current liabilities');
      h += row('tot', 'Total liabilities', money(B.totalLiab), mny(B.priorLiab));
      const eq = B.equity;
      h += '<tr class="sec"><td colspan="4">' + (FS.isProp ? 'Owner’s equity' : 'Stockholders’ equity') + '</td></tr>';
      if (FS.isProp) {
        h += row('', 'Balance, beginning of year', money(eq.begin), mny(M.prior.equity.beginning), '', true) + (eq.capitalAdd ? row('', 'Additional capital', money(eq.capitalAdd), null, '', true) : '') +
          row('', 'Drawings', '(' + plain(eq.drawings) + ')', M.prior.equity.drawings ? '(' + plain(M.prior.equity.drawings) + ')' : '—', S.adj.drawings == null ? 'Same as ' + py : 'Entered', true) + row('', 'Net income', money(eq.netIncome), mny(M.prior.sci.net_income), '', true);
      } else {
        const lines = (eq.lines || []).length ? eq.lines : [{ label: 'Equity', amount: eq.begin, kind: 'retained_earnings' }];
        let reDone = false;
        lines.forEach((l) => {
          let cur = Number(l.amount) || 0;
          if (!reDone && l.kind === 'retained_earnings') { cur = cur + eq.netIncome - eq.dividends; reDone = true; }
          if (l.kind === 'capital' && eq.capitalAdd) cur += eq.capitalAdd;
          h += row('', esc(l.label), money(cur), mny(l.amount), l.kind === 'retained_earnings' ? '+ net income ' + money(eq.netIncome) + (eq.dividends ? ' − dividends ' + money(eq.dividends) : '') : '', true);
        });
        if (!reDone) h += row('', 'Net income for the year', money(eq.netIncome), null, '', true);
      }
      h += row('tot', 'Total equity', money(eq.end), mny(B.priorEquity));
      h += row('grand', 'TOTAL LIABILITIES AND EQUITY', money(B.totalLE), mny(B.priorLE));
      h += '</table></div><p class="fs-note">' + (Math.abs(B.totalAssets - B.totalLE) < 0.01 ? '<span class="chk-ok">✓ Balanced</span>' : '<span class="chk-bad">Not balanced</span>') + ' · Cash is computed as the balancing figure, as in the office samples.</p>';
    }

    if (tab === 'cf') {
      const C = FS.scf;
      h += head('STATEMENTS OF CASH FLOWS · For the year ended December 31, ' + yr + ' (indirect method)');
      h += '<div class="scroll-x"><table class="fst"><tr><th></th><th>' + yr + '</th><th>' + py + '</th><th></th></tr>';
      h += '<tr class="sec"><td colspan="4">Cash flows from operating activities</td></tr>' + C.op.map((l) => row('', esc(l.label), money(l.amount), null, '', true)).join('');
      h += row('tot', 'Net cash provided by (used in) operating activities', money(C.cfo), mny(C.prior.cfo));
      h += '<tr class="sec"><td colspan="4">Cash flows from investing activities</td></tr>' + (C.inv.length ? C.inv.map((l) => row('', esc(l.label), money(l.amount), null, '', true)).join('') : row('zero', 'None', money(0), null, '', true));
      h += row('tot', 'Net cash provided by (used in) investing activities', money(C.cfi), mny(C.prior.cfi));
      h += '<tr class="sec"><td colspan="4">Cash flows from financing activities</td></tr>' + (C.fin.length ? C.fin.map((l) => row('', esc(l.label), money(l.amount), null, '', true)).join('') : row('zero', 'None', money(0), null, '', true));
      h += row('tot', 'Net cash provided by (used in) financing activities', money(C.cff), mny(C.prior.cff));
      h += row('sub', 'NET INCREASE (DECREASE) IN CASH', money(C.netChange), mny(C.prior.net_change));
      h += row('', 'Cash at beginning of year', money(C.cashBegin), mny(C.prior.cash_beginning));
      h += row('grand', 'CASH AT END OF YEAR', money(C.cashEnd), mny(C.prior.cash_end));
      h += '</table></div><p class="fs-note">' + (C.reconciles ? '<span class="chk-ok">✓ Agrees with the balance sheet cash</span>' : '<span class="chk-bad">Does not agree with the balance sheet cash — see Summary</span>') + '</p>';
    }

    if (tab === 'tpl') {
      const cs = M.corrections;
      h += '<div class="panel"><div class="panel-h"><h4>ITR Template vs the returns</h4><small>' + (cs.length ? (cs.length - nBad) + ' of ' + cs.length + ' match' : 'No checkpoints found in the template') + '</small></div><div class="scroll-x"><table class="tbl">' +
        '<tr class="sec"><td>Item</td><td class="amt">Template</td><td class="amt">Per returns</td><td class="amt">Difference</td><td>Cell</td></tr>' +
        cs.map((c) => '<tr><td>' + (c.ok ? '<span class="chk-ok">✓</span> ' : '<span class="chk-bad">✕</span> ') + esc(c.label) + '<span class="note">' + esc(c.source) + '</span></td><td class="amt">' + money(c.template) + '</td><td class="amt">' + money(c.verified) + '</td><td class="amt">' + (c.ok ? '—' : money(c.diff)) + '</td><td>' + esc(c.cell) + (c.isFormula ? ' <span class="note">formula' + (c.inputCell ? ' → fix ' + esc(c.inputCell) : '') + '</span>' : '') + '</td></tr>').join('') +
        '</table></div></div><p class="fs-note">The corrected template keeps your file’s layout: wrong input cells are fixed and highlighted, and two sheets are added — “Corrections” and “Annual ITR Summary”.</p>';
    }

    if (tab === 'docs') {
      h += '<div class="panel"><div class="panel-h"><h4>Documents read by Muse Spark</h4><small>' + plural(M.docs.length, 'file') + '</small></div><div class="scroll-x"><table class="tbl">' +
        M.docs.map((d) => '<tr><td><b>' + esc(d.label) + '</b> · ' + esc(d.file) + (d.data && d.data.confidence ? ' <span class="conf ' + esc(d.data.confidence) + '">' + esc(d.data.confidence) + '</span>' : '') +
          (d.error ? '<span class="note" style="color:var(--err)">' + esc(d.error) + '</span>' : d.data && d.data.notes ? '<span class="note">' + esc(d.data.notes) + '</span>' : '') +
          (d.url ? '<span class="note"><a href="' + esc(d.url) + '" target="_blank" rel="noopener">Open archived copy</a></span>' : '') + '</td><td class="amt">' + (d.data ? esc(summarize(d.type, d.data)) : '') + '</td></tr>').join('') + '</table></div></div>';
    }

    $('#results').innerHTML = h;
    if (animate) countUp();
    if (focusKey) { const el = $('[data-k="' + CSS.escape(focusKey) + '"]'); if (el) { el.focus(); if (el.select) el.select(); } }
  }

  function defaultAdj(k, M) {
    const c = M.fs.sfp.cur, e = M.fs.sfp.equity;
    return { receivables: c.receivables, inventoryEnd: c.inventory, prepaid: c.prepaid, otherCA: c.other_ca, otherNCA: c.other_nca, payables: c.payables, accrued: c.accrued, loans: c.loans, ppeAdditions: 0, depreciation: M.fs.dep, drawings: e.drawings, dividends: e.dividends, capitalAdditions: e.capitalAdd }[k] || 0;
  }
  function priorAdj(k, M) {
    const p = M.fs.sfp.prior, P = M.prior;
    return { receivables: p.receivables, inventoryEnd: p.inventory, prepaid: p.prepaid, otherCA: p.other_ca, otherNCA: p.other_nca, payables: p.payables, accrued: p.accrued, loans: p.loans, ppeAdditions: 0, depreciation: P.depreciation, drawings: P.equity.drawings, dividends: P.equity.dividends, capitalAdditions: 0 }[k] || 0;
  }

  function countUp() {
    $$('.kpi b[data-count]').forEach((el) => {
      const target = Number(el.dataset.count), t0 = performance.now(), dur = 900;
      const step = (t) => { const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 3); el.textContent = money(target * e); if (p < 1) requestAnimationFrame(step); else el.textContent = money(target); };
      requestAnimationFrame(step);
    });
  }

  const res = $('#results');
  res.addEventListener('click', (e) => {
    const b = e.target.closest('[data-basis]');
    if (b && !b.disabled) { S.adj.expenseBasis = b.dataset.basis; S.adj.lineOverrides = {}; renderResults(); return; }
    if (e.target.closest('[data-reset-pct]')) { delete S.adj.cosPct; delete S.adj.opexPct; S.adj.lineOverrides = {}; renderResults(); }
  });
  res.addEventListener('focusin', (e) => { if (e.target.matches('input')) setTimeout(() => e.target.select && e.target.select(), 0); });
  res.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.matches('input')) e.target.blur(); });
  res.addEventListener('change', (e) => {
    const el = e.target, k = el.dataset.k; if (!k) return;
    const v = parseNum(el.value);
    if (el.hasAttribute('data-pct')) {
      if (v < 0 || v > 100) { toast('Enter a rate between 0 and 100%.', 'err'); renderResults(); return; }
      S.adj[k] = v; S.adj.lineOverrides = {};
    } else if (el.hasAttribute('data-line')) {
      const cur = S.M && S.M.expenses[k.split('|')[0]].find((x) => x.key === k);
      if (cur && cur.edited && Math.abs((cur.extracted || 0) - v) < 0.005) delete S.adj.lineOverrides[k]; else S.adj.lineOverrides[k] = v;
    } else if (el.hasAttribute('data-adj')) S.adj[k] = v;
    setTimeout(() => renderResults(), 0);
  });

  /* ======================= downloads ======================= */
  const safe = (s) => String(s).replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').slice(0, 40);
  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  async function makeFile(kind, M) {
    const base = safe(S.client.name) + '_' + S.year;
    if (kind === 'itr' || kind === 'fs' || kind === 'tpl') await loadScript(LIBS.exceljs);
    if (kind === 'itr') return { blob: await window.ITRExcel.computation(M), name: (S.form === '1702RT' ? '1702RT' : '1701') + '_Computation_' + base + '.xlsx' };
    if (kind === 'fs') return { blob: await window.ITRExcel.statements(M), name: 'Financial_Statements_' + base + '.xlsx' };
    if (kind === 'tpl') {
      const buf = M.templateFile ? await M.templateFile.arrayBuffer() : null;
      return { blob: await window.ITRExcel.correctedTemplate(M, buf), name: 'ITR_Template_' + base + '_corrected.xlsx' };
    }
    if (kind === 'docx') { await loadScript(LIBS.docx); return { blob: await window.AFSDoc.docx(M), name: 'AFS_' + base + '.docx' }; }
    if (kind === 'pdf') { await loadScript(LIBS.pdfmake); await loadScript(LIBS.pdfmakeFonts); return { blob: await window.AFSDoc.pdf(M), name: 'AFS_' + base + '.pdf' }; }
  }
  let reported = false;
  $('#dlBar').addEventListener('click', async (e) => {
    const b = e.target.closest('[data-dl]'); if (!b || !S.X) return;
    const kind = b.dataset.dl;
    b.classList.add('loading');
    try {
      const M = model();
      const kinds = kind === 'all' ? ['itr', 'fs', 'tpl', 'docx', 'pdf'] : [kind];
      for (const k of kinds) { const f = await makeFile(k, M); saveBlob(f.blob, f.name); if (kinds.length > 1) await sleep(500); }
      toast(kinds.length > 1 ? '5 files downloaded.' : 'Downloaded.', 'ok');
      if (!reported) {
        reported = true;
        api('saveReport', { clientId: S.client.id, report: { form: S.form, year: S.year, method: M.methodLabel, revenue: M.revenueTotal, costOfSales: M.expenses.cosTotal, opex: M.expenses.opexTotal, taxable: M.tax.taxable, taxDue: M.tax.taxDue, credits: M.tax.totalCredits, taxPayable: M.tax.payable, netIncome: M.fs.sci.netIncome, totalAssets: M.fs.sfp.totalAssets, expenseBasis: M.expenses.basis, outputs: kinds.join(',') } }).catch(() => {});
      }
    } catch (err) {
      console.error(err);
      toast('Could not create the file: ' + err.message, 'err', 6000);
    } finally { b.classList.remove('loading'); }
  });

  /* ======================= boot ======================= */
  async function boot() {
    runIntro();
    refresh();
    showLanding();
    if (S.token && S.user) loadClients();
    if (!apiUrl()) setTimeout(openSettings, introDone ? 0 : 3000);
  }
  window.ITRApp = { state: S, model: () => model(), renderResults };
  boot();
})();
