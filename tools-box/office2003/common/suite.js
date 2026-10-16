/* Office 2003 suite services, shared by the start page and the apps (classic script, window.VO).
 *
 * - Which app opens which file (VO.APPS, VO.appFor).
 * - The host (VO.host): browser or desktop app, the files the system asks us to open.
 * - (Embedded copy: public/sw.js and its registration are gone — see the note at "service worker" below.)
 * - Files the start page hands to an app, and the Recent Files list. Both live in IndexedDB
 *   ('vibeoffice', stores 'handoff' and 'recent') in this browser only; nothing is uploaded.
 * - Unsaved versions: while a document has unsaved changes the app's own format is written (autosave)
 *   into its Recent Files entry as a draft; saving clears it. The Start Center, the apps' File menus and
 *   task panes and the Document Recovery pane all show the one list.
 *
 * An app page is opened as  <app>/?new           a blank document instead of the sample
 *                           <app>/?template=<id>  a new document from one of the app's templates
 *                           <app>/?open=<id>      a file handed over by VO.handoff (read once, then deleted)
 *                           <app>/?recent=<key>   a Recent Files entry (&draft: its unsaved version)
 * The app reads VO.launch while it starts, then calls VO.attach(app, { open, thumb }), and reports
 * VO.opened(file) after it opens a file and VO.saved(name, blob) after it saves in its own format.
 */
(function () {
  'use strict';
  const VO = (window.VO = window.VO || {});   // keep the language choice made before the first paint
  VO.VERSION = 'dev';   // the published copy says v<yyyy>.<mm>.<dd>-<6 hex> (tools/current-tree-version.sh, stamped by the sync script)
  if (window.__VO_HOST__ && window.__VO_HOST__.appVersion) VO.VERSION = window.__VO_HOST__.appVersion;   // the desktop app's build
  /** a version's date as a number (20261011), or null for a version without one (dev); copies before
   *  v<yyyy>.<mm>.<dd>-<hex> said <yyyy>.<mm>.<hex>, day 0 */
  const versionDay = (v) => {
    const m = /^v?(\d{4})\.(\d{2})(?:\.(\d{2})-[0-9a-f]{6})?/i.exec(v || '');
    return m ? +(m[1] + m[2] + (m[3] || '00')) : null;
  };
  /** compare two versions by their dates alone: < 0 if a is older than b, 0 for the same day, > 0 if newer;
   *  a version without a date (dev) is newer than any dated one */
  VO.compareVersions = (a, b) => {
    const x = versionDay(a), y = versionDay(b);
    return x === y ? 0 : x === null ? 1 : y === null ? -1 : Math.sign(x - y);
  };
  const BASE = new URL('..', document.currentScript.src).href;   // public/
  VO.base = BASE;

  /* ------------------------------------------------------------ apps */
  VO.APPS = {
    quire: { name: 'Quire', kind: 'Document', ext: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'htm', 'html', 'txt', 'md', 'markdown', 'zip'] },
    ledger: { name: 'Ledger', kind: 'Spreadsheet', ext: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'csv', 'tsv', 'prn', 'xml'] },
    lectern: { name: 'Lectern', kind: 'Presentation', ext: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt'] },
  };
  VO.ext = (name) => (/\.([^./\\]+)$/.exec(name || '') || ['', ''])[1].toLowerCase();
  /** the app for a file name, or null */
  VO.appFor = (name) => { const e = VO.ext(name); return Object.keys(VO.APPS).find((a) => VO.APPS[a].ext.includes(e)) || null; };
  VO.accept = () => Object.values(VO.APPS).flatMap((a) => a.ext.map((e) => '.' + e)).join(',');
  /** VO.url('quire', { new: true }) -> .../quire/index.html?new
   *  Embedded copy (tools-box): the folder form (.../quire/) relies on the static host
   *  answering with the directory's index.html; an explicit file needs nothing of the host. */
  VO.url = (app, params) => {
    const qs = Object.keys(params || {}).map((k) => (params[k] === true ? k : k + '=' + encodeURIComponent(params[k]))).join('&');
    return new URL(app + '/index.html', BASE).href + (qs ? '?' + qs : '');
  };

  /** the page's icon (browser tab) from an inline <svg> */
  VO.setIcon = (svg) => {
    if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    let link = document.querySelector('link[rel="icon"]');
    if (!link) link = document.head.appendChild(Object.assign(document.createElement('link'), { rel: 'icon' }));
    link.type = 'image/svg+xml';
    link.href = 'data:image/svg+xml,' + encodeURIComponent(svg);
  };

  /* ------------------------------------------------------------ the host */
  /** VO.host: the one interface between the pages and what runs them (docs/desktop.md, Host interface).
   *  kind 'web': a browser tab or the installed PWA; 'desktop': the native app (desktop/), which sets
   *  window.__VO_HOST__ and answers the host_* commands (desktop/src/host.rs). Pages use VO.host, never
   *  Tauri's or the system's file APIs directly. */
  const native = window.__VO_HOST__ && window.__TAURI_INTERNALS__ ? window.__VO_HOST__ : null;
  const call = (cmd, args) => window.__TAURI_INTERNALS__.invoke(cmd, args);
  VO.host = {
    kind: native ? 'desktop' : 'web',
    platform: native ? native.platform : 'web',
    version: native ? native.version : 0,
    /** whether this edition looks for its updates (updater/current.json): not the Mac App Store's, which the store updates */
    updates: !native || native.edition !== 'appstore',
    /** open a web or mail address in the system's browser: a new tab on the web, the default browser in the
     *  desktop app (which also sends every link that leaves the app there) */
    async openUrl(url) {
      if (native) return call('host_open_url', { url: String(url) });
      window.open(url, '_blank', 'noopener');
    },
    /** the whole UI's scale (1 = 100 %), or null where the page cannot set it (the web: the browser's own
     *  zoom, Ctrl + and Ctrl −, does it); the desktop app keeps it for the next start */
    async uiScale() { return native ? call('host_ui_scale') : null; },
    async setUiScale(scale) { if (native) return call('host_set_ui_scale', { scale: +scale }); },
    /** open one of the suite's pages in a window of its own: the desktop app makes a full window (host and
     *  all); the web opens a window or tab, false if the browser blocked it */
    async newWindow(url) {
      if (native) { await call('host_new_window', { url: new URL(url, location.href).href }); return true; }
      return !!window.open(url, '_blank');
    },
    /** close this window if it is an app window the suite opened; false for the Start Center's (and on the web) */
    async closeWindow() { return native ? call('host_close_window') : false; },
    /** Browse after a save: the Downloads folder, with this window's last saved file selected (desktop) */
    async showDownloads() { if (native) return call('host_show_downloads'); },
    /** fn(files) for the files the system asks us to open (Open with, a double-click on an associated
     *  file, the installed PWA's file handlers); the page that calls this takes them */
    onFiles(fn) {
      if (native) {
        // at start, and again when another start of the app hands this one its files (vo-files: Windows)
        const take = () => call('host_launch_files').then(async (list) => {
          if (!list.length) return;
          fn(await Promise.all(list.map(async (f) => new File([await call('host_read', { handle: f.handle })], f.name))));
        }).catch((e) => console.error('host:', e));
        take();
        window.addEventListener('vo-files', take);
      } else if ('launchQueue' in window) {
        window.launchQueue.setConsumer(async (params) => {
          if (params.files && params.files.length) fn(await Promise.all(params.files.map((h) => h.getFile())));
        });
      }
    },
  };

  /* ------------------------------------------------------------ updates */
  /** updater/current.json on the published site says, per policy and platform, the version an edition must
   *  be at least: older than "show-update-dialog" asks with a dialog, older than "show-update-button" shows an
   *  Update button. Versions compare by date (VO.compareVersions); a dev build never asks. The desktop app
   *  checks when the Start Center opens; nothing about the user or their files is sent. */
  VO.UPDATES = (native && native.updates) || 'https://vibeoffice.work/updater/current.json';
  /** updater/whatsnew.json: { "<version>": "<plain text>" }, one list for every platform */
  VO.WHATSNEW = (native && native.whatsnew) || 'https://vibeoffice.work/updater/whatsnew.json';
  /** the page the Update button opens (in the system's browser): it gets this edition's version and platform */
  VO.updateUrl = () => 'https://vibeoffice.work/updater/?from_version=' + encodeURIComponent(VO.VERSION)
    + '&current_platform=' + encodeURIComponent(VO.host.platform);
  /** { ask: 'dialog', 'button' or null (up to date), news: [{ version, text }] }: when an update is due, the
   *  whatsnew.json entries newer than this version, newest first (none if that file cannot be read).
   *  Throws when current.json cannot be read (offline). */
  VO.checkUpdate = async () => {
    const res = await fetch(VO.UPDATES, { cache: 'no-store' });
    if (!res.ok) throw new Error('updates: HTTP ' + res.status);
    const policy = (await res.json()) || {};
    const older = (k) => policy[k] && typeof policy[k][VO.host.platform] === 'string'
      && VO.compareVersions(VO.VERSION, policy[k][VO.host.platform]) < 0;
    const ask = older('show-update-dialog') ? 'dialog' : older('show-update-button') ? 'button' : null;
    let all = {};
    if (ask) try { const r = await fetch(VO.WHATSNEW, { cache: 'no-store' }); if (r.ok) all = (await r.json()) || {}; } catch (e) { /* the dialog without notes */ }
    const news = Object.entries(all)
      .filter(([v, text]) => typeof text === 'string' && versionDay(v) !== null && VO.compareVersions(VO.VERSION, v) < 0)
      .sort(([a], [b]) => VO.compareVersions(b, a))
      .map(([version, text]) => ({ version, text }));
    return { ask, news };
  };

  /* ------------------------------------------------------------ open source licenses */
  /* Embedded copy (tools-box): removed, together with public/office2003/oss-licenses.html — the page,
     its About links (common/ui.js, index.html) and every reference to it are gone. The license itself
     (Apache-2.0) is public/office2003/LICENSE, and the third-party notices are in the repository. */

  /* ------------------------------------------------------------ fonts on this computer */
  /** the font families Local Font Access listed (Chrome and Edge, after the user allows it), kept in this
   *  browser for every app of the suite; the Font boxes show them (L.localFonts in core.js) */
  const LOCAL_FONTS = 'vibeoffice.localFonts';
  let localFonts = null;
  window.addEventListener('storage', (e) => { if (e.key === LOCAL_FONTS || e.key === null) localFonts = null; });
  VO.localFonts = {
    supported: () => typeof window.queryLocalFonts === 'function',
    list() {
      if (!localFonts) { try { localFonts = JSON.parse(localStorage.getItem(LOCAL_FONTS)) || []; } catch (e) { localFonts = []; } }
      return localFonts;
    },
    /** ask the browser for the list (it asks the user the first time) and keep it; the number of families */
    async load() {
      const families = Array.from(new Set((await window.queryLocalFonts()).map((f) => f.family))).sort((x, y) => x.localeCompare(y));
      localFonts = families;
      try { localStorage.setItem(LOCAL_FONTS, JSON.stringify(families)); } catch (e) { /* storage off: kept for this page only */ }
      return families.length;
    },
    forget() { localFonts = []; try { localStorage.removeItem(LOCAL_FONTS); } catch (e) { /* storage off */ } },
  };

  /* Saving, for now: in the desktop app a save is a download into the Downloads folder; the native side says
     which file it wrote (the name only), and the app shows it with Browse (docs/desktop.md, Saving) */
  if (native) window.addEventListener('vo-download', (e) => {
    const ui = window.L && window.L.ui, d = e.detail || {};
    if (!ui) return;
    if (d.ok) ui.toast(VO.t('已保存到“下载”文件夹：{0}', d.name), 8000, { label: '浏览', run: () => VO.host.showDownloads().catch((err) => console.error('browse:', err)) });
    else ui.toast('无法将文件保存到“下载”文件夹。', 8000);
  });

  /* ------------------------------------------------------------ service worker */
  /* The published site, https://vibeoffice.work/, is cached for offline use, one version per page (sw.js); the
     worker is registered there and on 127.0.0.1 (any port: a published copy tried locally; the dev server's own
     files are unstamped, so the worker caches nothing there), and anywhere else (preview deployments) any earlier
     registration is removed. While the worker caches a version it reports to every page: 'vo-cache' on window, detail { state:
     'progress', 'done' or 'failed', update (an earlier version was cached: this is an update), progress (0 to 1),
     version }. updateViaCache 'none': the browser always asks the server for sw.js, so a new version is found on
     the next visit. */
  /* Embedded copy (tools-box): the registration is gone, together with public/office2003/sw.js. Upstream
     registers it only on vibeoffice.work and on 127.0.0.1, so it would never run on this site anyway; the
     'vo-cache' progress the Settings dialog shows therefore stays empty here (see index.html). */


  /* ------------------------------------------------------------ IndexedDB */
  let dbp = null;
  function db() {
    return dbp || (dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open('vibeoffice', 1);
      r.onupgradeneeded = () => {
        r.result.createObjectStore('recent', { keyPath: 'key' });
        r.result.createObjectStore('handoff', { keyPath: 'id' });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  }
  /** run fn(store) in one transaction; resolves with the result of the request fn returns, once committed */
  async function tx(name, mode, fn) {
    const t = (await db()).transaction(name, mode);
    return new Promise((resolve, reject) => {
      let out;
      const req = fn(t.objectStore(name));
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => resolve(out);
      t.onerror = t.onabort = () => reject(t.error);
    });
  }
  const chan = 'BroadcastChannel' in window ? new BroadcastChannel('vibeoffice') : null;
  const listeners = new Set();
  const notify = () => { VO.recent.refresh(); for (const fn of listeners) if (!fn.cacheOnly) fn(); };   // cache users redraw after the refresh
  const changed = () => { if (chan) chan.postMessage('recent'); notify(); };
  if (chan) chan.addEventListener('message', (e) => { if (e.data === 'recent') notify(); else if (e.data === 'look') applyLook(storedLook()); });
  /** call fn when the Recent Files list changes (here or in another page); returns the function that stops it */
  VO.onRecentChange = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };

  /* ------------------------------------------------------------ Looks */
  /* the chrome's colours (common/looks.css) and, for a look with layout 'ribbon', the ribbon instead of menus and
     toolbars (common/ribbon.js); one choice for the whole suite: <html data-look data-layout> is set before
     the first paint by each page's head script from localStorage 'vibeoffice.look'; changing it here
     repaints every open window of the suite. Canvas-drawn chrome (rulers, sheet headers) redraws on 'vo-look'. */
  const N = (s) => s;   // UI text, translated where it is shown (VO.N, common/i18n.js)
  N('空白文档'); N('空白工作簿'); N('空白演示文稿');   // templates.json's Blank entries (tools/templates.mjs)
  const LOOKS = [{ id: 'classic', name: N('经典'), note: 'Office 2003 的蓝色工具栏与菜单。' },
    { id: 'paper', name: N('纸张'), note: '明亮宁静：温暖的纸张色调、青绿色点缀、圆润的形状。' },
    { id: 'paper2016', name: N('纸张 2016'), layout: 'ribbon', note: '纸张风格，以按任务分组的命令选项卡代替菜单和工具栏。' }];
  const storedLook = () => { try { return localStorage.getItem('vibeoffice.look') || 'classic'; } catch (e) { return 'classic'; } };
  const applyLook = (id) => {
    const root = document.documentElement;
    if (!LOOKS.some((l) => l.id === id)) return;
    const changed = root.dataset.look !== id;
    root.dataset.look = id;
    root.dataset.layout = LOOKS.find((l) => l.id === id).layout || 'bars';
    const meta = document.querySelector('meta[name="theme-color"]');
    const c = getComputedStyle(root).getPropertyValue('--theme-color').trim();
    if (meta && c) meta.content = c;
    if (changed) window.dispatchEvent(new Event('vo-look'));
  };
  VO.look = {
    /** [{ id, name, note, layout }] */
    list: () => LOOKS.map((l) => ({ ...l })),
    get: () => document.documentElement.dataset.look || 'classic',
    /** 'bars' (menus and toolbars) or 'ribbon' */
    layout: () => document.documentElement.dataset.layout || 'bars',
    set(id) {
      if (!LOOKS.some((l) => l.id === id)) return;
      try { localStorage.setItem('vibeoffice.look', id); } catch (e) { /* this window still changes */ }
      applyLook(id);
      if (chan) chan.postMessage('look');
    },
  };
  applyLook(storedLook());   // a page without the head script, or a stored look that changed meanwhile

  /* ------------------------------------------------------------ Recent Files */
  const MAX_RECENT = 30;
  const MAX_BYTES = 50 * 1024 * 1024;   // bigger files are not kept (open them from disk again)
  const prune = async () => {
    /* entries with unsaved changes are never dropped */
    const all = (await VO.recent.list()).filter((e) => !e.draft);
    if (all.length > MAX_RECENT) await tx('recent', 'readwrite', (s) => { for (const e of all.slice(MAX_RECENT)) s.delete(e.key); });
  };
  /* Safari clears a site's data after about 7 days without a visit, unsaved versions included; once a
     document has unsaved changes, ask the browser to keep the data (Safari and Chrome decide without
     asking the user; Firefox may ask) */
  let persistAsked = false;
  const keepData = () => {
    if (persistAsked || !navigator.storage || !navigator.storage.persist) return;
    persistAsked = true;
    navigator.storage.persisted().then((yes) => yes || navigator.storage.persist()).catch(() => {});
  };
  VO.recent = {
    /** newest first: { key, app, name, size, time, file: Blob|null (opened or saved), thumb: Blob|null,
     *  draft: { file: Blob, name, time } | null (unsaved changes) }; a never-saved document has only a draft */
    async list() {
      const all = (await tx('recent', 'readonly', (s) => s.getAll())) || [];
      return all.sort((a, b) => b.time - a.time);
    },
    get: (key) => tx('recent', 'readonly', (s) => s.get(key)),
    /** the list as last read (for menus, which are built synchronously) */
    cache: [],
    refresh: (() => { let t = 0; return () => { clearTimeout(t); t = setTimeout(() => VO.recent.list().then((l) => { VO.recent.cache = l; for (const fn of listeners) if (fn.cacheOnly) fn(); }, () => {}), 30); }; })(),
    /** this app's entries, newest first, from the cache */
    forApp: (app) => VO.recent.cache.filter((e) => e.app === (app || (current && current.app))),
    /** add or refresh the entry for this file name (a newer copy replaces the older one, its thumbnail stays);
     *  keepDraft: opening a file keeps unsaved changes made to it, saving it clears them */
    async put(app, name, blob, keepDraft) {
      if (!app || !blob || blob.size > MAX_BYTES) return null;
      const key = app + ':' + name;
      const old = await VO.recent.get(key);
      await tx('recent', 'readwrite', (s) => s.put({ key, app, name, size: blob.size, time: Date.now(), file: blob, thumb: old ? old.thumb : null, draft: keepDraft && old ? old.draft || null : null }));
      await prune();
      changed();
      return key;
    },
    /** store the unsaved version of a document */
    async setDraft(key, app, name, draftName, blob, lossState) {
      const e = (await VO.recent.get(key)) || { key, app, name, size: 0, file: null, thumb: null };
      if (!e.file) e.name = name;
      e.draft = { file: blob, name: draftName || name, time: Date.now(), lossState };
      e.time = e.draft.time;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
      keepData();
    },
    /** throw the unsaved version away; a document that was never saved leaves the list */
    async discardDraft(key) {
      const e = await VO.recent.get(key);
      if (!e) return;
      if (!e.file) await tx('recent', 'readwrite', (s) => s.delete(key));
      else { e.draft = null; await tx('recent', 'readwrite', (s) => s.put(e)); }
      changed();
    },
    async setThumb(key, thumb) {
      const e = await VO.recent.get(key);
      if (!e) return;
      e.thumb = thumb;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
    },
    async remove(key) { await tx('recent', 'readwrite', (s) => s.delete(key)); changed(); },
    async clear() { await tx('recent', 'readwrite', (s) => s.clear()); changed(); },
  };
  /* Embedded copy (tools-box): every Recent Files write, chained in order, so VO.exit can wait
     for them. VO.saved() is "fire and forget" and an app navigates to the Start Center right
     after, but a navigation aborts an open IndexedDB transaction — the file that was just saved
     could disappear from Recent Files. */
  let recentWrites = Promise.resolve();
  for (const m of ['put', 'setDraft', 'discardDraft', 'setThumb', 'remove', 'clear']) {
    const fn = VO.recent[m];
    VO.recent[m] = (...a) => { const p = fn.apply(VO.recent, a); recentWrites = recentWrites.then(() => p).catch(() => {}); return p; };
  }

  /* ------------------------------------------------------------ handing files to apps */
  /** where an app page opens: a new window in the desktop app or when installed (the start page stays, as in
   *  LibreOffice), else this tab */
  VO.go = async (url) => {
    const standalone = matchMedia('(display-mode: standalone)').matches || matchMedia('(display-mode: window-controls-overlay)').matches;
    try { if ((VO.host.kind === 'desktop' || standalone) && await VO.host.newWindow(url)) return; } catch (e) { console.error('new window:', e); }
    location.href = url;
  };
  /** File ▸ Exit, after the app has closed its documents: an app window the Start Center opened
   * (installed) closes and shows the Start Center behind it; a tab goes to the Start Center */
  VO.exit = async () => {
    if (await VO.host.closeWindow().catch(() => false)) return;   // the desktop app's app windows
    if (window.opener && !window.opener.closed) { window.close(); if (window.closed) return; }
    /* Embedded copy (tools-box): let in-flight Recent Files writes commit before the page unloads
       (a navigation aborts an open IndexedDB transaction). Capped so a stuck write cannot block Exit. */
    await Promise.race([recentWrites, new Promise((r) => setTimeout(r, 2000))]);
    location.href = new URL('index.html', BASE).href;   // embedded copy: explicit file, see VO.url
  };
  /** open a file in the app made for it; false if no app opens this kind of file */
  VO.handoff = async (file) => {
    const app = VO.appFor(file.name);
    if (!app) return false;
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const day = Date.now() - 86400000;
    await tx('handoff', 'readwrite', (s) => {
      s.put({ id, name: file.name, file, time: Date.now() });
      s.openCursor().onsuccess = (e) => { const c = e.target.result; if (!c) return; if (c.value.time < day) c.delete(); c.continue(); };
    });
    VO.go(VO.url(app, { open: id }));
    return true;
  };

  /* ------------------------------------------------------------ the app side */
  const q = new URLSearchParams(location.search);
  /** how this app page was opened: null (plain: show the sample), { new }, { template: id }, { open: id } or { recent: key } */
  VO.launch = q.has('open') ? { open: q.get('open') } : q.has('recent') ? { recent: q.get('recent'), draft: q.has('draft') }
    : q.has('template') ? { template: q.get('template') } : q.has('new') ? { new: true } : null;

  let current = null;   // { app, hooks }
  /** an app is ready: open the file it was launched with, if any, then show the Document Recovery pane when
   *  other documents of this app have unsaved changes.
   *  hooks: open(file, { draft, name, saved }), thumb() -> a canvas (or image) of the first page, message(text),
   *  templates() -> [{ id, name }] (read by tools/templates.mjs); for autosave: snapshot(doc) -> Blob in the
   *  app's format, docInfo(doc) -> { name, draftName }, isDirty(doc), current() -> doc, autosave() -> false
   *  when turned off; showRecovery() opens the Document Recovery pane */
  VO.attach = async (app, hooks) => {
    current = { app, hooks: hooks || {} };
    VO.recent.refresh();
    const l = VO.launch;
    if (l && !l.new) {
      if (history.replaceState) history.replaceState(null, '', location.pathname);   // a reload starts clean
      if (l.open) {
        let rec = null;
        try {
          rec = await tx('handoff', 'readonly', (s) => s.get(l.open));
          if (rec) await tx('handoff', 'readwrite', (s) => s.delete(l.open));
        } catch (e) { console.error(e); }
        if (rec && rec.file) await current.hooks.open(asFile(rec.file, rec.name), {});
        else (current.hooks.message || alert)(VO.t('此处已无法访问该文件。请从计算机中重新打开。'));
      } else if (l.recent) await VO.openRecent(l.recent, { draft: l.draft });
    }
    setTimeout(async () => {
      const others = (await VO.recent.list()).filter((e) => e.app === app && e.draft && !openKeys.has(e.key));
      if (others.length && current.hooks.showRecovery) current.hooks.showRecovery();
    }, 700);
  };
  const asFile = (blob, name) => (blob instanceof File && blob.name === name ? blob : new File([blob], name, { type: blob.type }));
  /** open a Recent Files entry in this app: its unsaved version when draft is true, else the saved file */
  VO.openRecent = async (key, o) => {
    let e = null;
    try { e = await VO.recent.get(key); } catch (err) { console.error(err); }
    const draft = !!(o && o.draft && e && e.draft);
    const blob = e && (draft ? e.draft.file : e.file);
    if (!blob) { (current.hooks.message || alert)(VO.t('此处已无法访问该文件。请从计算机中重新打开。')); return; }
    const f = asFile(blob, draft ? e.draft.name : e.name);
    if (draft) f.voDraft = key;
    await current.hooks.open(f, { draft, name: e.name, saved: !!e.file, lossState: draft ? e.draft.lossState : undefined });
  };
  /** File ▸ New… inside an app: Blank and the app's templates with their previews (the Start Center's
   *  Create File gallery, common/templates.json), in the app's own dialog. create(id): id '' is Blank. */
  VO.newDialog = async (app, create) => {
    const ui = window.L && window.L.ui;
    if (!ui || !ui.dialog) return create('');
    let list = null;
    try { list = (await (await fetch(BASE + 'common/templates.json')).json())[app]; } catch (e) { console.warn('templates:', e); }
    if (!list || !list.length) list = [{ id: '', name: 'Blank', img: null }];
    if (!document.getElementById('vo-gallery-css')) {
      const st = document.createElement('style');
      st.id = 'vo-gallery-css';
      st.textContent = '.vo-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px;max-height:min(62vh,540px);overflow:auto;padding:6px;background:var(--surface);border:1px solid var(--field-border)}'
        + '.vo-tpl{display:flex;flex-direction:column;gap:4px;padding:5px;border:1px solid transparent;border-radius:2px;background:none;text-align:left;cursor:default;font:inherit;color:var(--ink)}'
        + '.vo-tpl:hover,.vo-tpl:focus-visible{background:var(--menu-hot);border-color:var(--card-hot-border);outline:none}'
        + '.vo-tpl span{display:flex;align-items:center;justify-content:center;height:112px;background:var(--card-thumb-bg);border:1px solid var(--card-thumb-line);overflow:hidden}'
        + '.vo-tpl img{max-width:100%;max-height:100%;box-shadow:0 1px 2px var(--tip-shadow-c)}'
        + '.vo-tpl b{font-weight:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
        + '.vo-gal-note{margin:0 0 6px}';
      document.head.appendChild(st);
    }
    const gal = document.createElement('div');
    gal.className = 'vo-gal';
    let d = null;
    for (const t of list) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'vo-tpl'; b.title = VO.t(t.name);
      const pic = document.createElement('span');
      if (t.img) { const img = document.createElement('img'); img.src = BASE + t.img; img.alt = ''; pic.append(img); }
      const name = document.createElement('b');
      name.textContent = VO.t(t.name);
      b.append(pic, name);
      b.addEventListener('click', () => { if (d) d.close(-1); create(t.id); });
      gal.append(b);
    }
    const body = document.createElement('div');
    const note = document.createElement('p');
    note.className = 'vo-gal-note';
    note.textContent = VO.t('选择空白文件或模板：');
    body.append(note, gal);
    d = ui.dialog({ title: VO.t({ quire: '新建文档', ledger: '新建电子表格', lectern: '新建演示文稿' }[app]), width: 700, body, buttons: [{ label: '取消', cancel: true }], focus: gal.firstChild });
  };
  /** the open app's templates: [{ id, name }] */
  VO.templates = () => (current && current.hooks.templates ? current.hooks.templates() : []);
  /** a JPEG (at most 320 px) of the open document's first page, or null */
  VO.thumbnail = async () => {
    const make = current && current.hooks.thumb;
    const src = make && (await make());
    if (!src) return null;
    const k = Math.min(1, 320 / src.width, 320 / src.height);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(src.width * k)); c.height = Math.max(1, Math.round(src.height * k));
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(src, 0, 0, c.width, c.height);
    return new Promise((r) => c.toBlob(r, 'image/jpeg', 0.88));
  };
  /** the thumbnail for a Recent Files entry, a little after the document has been drawn */
  function thumbLater(key) {
    if (!key) return;
    setTimeout(async () => {
      try { const blob = await VO.thumbnail(); if (blob) await VO.recent.setThumb(key, blob); }
      catch (e) { console.warn('thumbnail:', e); }
    }, 600);
  }
  /* ------------------------------------------------------------ documents, autosave */
  const docKeys = new WeakMap();   // the app's document object -> its Recent Files key
  const openKeys = new Set();      // keys of the documents open in this page (left out of Document Recovery)
  const timers = new WeakMap(), slow = new WeakMap();
  const DRAFT_DELAY = 15000;
  const bind = (doc, key) => { if (doc) docKeys.set(doc, key); openKeys.add(key); };
  /** the app opened a file (from any source) as doc */
  VO.opened = (file, doc) => {
    if (!current) return;
    if (file.voDraft) { bind(doc, file.voDraft); return; }   // an unsaved version: keep its entry as it is
    VO.recent.put(current.app, file.name, file, true).then((key) => { if (key) bind(doc, key); thumbLater(key); }, (e) => console.warn('recent files:', e));
  };
  /** the app saved doc in its own format as name: that is its entry now, without unsaved changes */
  VO.saved = (name, blob, doc) => {
    if (!current) return;
    const old = doc && docKeys.get(doc);
    if (doc) clearTimeout(timers.get(doc));
    VO.recent.put(current.app, name, blob, false).then(async (key) => {
      if (!key) return;
      if (old && old !== key) { await VO.recent.discardDraft(old); openKeys.delete(old); }
      bind(doc, key);
      thumbLater(key);
    }, (e) => console.warn('recent files:', e));
  };
  /** doc has new unsaved changes: its unsaved version is written a little later */
  VO.changed = (doc) => {
    if (!current || !doc || !current.hooks.snapshot) return;
    clearTimeout(timers.get(doc));
    timers.set(doc, setTimeout(() => saveDraft(doc), Math.max(DRAFT_DELAY, slow.get(doc) || 0)));
  };
  async function saveDraft(doc) {
    const H = current.hooks;
    if ((H.isDirty && !H.isDirty(doc)) || (H.autosave && H.autosave() === false)) return;
    const t0 = performance.now();
    let blob = null;
    try { blob = await H.snapshot(doc); } catch (e) { console.warn('autosave:', e); return; }
    slow.set(doc, (performance.now() - t0) * 30);   // a document that takes 1 s to write is saved every 30 s at most
    if (!blob || blob.size > MAX_BYTES || (H.isDirty && !H.isDirty(doc))) return;
    const info = H.docInfo(doc);
    let key = docKeys.get(doc);
    if (!key) { key = current.app + ':~' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); bind(doc, key); }
    try { await VO.recent.setDraft(key, current.app, info.name, info.draftName, blob, info.lossState); } catch (e) { console.warn('autosave:', e); return; }
    if (!H.current || H.current() === doc) thumbLater(key);
  }
  /** the user chose not to keep doc's changes (No in "Do you want to save the changes…?") */
  VO.discard = (doc) => {
    if (!doc) return;
    clearTimeout(timers.get(doc));
    const key = docKeys.get(doc);
    if (!key) return;
    docKeys.delete(doc);
    openKeys.delete(key);
    VO.recent.discardDraft(key).catch(() => {});
  };

  /* ------------------------------------------------------------ the list inside the apps */
  const when = (t) => {
    const d = new Date(t), today = new Date().toDateString() === d.toDateString();
    const loc = 'zh-CN';
    return today ? d.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' }) : d.toLocaleDateString(loc, { month: 'short', day: 'numeric' });
  };
  /* its text (common/text.js) */
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = VO.t ? VO.t(text) : text; return e; };
  const icon = (name) => (window.L && L.icons ? L.icons.get(name) : '');
  /** File menu: the app's four most recent entries ({ label, run } items for L.ui menus) */
  VO.recentMenuItems = (max) => VO.recent.forApp().slice(0, max || 4).map((e, i) => ({
    label: '&' + (i + 1) + ' ' + (e.draft ? VO.t('{0}（有未保存的更改）', e.name) : e.name).replace(/&/g, '&&'), raw: true,
    run: () => VO.openRecent(e.key, { draft: !!e.draft }),
  }));
  /** task pane: the app's recent entries as links (Getting Started ▸ Open); keeps itself up to date */
  VO.recentLinks = (max) => {
    const box = el('div', 'vo-recent');
    const draw = () => {
      if (!box.isConnected && box.dataset.drawn) { stop(); return; }
      box.dataset.drawn = '1';
      box.replaceChildren(...VO.recent.forApp().slice(0, max || 4).map((e) => {
        const b = el('button', 'tp-link');
        b.type = 'button';
        b.title = e.draft ? VO.t('{0} — {1} 的未保存更改', e.name, new Date(e.draft.time).toLocaleString('zh-CN')) : e.name;
        b.translate = false;
        const ic = el('span'); ic.innerHTML = icon(e.draft ? 'save' : 'open');
        const name = el('span');
        name.textContent = e.draft ? VO.t('{0}（未保存）', e.name) : e.name;
        b.append(ic, name);
        b.addEventListener('click', () => VO.openRecent(e.key, { draft: !!e.draft }));
        return b;
      }));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    draw();
    return box;
  };
  /** the Document Recovery task pane: this app's unsaved versions, each to open or to discard */
  VO.recoveryPane = (b) => {
    const list = el('div', 'vo-recovery');
    const draw = () => {
      if (!list.isConnected && list.dataset.drawn) { stop(); return; }
      list.dataset.drawn = '1';
      const drafts = VO.recent.forApp().filter((e) => e.draft);
      list.replaceChildren(...(drafts.length ? drafts.map((e) => {
        const item = el('div', 'tp-sec');
        const head = el('div', 'tp-h', e.name);
        const from = e.file ? VO.t('{0} 的未保存更改', when(e.draft.time)) : VO.t('{0} 的未保存更改（从未保存）', when(e.draft.time));
        const note = el('div', 'tp-note', openKeys.has(e.key) ? VO.t('{0} — 在此窗口打开', from) : from);
        const open = el('button', 'tp-link'); open.type = 'button';
        const oi = el('span'); oi.innerHTML = icon('open'); open.append(oi, el('span', null, '打开未保存的版本'));
        open.addEventListener('click', () => VO.openRecent(e.key, { draft: true }));
        const drop = el('button', 'tp-link'); drop.type = 'button';
        const di = el('span'); di.innerHTML = icon('delete') || icon('cut'); drop.append(di, el('span', null, '放弃更改'));
        drop.addEventListener('click', () => VO.recent.discardDraft(e.key));
        item.append(head, note, open, drop);
        return item;
      }) : [el('div', 'tp-note', '没有待恢复的未保存版本。')]));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    b.append(el('div', 'tp-note', VO.t('{0} 保留了这些文档中未保存的更改。请打开并保存需要保留的文档，放弃其余文档的更改。', current ? VO.APPS[current.app].name : 'Office 2003')), list);
    VO.recent.list().then((l) => { VO.recent.cache = l; draw(); }, () => draw());
  };
})();
