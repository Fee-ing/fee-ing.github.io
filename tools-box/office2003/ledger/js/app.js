/* Ledger — application controller: window chrome, workbooks, files, menus, toolbars, status bar, sheet tabs. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, O = L.ops, F = L.formula, NF = L.numfmt, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const G = () => L.grid;

  /* original application icon: a green ledger sheet with a bound spine */
  L.icons.app = function (size) {
    const sz = size || 16;
    return `<svg class="appico" width="${sz}" height="${sz}" viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 1.5h8l3 3v10h-11z" fill="#fff" stroke="#1e6b2e"/><path d="M11.5 1.5v3h3" fill="#d8eedc" stroke="#1e6b2e"/><rect x="1.5" y="4.5" width="7" height="8" fill="#2f8f45" stroke="#1a5a28"/><path d="M3.3 6.3v4.4h3.2" fill="none" stroke="#fff" stroke-width="1.4"/><path d="M10 7.5h3M10 9.5h3M10 11.5h3M9.5 6.5v6" stroke="#7aa886"/></svg>`;
  };
  /* the grid around commands, toolbar boxes and dialogs (common/ui.js) */
  ui.hooks = {
    beforeExec() {},
    refocus() { if (L.grid) L.grid.focus(); },
    commitEdit() { if (L.editor && L.editor.active) L.editor.commit(); },
  };
  const MAXR = M.MAXR, MAXC = M.MAXC;

  /* a new workbook as the Office of the UI's language makes it (common/locales.js): its font and paper */
  const LOCALE = (L.locale && L.locale.get().sheet) || {};
  const DEFAULT_OPTS = {
    statusBar: true, formulaBar: true, taskOpen: true, paneFollows: true, moveAfterEnter: true, enterDir: 'down', autoComplete: true, dragDrop: true,
    toolbars: { standard: true, formatting: true, drawing: false, chart: 'auto', auditing: false, reviewing: false, picture: 'auto', borders: false, list: 'auto' },
    userName: 'Ledger 用户', recent: [], defaultFont: LOCALE.font || 'Arial', defaultSize: LOCALE.sz || 10, sheetsInNew: 3, r1c1: false, calcMode: 'auto', iterate: false, iterateCount: 100, iterateDelta: 0.001, autoCalcFn: 'SUM',
  };
  const A = (L.app = Object.assign(L.app || {}, { opts: L.clone(DEFAULT_OPTS), books: [], cur: -1, lastFill: '#FFFF00', lastFontColor: '#FF0000', lastBorder: 'bottom' }));
  const book = () => A.books[A.cur];
  const wb = () => G().wb;
  const sh = () => G().sheet();

  /* ------------------------------------------------------------ options */
  /* the options that follow the UI's language, with their English values (common/locales.js) */
  const EN_OPTS = { defaultFont: 'Arial', defaultSize: 10 };
  A.loadOpts = () => { A.opts = L.deepMerge(L.clone(DEFAULT_OPTS), L.locale.savedOpts(L.store.get('opts', {}), EN_OPTS)); };
  A.saveOpts = () => { L.store.set('opts', L.locale.optsToSave(A.opts, DEFAULT_OPTS, EN_OPTS)); };
  A.applyOpts = () => {
    L.$('#statusbar').hidden = A.opts.statusBar === false;
    L.$('#fbar').hidden = A.opts.formulaBar === false;
    A.updateToolbars();
    A.saveOpts();
    G().paint();
  };

  /* ------------------------------------------------------------ messages */
  A.error = function (e) {
    if (!e) return;
    console.error(e);
    const msg = e.message || String(e);
    ui.msg(msg, { icon: e.code && e.code !== 'internal' ? 'warn' : 'error' });
  };
  let statusText = '';
  A.status = (t) => { statusText = t || ''; updateStatus(); };

  /* ------------------------------------------------------------ workbooks */
  A.blankWorkbook = function () {
    const font = { name: A.opts.defaultFont || 'Arial', sz: A.opts.defaultSize || 10 };
    if (LOCALE.charset && font.name === LOCALE.font) Object.assign(font, { family: LOCALE.family, charset: LOCALE.charset });
    const w = new M.Workbook({ font });
    if (LOCALE.paper) w.paper = LOCALE.paper;   // its sheets' paper (Sheet's print.paper)
    for (let i = 0; i < (A.opts.sheetsInNew || 3); i++) w.addSheet('Sheet' + (i + 1));
    w.props.creator = A.opts.userName || '';
    w.props.created = new Date().toISOString();
    return w;
  };
  let untitled = 0;
  A.newWorkbook = function () {
    untitled++;
    A.addBook(A.blankWorkbook(), 'Book' + untitled, { type: 'xlsx', untitled: true });
  };
  /** register and show a workbook */
  A.addBook = function (w, name, o) {
    o = o || {};
    w.fileName = name;
    const b = { wb: w, name, type: o.type || 'xlsx', untitled: !!o.untitled, saved: !o.untitled, dirty: false };
    w.undo.onChange = () => { b.dirty = true; updateTitle(); ui.refresh(); if (window.VO) VO.changed(b); };
    /* replace a pristine untitled workbook (like Excel does with Book1) */
    const cur = book();
    if (cur && cur.untitled && !cur.dirty && !o.untitled && A.books.length === 1) { A.books[A.cur] = b; }
    else { A.books.push(b); A.cur = A.books.length - 1; }
    A.cur = A.books.indexOf(b);
    showBook();
    return b;
  };
  function showBook() {
    const b = book();
    if (!b) return;
    if (L.editor.active) L.editor.cancel();
    G().setWorkbook(b.wb);
    renderTabs();
    updateTitle();
    L.editor.show();
    G().focus();
    ui.refresh();
    L.bus.emit('book');
    updateStatus();
  }
  A.switchBook = (i) => { if (i === A.cur || !A.books[i]) return; A.cur = i; showBook(); };
  /** close the current workbook (asking to save); false when the user cancels */
  A.closeWorkbook = async function () {
    const b = book();
    if (!b) return true;
    if (b.dirty) {
      const r = await ui.msg(L.t("要保存对“{0}”所做的更改吗？", b.name), { icon: 'warn', buttons: ['是(&Y)', '否(&N)', '取消'] });
      if (r === 2 || r == null) return false;
      if (r === 0) { const ok = await A.save(); if (!ok) return false; }
      if (r === 1 && window.VO) VO.discard(b);   // its unsaved version goes too
    }
    A.books.splice(A.cur, 1);
    if (!A.books.length) { A.newWorkbook(); return true; }
    A.cur = Math.min(A.cur, A.books.length - 1);
    showBook();
    return true;
  };
  function updateTitle() {
    const b = book();
    document.title = (b ? b.name + ' - ' : '') + 'Ledger';
  }

  /* ------------------------------------------------------------ open */
  A.openDialog = async function (accept) {
    const files = await L.pickFiles(accept || '.xlsx,.xlsm,.xltx,.xltm,.csv,.txt,.tsv,.prn,.xml', true);
    if (files.length) A.openFiles(files);
  };
  A.openFiles = async function (files) {
    for (const f of files) {
      try { await A.openFile(f); } catch (e) { A.error(e); }
    }
  };
  A.openFile = async function (file, password, opts) {
    opts = opts || {};
    const name = file.name || 'Book.xlsx';
    const ext = (name.split('.').pop() || '').toLowerCase();
    ui.busy(true, L.t('正在打开 {0}...', name));
    try {
      const bytes = new Uint8Array(await L.readAsArrayBuffer(file));
      let w, type = 'xlsx';
      if ((opts.wizard || /^(csv|txt|tsv|prn)$/.test(ext)) && !(bytes[0] === 0x50 && bytes[1] === 0x4b) && !(bytes[0] === 0xd0 && bytes[1] === 0xcf)) {
        ui.busy(false);
        const res = await L.csv.importDialog(bytes, name, { wizard: opts.wizard });
        if (!res) return;
        w = res.wb; type = ext === 'csv' ? 'csv' : 'txt';
      } else if (ext === 'xml' || ext === 'htm' || ext === 'html' || bytes[0] === 0x3c || (bytes[0] === 0xef && bytes[3] === 0x3c)) {
        const text = L.csv.decode(bytes).text;
        if (!/urn:schemas-microsoft-com:office:spreadsheet/.test(text) && /<table[\s>]/i.test(text)) { ui.busy(false); A.openHTML(text, name); return; }
        w = await L.xmlss.read(bytes);
        type = 'xmlss';
      } else {
        try { w = await L.xlsxRead.read(bytes, { password }); }
        catch (e) {
          if (e.code === 'badpassword') {
            /* as Excel: say so and cancel the open; File ▸ Open again to retry */
            ui.busy(false);
            await ui.msg(e.message, { icon: 'warn' });
            return;
          }
          if (e.code === 'password') {
            ui.busy(false);
            const pw = ui.password ? await ui.password(name) : await ui.prompt(L.t("“{0}”已受保护。\n\n密码：", name), '', '密码');
            if (pw == null) return;
            return A.openFile(file, pw, opts);
          }
          if (e.code === 'xls') { ui.busy(false); if (L.xls && L.xls.read) { w = await L.xls.read(bytes); type = 'xls'; } else throw Object.assign(new Error(L.t("“{0}”是 Excel 97–2003 二进制工作簿（.xls）。Ledger 可打开 .xlsx、.xlsm、CSV 和 XML Spreadsheet 2003 文件。请先在 Excel 或 LibreOffice 中打开此文件并另存为 .xlsx。", name)), { code: 'xls' }); }
          else throw e;
        }
        type = w.ooxmlFormat || (w.macroEnabled ? 'xlsm' : 'xlsx');
        if (w.repaired) ui.toast('此文件已损坏；Ledger 已尽可能修复。');
      }
      A.prepare(w);
      const b = A.addBook(w, name, { type });
      if (password) b.password = password;
      if (window.VO && file.name) VO.opened(file, b);
      updateTitle();
    } finally { ui.busy(false); }
  };
  /** after loading: dependency graph, values that were never calculated, row heights */
  A.prepare = function (w) {
    for (const s of w.sheets) if (s.kind === 'chartsheet') { s.view.grid = false; s.view.headings = false; for (const d of s.drawings) if (d.anchor && d.anchor.type === 'abs' && (!d.anchor.w || !d.anchor.h)) Object.assign(d.anchor, { x: 0, y: 0, w: 640, h: 440 }); }
    if (L.tstyle) L.tstyle.reset();
    C.rebuild(w);
    const nodes = Array.from(w.graph.byCell.values());
    const full = w.calcPr.fullCalcOnLoad;
    for (const n of nodes) if (full || n.cell.dirty || n.cell.ca) n.cell.dirty = true;
    try { C.evaluate(nodes.filter((n) => n.cell.dirty)); } catch (e) { console.error(e); }
    /* volatile functions update on open (Excel does the same) */
    try { const vol = Array.from(w.graph.volatile); if (vol.length && w.calcPr.mode !== 'manual') { for (const n of vol) n.cell.dirty = true; C.changed(w, []); } } catch (e) { console.error(e); }
    /* rows without a stored height grow for wrapped or large text */
    for (const s of w.sheets) {
      if (s.maxR > 20000) continue;
      s.rows.forEach((row, r) => {
        if (!row || row.ht != null || row.customHeight) return;
        let need = false;
        for (const k in row.cells) { const cl = row.cells[k]; if (!cl || (cl.v == null && !cl.s)) continue; const st = LY.styleOf(s, r, +k, cl); if ((st.align && (st.align.wrap || st.align.rot)) || (st.font && st.font.sz && st.font.sz > (w.defaultFont.sz || 11)) || (typeof cl.v === 'string' && cl.v.indexOf('\n') >= 0)) { need = true; break; } }
        if (need) { const pt = O.bestHeight(s, r); if (Math.abs(pt - M.defaultRowPt(s)) > 0.01) row.auto = pt; }
      });
      LY.invalidate(s);
    }
    w.undo.clear();
  };
  /** a web page: each <table> becomes a sheet, named after the heading just before it (as "Save as Web Page" writes them) */
  A.openHTML = function (html, name) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const tables = Array.from(doc.querySelectorAll('table')).filter((t) => !t.parentElement.closest('table'));
    if (!tables.length) { ui.msg(L.t("“{0}”不包含可打开的表格。", name), { icon: 'warn' }); return; }
    const styles = Array.from(doc.querySelectorAll('style')).map((x) => x.outerHTML).join('');
    const w = A.blankWorkbook();
    w.sheets.splice(1);
    const used = new Set();
    const sheetName = (t, i) => {
      let el = t.previousElementSibling;
      let s = el && /^H[1-6]$/.test(el.tagName) ? el.textContent.trim() : tables.length === 1 ? name.replace(/\.[^.]+$/, '') : 'Sheet' + (i + 1);
      s = s.replace(/[\\/?*[\]:]/g, '_').replace(/^'|'$/g, '').slice(0, 31) || 'Sheet' + (i + 1);
      let k = s, n = 2;
      while (used.has(k.toLowerCase())) k = s.slice(0, 28) + ' (' + n++ + ')';
      used.add(k.toLowerCase());
      return k;
    };
    w.sheets[0].name = sheetName(tables[0], 0);
    for (let i = 1; i < tables.length; i++) w.addSheet(sheetName(tables[i], i));
    A.prepare(w);
    const b = A.addBook(w, name, { type: 'html' });
    tables.forEach((t, i) => {
      A.activateSheet(i);
      G().select(0, 0);
      L.clip.pasteHTML('<html><head>' + styles + '</head><body>' + t.outerHTML + '</body></html>');
    });
    A.activateSheet(0);
    G().select(0, 0);
    w.undo.clear();
    b.dirty = false;
    updateTitle();
  };

  /* ------------------------------------------------------------ save */
  A.save = async function () {
    const b = book();
    if (!b) return false;
    if (b.untitled || b.type === 'xls') return L.dlg.saveAs();
    return A.saveAs(b.type, b.name);
  };
  /** type: xlsx | xlsm | xltx | csv | txt | html | xmlss | pdf */
  A.saveAs = async function (type, name) {
    const b = book();
    if (L.editor.active && !L.editor.commit()) return false;
    const w = b.wb;
    name = name || b.name.replace(/\.[^.]+$/, '') + '.' + ({ xmlss: 'xml', html: 'htm', txt: 'txt' }[type] || type);
    ui.busy(true, L.t('正在保存 {0}...', name));
    try {
      let blob;
      if (type === 'csv' || type === 'txt') {
        if (w.sheets.length > 1) {
          const r = await (ui.busy(false), ui.msg(`The selected file type does not support workbooks that contain multiple sheets.\n\n• To save only the active sheet, click OK.\n• To save all sheets, save them individually using a different file name for each, or choose a file type that supports multiple sheets.`, { icon: 'warn', buttons: ['确定', '取消'] }));
          if (r !== 0) return false;
          ui.busy(true, '正在保存...');
        }
        blob = L.csv.write(w, sh(), type === 'csv' ? ',' : '\t');
      } else if (type === 'html') blob = L.csv.writeHTML(w);
      else if (type === 'xmlss') blob = L.xmlss.write(w);
      else if (type === 'pdf') { ui.busy(false); await L.print.exportPDF(); return true; }
      else {
        blob = await L.xlsxWrite.write(w, { type, password: b.password });
      }
      ui.busy(false);
      const losses = L.opc.formats[type] ? await ui.compatibility(w, blob.dropped) : [];
      if (losses === null) return false;
      const res = await L.saveFile(name, blob);
      if (res === 'saved') {
        L.opc.acknowledge(w, losses);
        if (type !== 'csv' && type !== 'txt' && type !== 'html') { b.name = name; b.type = type; b.untitled = false; b.dirty = false; w.fileName = name; }
        else if (b.untitled) b.dirty = false;
        if (window.VO && type !== 'html') VO.saved(name, blob, b);
        updateTitle();
        A.status('');
        return true;
      }
      return false;
    } catch (e) { ui.busy(false); A.error(e); return false; }
    finally { ui.busy(false); }
  };

  /* ------------------------------------------------------------ undo */
  A.undo = function () {
    if (L.editor.active) { L.editor.cancel(); return; }
    const w = wb();
    const t = w.undo.undo();
    if (t) afterUndo(t);
  };
  A.redo = function () {
    if (L.editor.active) return;
    const w = wb();
    const t = w.undo.redo();
    if (t) afterUndo(t);
  };
  function afterUndo(t) {
    const w = wb();
    /* the sheet the step happened on comes back into view */
    if (t.sheet != null && w.sheets[t.sheet] && w.active !== t.sheet && w.sheets[t.sheet].state === 'visible') w.active = t.sheet;
    O.recalc(w, null, true);
    for (const s of w.sheets) { LY.invalidate(s); LY.touchMerges(s); s._cfIdx = null; }
    renderTabs();
    G().paint(); G().syncObjects(true);
    L.editor.show();
    ui.refresh();
  }

  /* ------------------------------------------------------------ sheet tabs */
  function renderTabs() {
    const tabs = L.$('#tabs');
    const w = wb();
    const selSet = new Set(A.selectedSheetsIdx());
    /* the same sheets: only which tab is on changes, so a press that activates a tab keeps the element under the
       pointer, and its click and double-click (Rename) still reach it */
    const key = w.sheets.map((s) => s.state + '|' + s.name + '|' + JSON.stringify(s.tabColor || null)).join('\n');
    if (tabs._w === w && tabs._key === key && !tabs.querySelector('input')) {
      for (const t of L.$$('.stab', tabs)) {
        const i = t._i;
        t.classList.toggle('on', i === w.active);
        t.classList.toggle('sel', selSet.has(i) && i !== w.active);
        t.setAttribute('aria-selected', i === w.active ? 'true' : 'false');
      }
      scrollToActive(tabs);
      return;
    }
    tabs._w = w; tabs._key = key;
    tabs.textContent = '';
    w.sheets.forEach((s, i) => {
      if (s.state !== 'visible') return;
      const t = h('div', { class: 'stab' + (i === w.active ? ' on' : '') + (selSet.has(i) && i !== w.active ? ' sel' : '') + (s.tabColor ? ' colored' : ''), role: 'tab', 'aria-selected': i === w.active ? 'true' : 'false', draggable: 'false' }, h('span', { translate: 'no', text: s.name }));
      t._i = i;
      if (s.tabColor) t.style.setProperty('--tabc', M.colorHex(w, s.tabColor, '#ECE9D8'));
      t.addEventListener('pointerdown', (e) => {
        if (e.target.closest('input')) return;   // the name being edited in place
        if (e.button === 2) { if (!A.selectedSheetsIdx().includes(i)) A.activateSheet(i); return; }
        e.preventDefault();
        if (e.button !== 0) return;
        if (e.ctrlKey || e.metaKey) { A.toggleSheetSel(i); return; }
        if (e.shiftKey) { A.rangeSheetSel(i); return; }
        A.activateSheet(i);
        /* drag to reorder (Ctrl copies): a mark shows where it goes; a cancel moves nothing */
        let target = null;
        const unmark = () => L.$$('.stab', tabs).forEach((x) => x.classList.remove('drop-before'));
        ui.drag(e, {
          capture: t,   // the tab itself, so its click and double-click (Rename) come to it
          threshold: 6,
          move(ev) {
            unmark();
            target = null;
            for (const el of L.$$('.stab', tabs)) { const r = el.getBoundingClientRect(); if (ev.clientX < r.left + r.width / 2) { el.classList.add('drop-before'); target = el; break; } }
          },
          up(ev, d) {
            unmark();
            if (!d.moved) return;
            const to = target ? w.sheets.findIndex((s2) => s2.name === target.textContent) : w.sheets.length;
            try {
              if (ev.ctrlKey || ev.metaKey) { const c = O.copySheet(w, w.sheets[i], to); w.active = w.sheets.indexOf(c); }
              else O.moveSheet(w, w.sheets[i], to);
            } catch (err) { A.error(err); }
            renderTabs(); G().paint(); ui.refresh();
          },
          cancel: unmark,
        });
      });
      t.addEventListener('dblclick', () => A.renameSheetInline(i));
      t.addEventListener('contextmenu', (e) => { e.preventDefault(); tabMenu(e, i); });
      tabs.appendChild(t);
    });
    scrollToActive(tabs);
  }
  function scrollToActive(tabs) {
    const at = tabs.querySelector('.stab.on');
    if (at) { const r = at.getBoundingClientRect(), pr = tabs.getBoundingClientRect(); if (r.right > pr.right || r.left < pr.left) tabs.scrollLeft += r.left - pr.left - 20; }
  }
  A.renderTabs = renderTabs;
  let sheetSel = null; /* grouped sheets */
  A.selectedSheetsIdx = () => (sheetSel && sheetSel.size > 1 ? Array.from(sheetSel).sort((a, b) => a - b) : [wb().active]);
  A.selectedSheets = () => A.selectedSheetsIdx().map((i) => wb().sheets[i]);
  A.toggleSheetSel = (i) => { if (!sheetSel) sheetSel = new Set([wb().active]); if (sheetSel.has(i) && i !== wb().active) sheetSel.delete(i); else sheetSel.add(i); renderTabs(); updateTitle(); };
  A.rangeSheetSel = (i) => { const a = wb().active; sheetSel = new Set(); for (let k = Math.min(a, i); k <= Math.max(a, i); k++) if (wb().sheets[k].state === 'visible') sheetSel.add(k); renderTabs(); };
  A.activateSheet = function (i) {
    L.ui.dragCancel();   /* another sheet: a drag on this one ends */
    const w = wb();
    if (L.editor.active) {
      if (L.editor.canPoint()) { w.active = i; sheetSel = null; renderTabs(); G().paint(true); G().syncObjects(true); L.editor.sheetChanged(); return; }
      if (!L.editor.commit()) return;
    }
    if (!sheetSel || !sheetSel.has(i)) sheetSel = null;
    w.active = i;
    G().deselectObject();
    renderTabs();
    G().paint(true);
    G().syncObjects(true);
    L.editor.show();
    ui.refresh();
    L.bus.emit('sheet');
    updateStatus();
  };
  A.nextSheet = function (d) {
    const w = wb();
    let i = w.active;
    do { i += d; } while (i >= 0 && i < w.sheets.length && w.sheets[i].state !== 'visible');
    if (i >= 0 && i < w.sheets.length) A.activateSheet(i);
  };
  function tabMenu(e, i) {
    ui.openMenu([
      { label: '插入(&I)...', run: () => L.dlg.insertSheetDlg() },
      { label: '删除(&D)', run: () => A.deleteSheets() },
      { label: '重命名(&R)', run: () => A.renameSheetInline(i) },
      { label: '移动或复制(&M)...', run: () => L.dlg.moveCopySheet() },
      { label: '选定全部工作表(&S)', run: () => { sheetSel = new Set(wb().sheets.map((s, k) => (s.state === 'visible' ? k : -1)).filter((k) => k >= 0)); renderTabs(); } },
      ...(sheetSel && sheetSel.size > 1 ? [{ label: '取消组合工作表(&U)', run: () => { sheetSel = null; renderTabs(); } }] : []),
      { label: '工作表标签颜色(&T)...', run: () => L.dlg.tabColor() },
      '-',
      { label: '查看代码(&V)', disabled: true },
    ], { x: e.clientX, y: e.clientY });
  }
  A.renameSheetInline = function (i) {
    const w = wb();
    i = i == null ? w.active : i;
    const s = w.sheets[i];
    const tab = L.$$('.stab', L.$('#tabs')).find((x) => x.textContent === s.name);
    if (!tab) return;
    const inp = h('input', { type: 'text', value: s.name, style: 'position:relative;z-index:2;width:' + Math.max(60, tab.offsetWidth - 10) + 'px;height:15px;font:11px Tahoma,sans-serif;border:1px solid #000;padding:0 2px' });
    const span = tab.querySelector('span');
    span.replaceWith(inp);
    inp.focus(); inp.select();
    let done = false;
    const finish = (ok) => {
      if (done) return; done = true;
      if (ok && inp.value !== s.name) { try { O.renameSheet(w, s, inp.value.trim()); } catch (e) { A.error(e); } }
      renderTabs(); G().paint(); G().focus(); ui.refresh();
    };
    inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') finish(true); if (e.key === 'Escape') finish(false); });
    inp.addEventListener('blur', () => finish(true));
  };
  function buildTabNav() {
    const nav = L.$('#tabnav');
    const arrow = (d) => `<svg width="9" height="9" viewBox="0 0 9 9">${{ first: '<path d="M1.5 1v7" stroke="#000"/><path d="M7 1.5L3 4.5 7 7.5z" fill="#000"/>', prev: '<path d="M6.5 1.5L2.5 4.5 6.5 7.5z" fill="#000"/>', next: '<path d="M2.5 1.5L6.5 4.5 2.5 7.5z" fill="#000"/>', last: '<path d="M7.5 1v7" stroke="#000"/><path d="M2 1.5L6 4.5 2 7.5z" fill="#000"/>' }[d]}</svg>`;
    const tabs = L.$('#tabs');
    const b = (d, tip, fn) => { const x = h('button', { type: 'button', tabindex: '-1', 'aria-label': tip, 'data-tip': tip, html: arrow(d) }); x.addEventListener('click', fn); x.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu(wb().sheets.map((s, k) => (s.state === 'visible' ? { label: s.name.replace(/&/g, '&&'), raw: true, checked: () => k === wb().active, radio: true, run: () => A.activateSheet(k) } : null)).filter(Boolean), { x: e.clientX, y: e.clientY }); }); nav.appendChild(x); };
    b('first', '第一张工作表', () => { tabs.scrollLeft = 0; });
    b('prev', '上一张工作表', () => { tabs.scrollLeft -= 60; });
    b('next', '下一张工作表', () => { tabs.scrollLeft += 60; });
    b('last', '最后一张工作表', () => { tabs.scrollLeft = tabs.scrollWidth; });
    /* the split bar between tabs and the horizontal scroll bar */
    const sp = L.$('#tabsplit');
    sp.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      if (e.button !== 0) return;
      const x0 = e.clientX, w0 = tabs.offsetWidth, before = tabs.style.width;
      ui.drag(e, {
        capture: sp,
        move(ev) { tabs.style.width = L.clamp(w0 + ev.clientX - x0, 40, L.$('#tabrow').offsetWidth - 120) + 'px'; G().paint(); },
        cancel() { tabs.style.width = before; G().paint(); },
      });
    });
  }
  A.insertSheet = function () {
    const w = wb();
    const s = O.addSheet(w, null, w.active);
    w.active = w.sheets.indexOf(s);
    renderTabs(); G().paint(true); G().syncObjects(true); ui.refresh();
  };
  A.deleteSheets = async function () {
    const w = wb();
    const list = A.selectedSheets();
    const nonEmpty = list.some((s) => s.maxR >= 0 || s.drawings.length);
    if (nonEmpty) {
      const r = await ui.msg('选定要删除的工作表可能包含数据。要永久删除这些数据，请按“删除”。', { icon: 'warn', buttons: ['删除(&D)', '取消'] });
      if (r !== 0) return;
    }
    try { O.tx(w, '删除工作表', () => { for (const s of list) O.deleteSheet(w, s); }); } catch (e) { A.error(e); }
    sheetSel = null;
    renderTabs(); G().paint(true); G().syncObjects(true); ui.refresh();
  };
  A.hideSheets = function () {
    const w = wb();
    const list = A.selectedSheets();
    if (w.sheets.filter((s) => s.state === 'visible').length <= list.length) { ui.msg('工作簿必须至少包含一张可见的工作表。\n\n要隐藏、删除或移动选定工作表，请先插入一张新工作表，或取消隐藏一张已有工作表。', { icon: 'warn' }); return; }
    O.tx(w, '隐藏工作表', () => { for (const s of list) O.sheetProp(w, s, 'state', 'hidden'); });
    sheetSel = null;
    const vis = w.sheets.findIndex((s) => s.state === 'visible');
    if (w.sheets[w.active].state !== 'visible') w.active = vis;
    renderTabs(); G().paint(true); G().syncObjects(true);
  };

  /* ------------------------------------------------------------ menus */
  function buildMenus() {
    const recent = () => (window.VO ? VO.recentMenuItems(4) : []);
    const menus = [
      { label: '文件(&F)', items: () => ['newBook', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'pageSetup', { label: '打印区域(&t)', sub: ['setPrintArea', 'clearPrintArea'] }, 'printPreview', 'print', '-', { label: '发送到(&d)', sub: ['exportPDF', 'exportCSV'] }, 'properties', ...((recentList) => (recentList.length ? ['-', ...recentList] : []))(recent()), '-', 'exitApp'] },
      {
        label: '编辑(&E)', items: () => ['undo', 'redo', '-', 'cut', 'copy', 'officeClipboard', 'paste', 'pasteSpecial', 'pasteHyperlink', '-',
          { label: '填充(&i)', sub: ['fillDown', 'fillRight', 'fillUp', 'fillLeft', 'fillAcross', 'fillSeries', 'fillJustify'] },
          { label: '清除(&a)', sub: ['clearAll', 'clearFormats', 'clearContents', 'clearComments'] }, 'deleteCells', 'deleteSheet', 'moveCopySheet', '-', 'find', 'replace', 'goTo', '-', 'links'],
      },
      {
        label: '视图(&V)', items: () => ['viewNormal', 'viewPageBreak', '-', 'taskPane',
          { label: '工具栏(&T)', sub: ['tb_standard', 'tb_formatting', 'tb_borders', 'tb_chart', 'tb_drawing', 'tb_auditing', 'tb_reviewing', 'tb_picture', 'tb_list', '-', 'customize'] },
          'formulaBar', 'statusBar', '-', 'headerFooter', 'showComments', '-', 'fullScreen', 'zoomDlg'],
      },
      {
        label: '插入(&I)', items: () => ['insertCells', 'insertRows', 'insertCols', 'insertSheet', 'insertChart', 'insertSymbol', 'insertPageBreak', '-', 'insertFunction',
          { label: '名称(&N)', sub: ['nameDefine', 'pasteName', 'nameCreate', 'nameApply'] }, 'insertComment', '-',
          { label: '图片(&P)', icon: 'picture', sub: ['insertClipArt', 'insertPicture', '-', 'insertShape', 'insertTextBox'] }, 'hyperlink'],
      },
      {
        label: '格式(&o)', items: () => ['formatCells', { label: '行(&R)', sub: ['rowHeight', 'rowAutofit', 'rowHide', 'rowUnhide'] }, { label: '列(&C)', sub: ['colWidth', 'colAutofit', 'colHide', 'colUnhide', 'colStandard'] },
          { label: '工作表(&S)', sub: ['sheetRename', 'sheetHide', 'sheetUnhide', 'sheetBackground', 'tabColor'] }, 'autoFormat', 'conditionalFormat', 'styleDlg', ...(G().selectedObject() ? ['-', 'formatObject'] : [])],
      },
      {
        label: '工具(&T)', items: () => ['errorChecking', '-', { label: '保护(&P)', sub: ['protectSheet', 'allowEditRanges', 'protectWorkbook'] }, '-', 'goalSeek', 'scenarios',
          { label: '公式审核(&u)', sub: ['tracePrecedents', 'traceDependents', 'traceError', 'removeArrows', '-', 'evaluateFormula', '-', 'circleInvalid', 'clearCircles', '-', 'watchWindow', 'showFormulas', 'tb_auditing'] }, '-',
          'calcNow', { label: '宏(&M)', sub: ['macros'] }, 'addIns', 'autoCorrectDlg', 'customize', 'options'],
      },
      {
        label: '数据(&D)', items: () => ['sortDlg', { label: '筛选(&F)', sub: ['autoFilter', 'showAll', 'advancedFilter'] }, 'dataForm', 'subtotals', 'validation', '-', 'dataTable', 'textToColumns', 'consolidate',
          { label: '组及分级显示(&G)', sub: ['hideDetail', 'showDetail', '-', 'groupRows', 'ungroupRows', 'autoOutline', 'clearOutline', '-', 'outlineSettings'] }, '-', 'pivotTable',
          { label: '导入外部数据(&D)', sub: ['importText'] }, '-', { label: '列表(&L)', sub: ['createList', 'totalRow', 'convertToRange'] }, '-', 'refreshData'],
      },
      { label: '窗口(&W)', items: () => ['newWindow', 'arrangeWindows', '-', 'splitWindow', 'freezePanes', '-', ...A.books.map((b, i) => ({ label: `&${i + 1} ${b.name.replace(/&/g, '&&')}`, raw: true, radio: true, checked: () => i === A.cur, run: () => A.switchBook(i) }))] },
      { label: '帮助(&H)', items: ['help', 'showAssistant', '-', 'keyboardHelp', 'functionsHelp', '-', 'about'] },
    ];
    const row = L.$('#menurow');
    A.menuBar = ui.menuBar(row, menus);
    row.appendChild(ui.searchBox());
  }

  /* ------------------------------------------------------------ toolbars */
  const TB = {};
  /* the toolbars' boxes and buttons with menus, also used by the ribbon (js/ribbon.js): each call makes a new one */
  A.parts = {
    font: (o) => ui.combo(Object.assign({ id: 'fontname', width: 128, tip: '字体', value: () => { const st = curStyle(); return LY.fontName(wb(), st.font || {}); }, options: () => L.fontSections({ major: wb().theme?.major, minor: wb().theme?.minor }), preview: (v) => `font-family:${LY.fontStack(v)}`, note: L.fontNote, listCls: 'fonts', onChange: (v) => { if (G().guardProtect('formatCells')) return; L.addRecentFont(v); O.format(sh(), G().ranges(), { font: { name: v, scheme: undefined } }, L.N('字体')); G().paint(); } }, o)),
    size: (o) => ui.combo(Object.assign({ id: 'fontsize', width: 42, tip: '字号', value: () => String((curStyle().font || {}).sz || 10), options: () => L.SIZE_LIST, onChange: (v) => { const n = parseFloat(v); if (!(n >= 1 && n <= 409)) { ui.msg('数字必须在 1 到 409 之间。', { icon: 'warn' }); return; } if (G().guardProtect('formatCells')) return; O.format(sh(), G().ranges(), { font: { sz: n } }, '字号'); O.tx(wb(), L.N('自动调整'), () => { for (const r of A.selRows()) O.autoRow(sh(), r); }); G().paint(); } }, o)),
    zoom: (o) => ui.combo(Object.assign({ id: 'zoombox', width: 52, tip: '缩放', value: () => (sh().view.zoom || 100) + '%', options: () => ['200%', '100%', '75%', '50%', '25%', { value: '选区', label: L.t('选区') }], onChange: (v) => { if (/^sel/i.test(v)) { A.zoomToSelection(); return; } const n = parseInt(v, 10); if (n >= 10 && n <= 400) G().setZoom(n); else ui.msg('请输入 10 到 400 之间的数字。', { icon: 'warn' }); } }, o)),
  };
  A.menus = {
    autoSum: (at) => ui.openMenu([
      { label: '求和(&S)', run: () => A.autoSum('SUM') }, { label: '平均值(&A)', run: () => A.autoSum('AVERAGE') }, { label: '计数(&C)', run: () => A.autoSum('COUNT') }, { label: '最大值(&M)', run: () => A.autoSum('MAX') }, { label: '最小值(&i)', run: () => A.autoSum('MIN') }, '-', { label: '其他函数(&F)...', run: () => L.dlg.insertFunction() },
    ], at),
    fill: (at, el) => ui.colorMenu(el, { mode: 'fill', grid: true }, (c) => { if (c.none) A.applyFill(null); else if (typeof c === 'string') { A.lastFill = c; A.applyFill(c); } ui.refresh(); }),
    fontColor: (at, el) => ui.colorMenu(el, { mode: 'font', grid: true }, (c) => { if (c.auto) A.applyFontColor(null); else if (typeof c === 'string') { A.lastFontColor = c; A.applyFontColor(c); } ui.refresh(); }),
    borders: (at, el) => borderPalette(el),
    draw: (at) => ui.openMenu([{ label: '置于顶层(&B)', run: () => L.drawing.order('front') }, { label: '置于底层(&S)', run: () => L.drawing.order('back') }], at),
  };
  function buildToolbars() {
    const host = L.$('#toolbars');
    host.textContent = '';
    const fontCombo = A.parts.font(), sizeCombo = A.parts.size(), zoomCombo = A.parts.zoom();
    const sumSplit = ui.tbSplit('autoSum', A.menus.autoSum);
    const fillSplit = ui.tbSplit('fillColorApply', A.menus.fill);
    const fontSplit = ui.tbSplit('fontColorApply', A.menus.fontColor);
    const borderSplit = ui.tbSplit('bordersApply', A.menus.borders);
    TB.standard = ui.toolbar('standard', '常用', ['newBook', 'open', 'save', '|', 'printQuick', 'printPreview', '|', 'cut', 'copy', 'paste', 'formatPainter', '|', 'undo', 'redo', '|', 'hyperlink', sumSplit, 'sortAsc', 'sortDesc', '|', 'chartWizard', ui.tbButton('tb_drawing', { icon: 'drawing' }), zoomCombo, '|', 'help']);
    TB.formatting = ui.toolbar('formatting', '格式设置', [fontCombo, sizeCombo, '|', 'bold', 'italic', 'underline', '|', 'alignLeft', 'alignCenter', 'alignRight', 'mergeCenter', '|', 'currencyStyle', 'percentStyle', 'commaStyle', 'incDecimal', 'decDecimal', '|', 'decIndent', 'incIndent', '|', borderSplit, fillSplit, fontSplit]);
    TB.borders = ui.toolbar('borders', 'Borders', [ui.tbSplit('bordersApply', A.menus.borders), ui.tbButton('borderNone', { icon: 'bordersNone' }), ui.tbButton('borderOutline', { icon: 'bordersOut' })]);
    TB.auditing = ui.toolbar('auditing', '公式审核', ['traceError', '|', 'tracePrecedents', 'traceDependents', 'removeArrows', '|', 'insertComment', 'validation', 'circleInvalid', 'clearCircles', '|', 'watchWindow', 'evaluateFormula']);
    TB.reviewing = ui.toolbar('reviewing', '审阅', ['insertComment', 'showHideComment', 'deleteComment', '|', 'showComments']);
    TB.list = ui.toolbar('list', '列表', [ui.tbDrop('列表(&L)', null, (at) => ui.openMenu(['insertRows', 'insertCols', '-', 'sortDlg', 'autoFilter', '-', 'totalRow', 'convertToRange'], at)), 'totalRow']);
    TB.chart = ui.toolbar('chart', '图表', [ui.tbDrop('图表类型', 'chart', (at) => L.chartWizard && L.chartWizard.typeMenu(at)), ui.tbButton('formatObject', { label: '格式' })]);
    TB.drawing = ui.toolbar('drawing', '绘图', [ui.tbDrop('绘制(&r)', 'draw', A.menus.draw), '|',
      ui.tbButton('insertShape', { icon: 'rect' }), ui.tbButton('insertTextBox'), ui.tbButton('insertPicture'), '|', ui.tbButton('fillColorApply')]);
    const row1 = h('div', { class: 'tb-row' }), row2 = h('div', { class: 'tb-row' });
    row1.append(TB.standard, TB.formatting);
    row2.append(TB.borders, TB.auditing, TB.reviewing, TB.list, TB.chart, TB.drawing);
    host.append(row1, row2);
    A.updateToolbars();
  }
  function borderPalette(el) {
    const opts = [['none', 'bordersNone', '无框线'], ['bottom', 'bordersBottom', '下框线'], ['left', 'bordersLeft', '左框线'], ['right', 'bordersRight', '右框线'],
      ['bottomDouble', 'bordersBottomDouble', '双下框线'], ['thickBottom', 'bordersThickBottom', '粗下框线'], ['topBottom', 'bordersTopBottom', '上下框线'], ['topThickBottom', 'bordersTopThickBottom', '上框线和粗下框线'],
      ['topDoubleBottom', 'bordersTopDoubleBottom', '上框线和双下框线'], ['all', 'bordersAll', '所有框线'], ['outside', 'bordersOut', '外侧框线'], ['thickBox', 'bordersThickBox', '粗匣框线']];
    const r = el.getBoundingClientRect();
    ui.openMenu([{ custom: (close) => {
      const g = h('div', { class: 'border-grid', style: 'grid-template-columns:repeat(4,24px)' });
      for (const [k, ic, tip] of opts) { const b = h('button', { type: 'button', 'data-tip': tip, 'aria-label': tip, html: L.icons.get(ic) }); b.addEventListener('click', () => { close(); A.lastBorder = k; A.applyBorderPreset(k); }); g.appendChild(b); }
      const box = h('div', null, g, h('button', { class: 'cp-wide', type: 'button', style: 'width:100%', text: '绘制边框...', onclick: () => { close(); L.dlg.formatCells(3); } }));
      return box;
    } }], { left: r.left, bottom: r.bottom }, { cls: 'grid-menu' });
  }
  A.toolbarVisible = (id) => { const v = A.opts.toolbars[id]; if (v === 'auto') return id === 'chart' ? !!(G().selectedObject() && G().selectedObject().kind === 'chart') : id === 'picture' ? !!(G().selectedObject() && G().selectedObject().kind === 'image') : id === 'list' ? !!A.tableAt() : false; return !!v; };
  /** what is selected, for the ribbon's contextual tabs: 'chart', 'picture', 'shape', 'list' */
  A.inContext = (k) => {
    const o = G().selectedObject();
    if (k === 'list') return !!A.tableAt();
    return !!o && (k === 'chart' ? o.kind === 'chart' : k === 'picture' ? o.kind === 'image' : k === 'shape' ? o.kind === 'shape' : false);
  };
  A.toggleToolbar = (id, force) => { const cur = A.toolbarVisible(id); A.opts.toolbars[id] = force != null ? force : !cur; A.updateToolbars(); A.saveOpts(); };
  A.updateToolbars = function () {
    for (const id in TB) { const v = A.toolbarVisible(id); TB[id].hidden = !v; }
    const r2 = TB.borders && TB.borders.parentElement;
    if (r2) r2.hidden = !Array.from(r2.children).some((x) => !x.hidden);
    ui.refresh();
  };
  const curStyle = () => { const s = G().sel(); return LY.styleOf(sh(), s.r, s.c, sh().get(s.r, s.c)); };

  /* ------------------------------------------------------------ status bar */
  let sbMode, sbText, sbCalc, sbInd = {};
  function buildStatus() {
    const sb = L.$('#statusbar');
    sbMode = h('div', { class: 'sb-cell mode', text: L.N('就绪') });
    sbText = h('div', { class: 'sb-cell grow' });
    sbCalc = h('div', { class: 'sb-cell calc', 'data-tip': '右键单击可选择自动计算函数' });
    sbCalc.addEventListener('contextmenu', (e) => { e.preventDefault(); calcMenu(e); });
    sbCalc.addEventListener('click', (e) => calcMenu(e));
    sb.append(sbMode, sbText, sbCalc);
    /* the keys stay CAPS/NUM/SCRL/END — the code below reads sbInd.CAPS and friends; only the label is Chinese */
    for (const [k, label] of [['CAPS', '大写'], ['NUM', '数字'], ['SCRL', '滚动'], ['END', '结束']]) { sbInd[k] = h('div', { class: 'sb-cell ind off', text: label }); sb.appendChild(sbInd[k]); }
    document.addEventListener('keydown', (e) => { if (e.getModifierState) { sbInd.CAPS.classList.toggle('off', !e.getModifierState('CapsLock')); sbInd.NUM.classList.toggle('off', !e.getModifierState('NumLock')); sbInd.SCRL.classList.toggle('off', !e.getModifierState('ScrollLock')); G().scrollLock = e.getModifierState('ScrollLock'); } }, true);
  }
  function calcMenu(e) {
    const fns = [['NONE', '无(&N)'], ['AVERAGE', '平均值(&A)'], ['COUNTA', '计数(&C)'], ['COUNT', '数值计数(&o)'], ['MAX', '最大值(&M)'], ['MIN', '最小值(&i)'], ['SUM', '求和(&S)']];
    ui.openMenu(fns.map(([k, l]) => ({ label: l, radio: true, checked: () => A.opts.autoCalcFn === k, run: () => { A.opts.autoCalcFn = k; A.saveOpts(); updateStatus(); } })), { x: e.clientX, y: e.clientY - 150 });
  }
  /* circular references: "Circular: B4" in the status bar, and Excel's warning the first time one appears */
  let warnedCircular = new WeakSet();
  function circularCells() {
    const w = wb();
    if (!w || !C.circular || !C.circular.size || w.calcPr.iterate) return [];
    const out = [];
    for (const n of w.graph.byCell.values()) if (C.circular.has(n.cell)) out.push(n);
    return out;
  }
  function circularText() {
    const list = circularCells();
    if (!list.length) return '';
    const n = list.find((x) => x.sh === sh()) || list[0];
    return L.t('循环引用：{0}', (n.sh === sh() ? '' : F.quoteSheet(n.sh.name) + '!') + F.cellName(n.r, n.c));
  }
  L.bus.on('changed', () => {
    const fresh = circularCells().filter((n) => !warnedCircular.has(n.cell));
    if (!fresh.length) return;
    for (const n of circularCells()) warnedCircular.add(n.cell);
    setTimeout(() => ui.msg('Ledger 无法计算公式。公式中的单元格引用指向该公式的结果，形成了循环引用。\n\n请删除或更改该引用，或启用迭代计算（“工具 ▸ 选项 ▸ 计算”）。', { icon: 'warn' }), 0);
  });
  const updateStatus = L.debounce(() => {
    if (!sbMode || !G().wb) return;
    const ed = L.editor;
    sbMode.textContent = ed.active ? (ed.mode === 'point' ? L.t('点选') : ed.mode === 'edit' ? L.t('编辑') : L.t('输入')) : L.t('就绪');
    sbText.textContent = statusText || (wb().calcPr.mode === 'manual' && wb().needsCalc ? L.t('计算') : '') || (A.painter.on() ? L.t('单击单元格以应用格式') : '') || circularText();
    sbInd.END.classList.toggle('off', !G().endMode());
    /* AutoCalculate */
    const fn = A.opts.autoCalcFn || 'SUM';
    const s = G().sel();
    const multi = s.ranges.length > 1 || s.ranges.some((r) => r.r1 !== r.r2 || r.c1 !== r.c2);
    if (!multi || fn === 'NONE') { sbCalc.textContent = ''; return; }
    let sum = 0, cnt = 0, cnta = 0, mx = -Infinity, mn = Infinity, err = false, guard = 0;
    for (const rg of s.ranges) {
      const rr = O.clip(sh(), rg);
      sh().each(rr.r1, rr.c1, rr.r2, rr.c2, (cell, r, c) => {
        if (++guard > 200000) return false;
        const row = sh().rows[r];
        if (row && row.hidden && row.filtered) return;
        const v = cell.dirty ? C.cellValue(sh(), r, c) : cell.v;
        if (v == null || v === '') return;
        cnta++;
        if (M.isErr(v)) err = true;
        if (typeof v === 'number') { sum += v; cnt++; mx = Math.max(mx, v); mn = Math.min(mn, v); }
      });
    }
    const st = curStyle();
    const show = (v) => (typeof v === 'number' ? NF.text(st.nf && !/^General$/i.test(st.nf) ? st.nf : 'General', v, { date1904: wb().date1904 }) : v);
    let t = '';
    if (err && fn !== 'COUNTA' && fn !== 'COUNT') t = '';
    else if (fn === 'SUM') t = L.t('求和={0}', show(NF.r15(sum)));
    else if (fn === 'AVERAGE') t = cnt ? L.t('平均值={0}', show(NF.r15(sum / cnt))) : '';
    else if (fn === 'COUNTA') t = L.t('计数={0}', cnta);
    else if (fn === 'COUNT') t = L.t('数值计数={0}', cnt);
    else if (fn === 'MAX') t = L.t('最大值={0}', show(cnt ? mx : 0));
    else if (fn === 'MIN') t = L.t('最小值={0}', show(cnt ? mn : 0));
    sbCalc.textContent = t;
  }, 60);
  A.updateStatus = updateStatus;

  /* ------------------------------------------------------------ name box */
  function buildNameBox() {
    const nb = L.$('#namebox');
    const show = () => {
      if (document.activeElement === nb) return;
      const s = G().sel(), w = wb();
      const rg = s.ranges[s.active];
      const d = G().selectedObject();
      if (d) { nb.value = d.name || L.t('对象'); return; }
      /* a defined name that matches the selection exactly is shown instead of the address */
      const exact = w.names.find((n) => !n.hidden && (n.scope == null || n.scope === w.active) && (() => { try { const a = F.parse(n.ref); return a.t === 'area' || a.t === 'ref' ? sameRange(a, rg) : false; } catch (e) { return false; } })());
      if (exact) { nb.value = exact.name; return; }
      if (G().state.fillTo || (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) && L.editor.active === false && A._dragging) { nb.value = L.t('{0}R x {1}C', rg.r2 - rg.r1 + 1, rg.c2 - rg.c1 + 1); return; }
      nb.value = w.r1c1 ? 'R' + (s.r + 1) + 'C' + (s.c + 1) : F.cellName(s.r, s.c);
    };
    const sameRange = (a, rg) => { if (a.sheet && a.sheet.toLowerCase() !== sh().name.toLowerCase()) return false; if (a.t === 'ref') return rg.r1 === a.r && rg.r2 === a.r && rg.c1 === a.c && rg.c2 === a.c; return rg.r1 === Math.min(a.r1, a.r2) && rg.r2 === Math.max(a.r1, a.r2) && rg.c1 === Math.min(a.c1, a.c2) && rg.c2 === Math.max(a.c1, a.c2); };
    L.bus.on('selection', show); L.bus.on('sheet', show); L.bus.on('book', show); L.bus.on('changed', show);
    /* the Name Box takes the keyboard: what is being typed in a cell is entered first, as a click elsewhere does */
    nb.addEventListener('focus', () => { if (L.editor.active && !L.editor.commit({ keepFocus: true })) { nb.blur(); return; } nb.select(); });
    nb.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { nb.blur(); show(); G().focus(); return; }
      if (e.key !== 'Enter') return;
      e.preventDefault();
      const t = nb.value.trim();
      nb.blur();
      if (!t) { show(); G().focus(); return; }
      if (!A.goToRef(t)) {
        /* a new name for the selection */
        if (/^[A-Za-z_\\][\w.\\]*$/.test(t) && !F.parseCell(t)) {
          const s = G().sel(), rg = s.ranges[s.active];
          O.setNames(wb(), wb().names.concat([{ name: t, ref: F.quoteSheet(sh().name) + '!' + F.absRangeName(rg), scope: null }]));
          O.recalc(wb(), null, true);
          show();
        } else ui.msg('引用无效。', { icon: 'warn' });
      }
      G().focus();
    });
    L.$('#namebox-dd').addEventListener('pointerdown', (e) => {
      e.preventDefault();
      const r = L.$('#namebox-wrap').getBoundingClientRect();
      const names = wb().names.filter((n) => !n.hidden && !/^_xlnm\./.test(n.name) && (n.scope == null || n.scope === wb().active));
      const tables = wb().tables.map((t) => t.name);
      ui.openMenu(names.length || tables.length ? names.map((n) => ({ label: n.name.replace(/&/g, '&&'), raw: true, run: () => A.goToRef(n.name) })).concat(tables.map((t) => ({ label: t.replace(/&/g, '&&'), raw: true, run: () => A.goToRef(t) }))) : [{ label: '（无名称）', disabled: true }], { left: r.left, bottom: r.bottom });
    });
  }
  /** select a reference typed in the Name Box or Go To: A1, A1:B5, Sheet2!C3, a name or a table */
  A.goToRef = function (t) {
    const w = wb();
    let target = null;
    try {
      const ast = F.parse(t.replace(/^=/, ''));
      const ctx = C.ctx(w, sh(), G().sel().r, G().sel().c);
      let v = C.ev(ast, ctx);
      if (v instanceof C.RefList) v = v.list[0];
      if (v instanceof C.Ref) target = v;
    } catch (e) { target = null; }
    if (!target) {
      /* R1C1 style */
      const m = /^R(\d+)C(\d+)$/i.exec(t);
      if (m) target = new C.Ref(sh(), +m[1] - 1, +m[2] - 1);
    }
    if (!target) return false;
    const i = w.sheets.indexOf(target.sheet);
    if (i < 0) return false;
    if (i !== w.active) A.activateSheet(i);
    G().selectRange({ r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2 }, { r: target.r1, c: target.c1 });
    A.lastGoTo = (A.lastGoTo || []).filter((x) => x !== t); A.lastGoTo.unshift(t);
    return true;
  };

  /* ------------------------------------------------------------ keyboard */
  A.SHORTCUTS = {};
  function keyName(e) {
    let k = e.key;
    if (k === ' ') k = 'Space';
    if (k.length === 1) k = k.toUpperCase();
    const parts = [];
    if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
    if (e.altKey) parts.push('Alt');
    if (e.shiftKey) parts.push('Shift');
    parts.push(k);
    return parts.join('+');
  }
  function buildShortcuts() {
    for (const id in ui.cmds) {
      const c = ui.cmds[id];
      if (!c.key) continue;
      let k = c.key.replace('Del', 'Delete');
      A.SHORTCUTS[k.split('+').map((p) => (p.length === 1 ? p.toUpperCase() : p)).join('+')] = id;
    }
    Object.assign(A.SHORTCUTS, {
      'Ctrl+Shift+~': 'fmtGeneral', 'Ctrl+Shift+`': 'fmtGeneral', 'Ctrl+Shift+$': 'fmtCurrency', 'Ctrl+Shift+4': 'fmtCurrency', 'Ctrl+Shift+%': 'percentStyle', 'Ctrl+Shift+5': 'percentStyle', 'Ctrl+Shift+^': 'fmtExp', 'Ctrl+Shift+6': 'fmtExp', 'Ctrl+Shift+#': 'fmtDate', 'Ctrl+Shift+3': 'fmtDate', 'Ctrl+Shift+@': 'fmtTime', 'Ctrl+Shift+2': 'fmtTime', 'Ctrl+Shift+!': 'fmtNumber', 'Ctrl+Shift+1': 'fmtNumber',
      'Ctrl+Shift+&': 'borderOutline', 'Ctrl+Shift+7': 'borderOutline', 'Ctrl+Shift+_': 'borderNone', 'Ctrl+Shift+-': 'borderNone',
      'Ctrl+-': 'deleteCells', 'Ctrl+Shift++': 'insertCells', 'Ctrl+Shift+=': 'insertCells', 'Ctrl++': 'insertCells', 'Ctrl+=': 'calcNow',
      'Ctrl+;': 'insertDate', 'Ctrl+Shift+:': 'insertTime', 'Ctrl+Shift+;': 'insertTime', "Ctrl+'": 'copyAbove', 'Ctrl+Shift+"': 'copyValueAbove', "Ctrl+Shift+'": 'copyValueAbove', 'Ctrl+`': 'showFormulas',
      'Ctrl+9': 'rowHide', 'Ctrl+0': 'colHide', 'Ctrl+Shift+(': 'rowUnhide', 'Ctrl+Shift+9': 'rowUnhide', 'Ctrl+Shift+)': 'colUnhide', 'Ctrl+Shift+0': 'colUnhide',
      'Ctrl+5': 'strike', 'Ctrl+2': 'bold', 'Ctrl+3': 'italic', 'Ctrl+4': 'underline', 'Ctrl+1': 'formatCells', 'Ctrl+Shift+F': 'formatCells', 'Ctrl+Shift+P': 'formatCells',
      'Ctrl+Shift+*': 'selectRegion', 'Ctrl+Shift+8': 'selectRegion', 'Ctrl+*': 'selectRegion', 'Ctrl+/': 'selectArray',
      'Alt+=': 'autoSum', 'Alt+Shift+=': 'autoSum', 'F4': 'repeatLast', 'Ctrl+Y': 'redo', 'F5': 'goTo', 'Ctrl+G': 'goTo', 'Shift+F5': 'find', 'Shift+F4': 'find',
      'F11': 'insertChart', 'Alt+F1': 'insertChart', 'Shift+F11': 'insertSheet', 'Alt+Shift+F1': 'insertSheet', 'F9': 'calcNow', 'Shift+F9': 'calcSheet', 'Ctrl+Alt+F9': 'calcNow',
      'F12': 'saveAs', 'Ctrl+F12': 'open', 'Ctrl+Shift+F12': 'print', 'Ctrl+F2': 'printPreview', 'Ctrl+F1': 'taskPane', 'F1': 'help', 'Shift+F10': 'contextKey', 'ContextMenu': 'contextKey',
      'Ctrl+K': 'hyperlink', 'Ctrl+L': 'createList', 'Ctrl+F3': 'nameDefine', 'Ctrl+Shift+F3': 'nameCreate', 'Shift+F3': 'insertFunction', 'Shift+F2': 'insertComment',
      'Alt+Shift+ArrowRight': 'groupRows', 'Alt+Shift+ArrowLeft': 'ungroupRows', 'Ctrl+8': 'outlineSymbolsToggle', 'Ctrl+PageDown': 'nextSheet', 'Ctrl+PageUp': 'prevSheet', 'Ctrl+W': 'close', 'Ctrl+F4': 'close',
    });
    delete A.SHORTCUTS['Ctrl+C']; delete A.SHORTCUTS['Ctrl+X']; delete A.SHORTCUTS['Ctrl+V'];
    delete A.SHORTCUTS['Delete']; delete A.SHORTCUTS['Alt+ArrowDown'];
  }
  /** grid key handler asks here first */
  A.shortcut = function (e) {
    const k = keyName(e);
    if (k === 'Ctrl+A') { A.selectAllSmart(); return true; }
    const id = A.SHORTCUTS[k];
    if (!id || !ui.cmds[id]) return false;
    ui.exec(id);
    return true;
  };
  document.addEventListener('keydown', (e) => {
    if (ui.dialogOpen()) { ui.dialogKey(e); return; }
    if (ui.menuOpen()) { if (ui.menuKey(e)) { e.preventDefault(); e.stopPropagation(); } return; }
    if ((e.key === 'F10' && !e.shiftKey) || (e.key === 'Alt' && false)) { e.preventDefault(); A.menuBar.openIndex(0); return; }
    if (e.altKey && !e.ctrlKey && !e.shiftKey && e.key.length === 1 && /[a-z]/i.test(e.key)) { if (A.menuBar.openByKey(e.key)) { e.preventDefault(); return; } }
    const t = e.target;
    if (t === G().input || t === L.$('#fbinput')) return; /* the grid / editor handle their own keys */
    const inInput = t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable);
    if (inInput) { if (e.key === 'Escape') { t.blur(); G().focus(); } return; }
    if (A.shortcut(e)) { e.preventDefault(); return; }
    /* stray typing goes to the grid */
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) G().focus();
  });

  /* ------------------------------------------------------------ context menus */
  A.contextMenu = function (e, z) {
    const s = G().sel();
    const d = G().selectedObject();
    if (d) { L.drawing.contextMenu(e, d); return; }
    const z0 = z || {};
    let items;
    if (z0.kind === 'colhdr' || (G().isWholeCols(s.ranges[0]) && z0.kind !== 'rowhdr' && z0.kind !== 'cell')) items = ['cut', 'copy', 'paste', 'pasteSpecial', '-', { label: '插入(&I)', run: () => A.insertLines('c') }, { label: '删除(&D)', run: () => A.deleteLines('c') }, { label: '清除内容(&n)', cmd: 'clearContents' }, '-', 'formatCells', 'colWidth', 'colHide', 'colUnhide'];
    else if (z0.kind === 'rowhdr' || G().isWholeRows(s.ranges[0])) items = ['cut', 'copy', 'paste', 'pasteSpecial', '-', { label: '插入(&I)', run: () => A.insertLines('r') }, { label: '删除(&D)', run: () => A.deleteLines('r') }, { label: '清除内容(&n)', cmd: 'clearContents' }, '-', 'formatCells', 'rowHeight', 'rowHide', 'rowUnhide'];
    else {
      const cm = sh().comments.get(M.key(s.r, s.c));
      const link = sh().links.find((l) => M.rangeContains(l.ref, s.r, s.c));
      items = ['cut', 'copy', 'paste', 'pasteSpecial', '-', { label: '插入(&I)...', cmd: 'insertCells' }, { label: '删除(&D)...', cmd: 'deleteCells' }, { label: '清除内容(&n)', cmd: 'clearContents' }, '-',
        ...(cm ? ['insertComment', 'deleteComment', 'showHideComment'] : ['insertComment']), '-', 'formatCells', 'pickList', 'addWatch', ...(A.tableAt() ? ['-', { label: '列表(&L)', sub: ['totalRow', 'convertToRange'] }] : ['createList']), '-',
        ...(link ? [{ label: '编辑超链接(&E)...', cmd: 'hyperlink' }, { label: '打开超链接(&O)', run: () => A.followLink(link) }, 'removeHyperlink'] : ['hyperlink'])];
    }
    ui.openMenu(items, { x: e.clientX, y: e.clientY });
  };
  A.contextMenuAtSel = function () {
    const s = G().sel();
    const q = G().cellRect(s.r, s.c);
    const w = L.$('#gridwrap').getBoundingClientRect();
    A.contextMenu({ clientX: w.left + q.x + q.w / 2, clientY: w.top + q.y + q.h / 2, preventDefault() {} }, { kind: 'cell', r: s.r, c: s.c });
  };

  /* ------------------------------------------------------------ cell actions */
  A.afterEntry = function (r, c) {
    G().paint();
    updateStatus();
    void r; void c;
  };
  A.followLink = function (l) {
    if (l.location) { if (!A.goToRef(l.location)) ui.msg('引用无效。', { icon: 'warn' }); return; }
    if (!l.target) return;
    if (/^#/.test(l.target)) { A.goToRef(l.target.slice(1)); return; }
    if (/^(https?:|mailto:|ftp:)/i.test(l.target)) { const a = h('a', { href: l.target, target: '_blank', rel: 'noopener noreferrer' }); document.body.appendChild(a); a.click(); a.remove(); return; }
    ui.msg(L.t('无法打开指定的文件。\n\n{0}', l.target), { icon: 'warn' });
  };
  /** Alt+Down: validation list, else the column's distinct entries */
  A.pickList = function () {
    const s = G().sel(), shx = sh();
    if (G().guardProtect(null, true)) return;
    const dv = L.cf.dvAt(shx, s.r, s.c);
    let items;
    if (dv && dv.type === 'list') items = L.cf.dvList(shx, dv, s.r, s.c);
    else {
      const set = new Map();
      const scan = (dir) => { let r = s.r + dir; while (r >= 0 && r < MAXR) { const cl = shx.get(r, s.c); if (!cl || cl.v == null || cl.v === '') break; if (typeof cl.v === 'string') set.set(cl.v.toLowerCase(), cl.v); r += dir; } };
      scan(-1); scan(1);
      items = Array.from(set.values()).sort((a, b) => C.strCmp(a, b));
    }
    if (!items.length) return;
    L.editor.pickList(items, (v) => { O.tx(wb(), L.N('键入'), () => O.enter(shx, s.r, s.c, typeof v === 'number' ? String(v) : String(v).replace(/^=/, "'="))); G().paint(); L.editor.show(); });
  };
  /** double-click on the fill handle: fill down as far as the column to the left has data */
  A.fillDownToData = function () {
    const s = G().sel(), shx = sh();
    const rg = s.ranges[0];
    const has = (r, c) => { const cl = shx.get(r, c); return cl && (cl.v != null || cl.f != null); };
    let last = rg.r2;
    const probe = rg.c1 > 0 && has(rg.r2 + 1, rg.c1 - 1) ? rg.c1 - 1 : has(rg.r2 + 1, rg.c2 + 1) ? rg.c2 + 1 : -1;
    if (probe < 0) return;
    while (last + 1 < MAXR && has(last + 1, probe) && !has(last + 1, rg.c1)) last++;
    if (last === rg.r2) return;
    O.autoFill(shx, rg, { r1: rg.r1, c1: rg.c1, r2: last, c2: rg.c2 }, 'auto');
    G().selectRange({ r1: rg.r1, c1: rg.c1, r2: last, c2: rg.c2 }, { r: rg.r1, c: rg.c1 });
  };
  A.fillDir = function (dir) {
    const shx = sh();
    // Resolve destinations before checking protection or changing the used extent.
    const ranges = G().ranges().map(rg0 => {
      const rg = { ...rg0 }, used = O.clip(shx, rg0);
      if (rg.r1 === 0 && rg.r2 === MAXR - 1) rg.r2 = used.r2;
      if (rg.c1 === 0 && rg.c2 === MAXC - 1) rg.c2 = used.c2;
      return rg;
    });
    if (G().guardProtect(null, true, ranges)) return;
    O.tx(wb(), { down: '向下填充', up: '向上填充', right: '向右填充', left: '向左填充' }[dir], () => {
      for (const rg of ranges) {
        if (dir === 'down') { const src = rg.r1 === rg.r2 ? { r1: rg.r1 - 1, c1: rg.c1, r2: rg.r1 - 1, c2: rg.c2 } : { r1: rg.r1, c1: rg.c1, r2: rg.r1, c2: rg.c2 }; if (src.r1 < 0) continue; O.autoFill(shx, src, { r1: src.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, 'copy'); }
        if (dir === 'up') { const src = rg.r1 === rg.r2 ? { r1: rg.r2 + 1, c1: rg.c1, r2: rg.r2 + 1, c2: rg.c2 } : { r1: rg.r2, c1: rg.c1, r2: rg.r2, c2: rg.c2 }; O.autoFill(shx, src, { r1: rg.r1, c1: rg.c1, r2: src.r2, c2: rg.c2 }, 'copy'); }
        if (dir === 'right') { const src = rg.c1 === rg.c2 ? { r1: rg.r1, c1: rg.c1 - 1, r2: rg.r2, c2: rg.c1 - 1 } : { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c1 }; if (src.c1 < 0) continue; O.autoFill(shx, src, { r1: rg.r1, c1: src.c1, r2: rg.r2, c2: rg.c2 }, 'copy'); }
        if (dir === 'left') { const src = rg.c1 === rg.c2 ? { r1: rg.r1, c1: rg.c2 + 1, r2: rg.r2, c2: rg.c2 + 1 } : { r1: rg.r1, c1: rg.c2, r2: rg.r2, c2: rg.c2 }; O.autoFill(shx, src, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: src.c2 }, 'copy'); }
      }
    });
  };
  /** Edit ▸ Fill ▸ Justify: rejoin text in the first column and wrap it to the selection width */
  A.justifyFill = function () {
    const shx = sh(), rg = G().ranges()[0];
    const words = [];
    for (let r = rg.r1; r <= rg.r2; r++) { const v = shx.val(r, rg.c1); if (typeof v === 'string') words.push(...v.split(/\s+/).filter(Boolean)); else if (v != null) { ui.msg('无法对包含数字或公式的单元格进行两端对齐。', { icon: 'warn' }); return; } }
    const g = LY.geo(shx);
    let width = 0; for (let c = rg.c1; c <= rg.c2; c++) width += g.cols.size(c);
    const st = LY.styleOf(shx, rg.r1, rg.c1, shx.get(rg.r1, rg.c1));
    const font = LY.cssFont(Object.assign({}, st.font, { name: LY.fontName(wb(), st.font || {}) }), 1);
    const lines = L.render.wrapLines(font, words.join(' '), width - 6);
    O.tx(wb(), L.N('两端对齐'), () => {
      for (let r = rg.r1; r <= Math.max(rg.r2, rg.r1 + lines.length - 1); r++) O.setValue(shx, r, rg.c1, lines[r - rg.r1] != null ? lines[r - rg.r1] : null);
    });
  };
  A.copyFromAbove = function (valueOnly) {
    const s = G().sel(), shx = sh();
    if (s.r === 0) return;
    const above = shx.get(s.r - 1, s.c);
    let text = '';
    if (above) {
      if (!valueOnly && above.f != null) text = '=' + F.display(above.f);
      else text = L.editor.cellText(shx, s.r - 1, s.c).replace(/^=/, valueOnly ? '' : '=');
      if (valueOnly && above.f != null) { const v = above.v; text = v == null ? '' : typeof v === 'number' ? L.editor.barNumber(v) : String(M.isErr(v) ? v.e : v); }
    }
    L.editor.begin('edit', text);
  };
  A.selectAllSmart = function () {
    const s = G().sel(), shx = sh();
    const reg = O.currentRegion(shx, s.r, s.c);
    const rg = s.ranges[0];
    const single = s.ranges.length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2;
    const hasData = shx.get(s.r, s.c) || reg.r1 !== reg.r2 || reg.c1 !== reg.c2;
    if (single && hasData && !(reg.r1 === rg.r1 && reg.r2 === rg.r2 && reg.c1 === rg.c1 && reg.c2 === rg.c2)) G().selectRange(reg, { r: s.r, c: s.c }, { noScroll: true });
    else G().selectAll();
  };
  A.selRows = function (unhide) {
    const out = new Set();
    for (const rg of G().ranges()) {
      let r1 = rg.r1, r2 = Math.min(rg.r2, Math.max(sh().maxR + 1, rg.r1));
      if (unhide && r1 === r2) { /* unhide around a single row */ }
      for (let r = r1; r <= r2 && out.size < 100000; r++) out.add(r);
    }
    return Array.from(out);
  };
  A.selCols = function () {
    const out = new Set();
    for (const rg of G().ranges()) for (let c = rg.c1; c <= Math.min(rg.c2, Math.max(sh().maxC + 1, rg.c1)); c++) out.add(c);
    return Array.from(out);
  };

  /* insert / delete */
  A.insertLines = function (axis) {
    if (G().guardProtect(axis === 'r' ? 'insertRows' : 'insertColumns')) return;
    const rg = G().range();
    const n = axis === 'r' ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
    try { O.insertLines(sh(), axis, axis === 'r' ? rg.r1 : rg.c1, Math.min(n, axis === 'r' ? 100000 : 16384)); } catch (e) { A.error(e); }
    G().paint(); G().syncObjects(true);
  };
  A.deleteLines = function (axis) {
    if (G().guardProtect(axis === 'r' ? 'deleteRows' : 'deleteColumns')) return;
    const list = axis === 'r' ? A.selRows() : A.selCols();
    if (!list.length) return;
    try {
      /* delete from the bottom so indexes stay valid */
      const runs = [];
      list.sort((a, b) => a - b);
      for (const i of list) { const last = runs[runs.length - 1]; if (last && last[1] === i - 1) last[1] = i; else runs.push([i, i]); }
      O.tx(wb(), axis === 'r' ? '删除行' : '删除列', () => { for (let k = runs.length - 1; k >= 0; k--) O.insertLines(sh(), axis, runs[k][0], -(runs[k][1] - runs[k][0] + 1)); });
    } catch (e) { A.error(e); }
    G().paint(); G().syncObjects(true);
  };
  A.insertSelection = function () {
    const rg = G().range();
    if (G().isWholeRows(rg)) { A.insertLines('r'); return; }
    if (G().isWholeCols(rg)) { A.insertLines('c'); return; }
    L.dlg.insertCells();
  };
  A.deleteSelection = function () {
    const rg = G().range();
    if (G().isWholeRows(rg)) { A.deleteLines('r'); return; }
    if (G().isWholeCols(rg)) { A.deleteLines('c'); return; }
    L.dlg.deleteCells();
  };

  /* formatting helpers */
  A.applyFill = function (hex) {
    if (G().guardProtect('formatCells')) return;
    O.format(sh(), G().ranges(), { fill: hex ? { pattern: 'solid', fg: M.rgb(hex), bg: { indexed: 64 } } : null }, 'Fill Color');
    G().paint();
  };
  A.applyFontColor = function (hex) {
    if (G().guardProtect('formatCells')) return;
    O.format(sh(), G().ranges(), { font: { color: hex ? M.rgb(hex) : undefined } }, '字体颜色');
    G().paint();
  };
  A.applyBorderPreset = function (k) {
    if (G().guardProtect('formatCells')) return;
    const thin = { style: 'thin' }, thick = { style: 'thick' }, dbl = { style: 'double' }, med = { style: 'medium' };
    const spec = {
      none: { none: true }, bottom: { bottom: thin }, left: { left: thin }, right: { right: thin }, top: { top: thin }, bottomDouble: { bottom: dbl }, thickBottom: { bottom: thick },
      topBottom: { top: thin, bottom: thin }, topThickBottom: { top: thin, bottom: thick }, topDoubleBottom: { top: thin, bottom: dbl }, all: { all: true, line: thin }, outside: { outline: true, line: thin }, thickBox: { outline: true, line: thick }, medBox: { outline: true, line: med },
    }[k];
    if (!spec) return;
    O.borders(sh(), G().ranges(), spec);
    G().paint();
  };
  A.confirmMerge = function (fn) {
    const shx = sh();
    let multi = 0;
    for (const rg of G().ranges()) shx.each(rg.r1, rg.c1, Math.min(rg.r2, rg.r1 + 2000), Math.min(rg.c2, rg.c1 + 200), (cl) => { if (cl.v != null || cl.f != null) multi++; });
    if (multi > 1) {
      ui.msg('选区包含多个数据值。合并为一个单元格时，只会保留左上角的数据。', { icon: 'warn', buttons: ['确定', '取消'] }).then((i) => { if (i === 0) { fn(); G().paint(); ui.refresh(); } });
      return;
    }
    fn();
  };
  /* Format Painter */
  /* Format Painter: carries the selected range's formats; the grid calls release() when the cells to
     paint are dragged out */
  A.painter = ui.brush({
    el: () => L.$('#gridwrap'),
    pick: () => { const s = G().sel(); return { sh: sh(), range: Object.assign({}, s.ranges[s.active]) }; },
    apply: (p) => {
      const s = G().sel();
      try { O.paintFormats(sh(), p.range, p.sh, s.ranges[s.active]); } catch (e) { A.error(e); }
      G().paint();
      return true;
    },
    status: () => updateStatus(),
  });

  /* AutoSum */
  A.autoSum = function (fn) {
    const shx = sh(), s = G().sel();
    const rg = s.ranges[s.active];
    const isNum = (r, c) => { const cl = shx.get(r, c); if (!cl) return false; const v = cl.dirty ? C.cellValue(shx, r, c) : cl.v; return typeof v === 'number'; };
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (single) {
      /* look above, then left, for a run of numbers */
      const r = s.r, c = s.c;
      let r1 = r - 1;
      while (r1 >= 0 && !isNum(r1, c) && r1 >= r - 1 && shx.get(r1, c) == null) r1--;
      let a = null;
      if (r1 >= 0 && isNum(r1, c)) { let t = r1; while (t - 1 >= 0 && (isNum(t - 1, c) || (shx.get(t - 1, c) == null && false))) t--; a = { r1: t, c1: c, r2: r1, c2: c }; }
      else if (c > 0 && isNum(r, c - 1)) { let t = c - 1; while (t - 1 >= 0 && isNum(r, t - 1)) t--; a = { r1: r, c1: t, r2: r, c2: c - 1 }; }
      const ref = a ? F.rangeName(a) : '';
      const sub = fn === 'SUM' ? 'SUM' : fn;
      L.editor.begin('enter', '=' + sub + '(' + ref + ')');
      if (a) { const t = L.editor.text(); const start = t.indexOf('(') + 1; try { G().input.setSelectionRange(start, start + ref.length); } catch (e) { /* ignore */ } }
      else { const t = L.editor.text(); try { G().input.setSelectionRange(t.length - 1, t.length - 1); } catch (e) { /* ignore */ } }
      L.editor.mode = 'point';
      L.editor.highlight();
      return;
    }
    /* a range: totals below each column (or right of each row) */
    const rr = O.clip(shx, rg);
    O.tx(wb(), L.N('自动求和'), () => {
      const lastRowEmpty = (() => { for (let c = rr.c1; c <= rr.c2; c++) { const cl = shx.get(rr.r2, c); if (cl && cl.v != null) return false; } return true; })();
      const lastColEmpty = (() => { for (let r = rr.r1; r <= rr.r2; r++) { const cl = shx.get(r, rr.c2); if (cl && cl.v != null) return false; } return true; })();
      if (rr.r1 !== rr.r2 && (lastRowEmpty || !lastColEmpty)) {
        const tr = lastRowEmpty ? rr.r2 : rr.r2 + 1, end = lastRowEmpty ? rr.r2 - 1 : rr.r2;
        for (let c = rr.c1; c <= rr.c2; c++) { let any = false; for (let r = rr.r1; r <= end; r++) if (isNum(r, c)) any = true; if (any) O.enter(shx, tr, c, '=' + fn + '(' + F.rangeName({ r1: rr.r1, c1: c, r2: end, c2: c }) + ')'); }
      } else {
        const tc = lastColEmpty ? rr.c2 : rr.c2 + 1, end = lastColEmpty ? rr.c2 - 1 : rr.c2;
        for (let r = rr.r1; r <= rr.r2; r++) { let any = false; for (let c = rr.c1; c <= end; c++) if (isNum(r, c)) any = true; if (any) O.enter(shx, r, tc, '=' + fn + '(' + F.rangeName({ r1: r, c1: rr.c1, r2: r, c2: end }) + ')'); }
      }
    });
    G().paint();
  };
  /* sort buttons */
  A.quickSort = function (desc) {
    const shx = sh(), s = G().sel();
    let rg = s.ranges[s.active];
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    const t = A.tableAt();
    if (single) rg = t ? { r1: t.ref.r1, c1: t.ref.c1, r2: t.ref.r2 - (t.totals ? 1 : 0), c2: t.ref.c2 } : O.currentRegion(shx, s.r, s.c);
    rg = O.clip(shx, rg);
    if (rg.r1 === rg.r2) return;
    const header = t ? t.header !== false : single ? O.guessHeader(shx, rg) : false;
    const col = Math.max(rg.c1, Math.min(rg.c2, s.c)) - rg.c1;
    try { O.sort(shx, rg, [{ index: col, desc }], { header }); } catch (e) { A.error(e); }
    G().paint();
  };
  A.tableAt = function () { const s = G().sel(); return wb().tableAt(sh(), s.r, s.c); };
  A.toggleTotalRow = function () {
    const t = A.tableAt(); if (!t) return;
    const shx = sh();
    O.structural(shx, '汇总行', () => {
      if (t.totals) { t.totals = false; t.ref = Object.assign({}, t.ref, { r2: t.ref.r2 - 1 }); for (let c = t.ref.c1; c <= t.ref.c2; c++) { const row = shx.rows[t.ref.r2 + 1]; if (row) delete row.cells[c]; } }
      else {
        t.totals = true; t.ref = Object.assign({}, t.ref, { r2: t.ref.r2 + 1 });
        const r = t.ref.r2;
        t.columns.forEach((col, i) => {
          const c = t.ref.c1 + i;
          const row = shx.rowObj(r);
          if (i === 0) { row.cells[c] = { v: 'Total' }; col.totalsLabel = 'Total'; }
          else if (i === t.columns.length - 1) { col.totalsFn = 'sum'; row.cells[c] = { f: 'SUBTOTAL(109,' + t.name + '[' + col.name.replace(/([\[\]#'])/g, "'$1") + '])', v: null, dirty: true }; }
        });
        if (r > shx.maxR) shx.maxR = r;
      }
    });
    G().paint();
  };
  A.convertTableToRange = async function () {
    const t = A.tableAt(); if (!t) return;
    const r = await ui.msg('要将列表转换为普通区域吗？', { icon: 'question', buttons: ['是(&Y)', '否(&N)'] });
    if (r !== 0) return;
    const w = wb();
    O.tx(w, '转换为区域', () => {
      const old = w.tables.slice();
      /* structured references become ordinary ones */
      for (const s of w.sheets) s.rows.forEach((row, rr) => { if (!row) return; row.cells.forEach((cell, cc) => {
        if (!cell || cell.f == null || cell.f.indexOf('[') < 0) return;
        const ast = C.astOf(cell); if (!ast) return;
        let ch = false;
        const out = F.map(ast, (n) => { if (n.t !== 'struct') return null; const tb = n.table ? w.findTable(n.table) : w.tableAt(s, rr, cc); if (tb !== t) return null; const g = C.structRange(t, n.spec, rr); if (!g || M.isErr(g)) return null; ch = true; return g.r1 === g.r2 && g.c1 === g.c2 ? { t: 'ref', r: g.r1, c: g.c1, ra: false, ca: true, sheet: s !== t.sheet ? t.sheet.name : undefined } : { t: 'area', r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, ra1: true, ca1: true, ra2: true, ca2: true, sheet: s !== t.sheet ? t.sheet.name : undefined }; });
        if (ch) { w.undo.note(s, rr, cc); cell.f = F.toText(out, { store: true }); cell.ast = out; cell._fsrc = cell.f; }
      }); });
      w.tables = w.tables.filter((x) => x !== t);
      const now = w.tables.slice();
      w.undo.op(() => { w.tables = old; }, () => { w.tables = now; });
      O.markStructural();
    });
    G().paint(); ui.refresh();
  };
  /* outline / grouping */
  A.group = function (d) {
    const rg = G().range();
    const axis = G().isWholeCols(rg) ? 'c' : G().isWholeRows(rg) ? 'r' : null;
    const go = (ax) => {
      const shx = sh();
      const list = ax === 'r' ? A.selRows() : A.selCols();
      O.structural(shx, d > 0 ? L.N('组合') : L.N('取消组合'), () => {
        for (const i of list) {
          const o = ax === 'r' ? shx.rowObj(i) : shx.colObj(i);
          const lvl = Math.max(0, Math.min(7, (o.level || 0) + d));
          if (lvl) o.level = lvl; else delete o.level;
          if (!lvl && o.hidden && d < 0) delete o.hidden;
        }
        const maxLvl = (arr) => arr.reduce((m, o) => Math.max(m, (o && o.level) || 0), 0);
        if (ax === 'r') shx.outline.levelRow = maxLvl(shx.rows); else shx.outline.levelCol = maxLvl(shx.cols);
        LY.invalidate(shx);
      });
      L.outline && L.outline.refresh();
      G().paint();
    };
    if (axis) { go(axis); return; }
    L.dlg.groupRowsCols(d).then((ax) => { if (ax) go(ax); });
  };
  A.outlineDetail = function (show) {
    const shx = sh();
    const rows = A.selRows();
    O.structural(shx, show ? '显示明细数据' : '隐藏明细数据', () => {
      for (const r of rows) { const o = shx.rows[r]; if (o && o.level) { if (show) delete o.hidden; else o.hidden = true; } }
      LY.invalidate(shx);
    });
    G().paint();
  };
  A.autoOutline = function () {
    /* group rows above formula rows that SUM the cells above them */
    const shx = sh();
    O.structural(shx, '自动建立分级显示', () => {
      shx.rows.forEach((row, r) => { if (!row) return; for (const k in row.cells) { const cl = row.cells[k]; if (!cl || cl.f == null) continue; const m = /^(?:SUM|SUBTOTAL\(\d+,)\(?([A-Z]+)(\d+):([A-Z]+)(\d+)\)/i.exec(cl.f); if (m && +m[4] === r) { for (let x = +m[2] - 1; x < r; x++) { const o = shx.rowObj(x); o.level = Math.max(o.level || 0, 1); } } } });
      shx.outline.levelRow = shx.rows.reduce((m, o) => Math.max(m, (o && o.level) || 0), 0);
      LY.invalidate(shx);
    });
    G().paint();
  };
  A.clearOutline = function () {
    const shx = sh();
    O.structural(shx, '清除分级显示', () => {
      shx.rows.forEach((o) => { if (o) { delete o.level; delete o.collapsed; } });
      shx.cols.forEach((o) => { if (o) { delete o.level; delete o.collapsed; } });
      shx.outline.levelRow = 0; shx.outline.levelCol = 0;
      LY.invalidate(shx);
    });
    G().paint();
  };
  /* freeze */
  A.toggleFreeze = function () {
    const shx = sh();
    if (shx.view.freeze) { O.tx(wb(), '取消冻结窗格', () => O.setFreeze(shx, { f: null, t: { r: 0, c: 0 } })); G().vs().scrollR = Math.max(0, G().vs().scrollR); G().paint(); ui.refresh(); return; }
    const s = G().sel();
    const vs = G().vs();
    let r = s.r - vs.scrollR, c = s.c - vs.scrollC;
    const rg = s.ranges[s.active];
    if (G().isWholeRows(rg)) c = 0;
    if (G().isWholeCols(rg)) r = 0;
    if (r <= 0 && c <= 0) {
      /* freezing at the top-left cell splits the window in the middle */
      const fr = G().frame();
      const main = fr.panes[fr.panes.length - 1];
      r = Math.max(1, Math.floor((main.rEnd - main.r0) / 2)); c = Math.max(1, Math.floor((main.cEnd - main.c0) / 2));
    }
    r = Math.max(0, r); c = Math.max(0, c);
    O.tx(wb(), '冻结窗格', () => O.setFreeze(shx, { f: { r, c, top: { r: vs.scrollR + r, c: vs.scrollC + c } }, t: { r: vs.scrollR, c: vs.scrollC } }));
    vs.scrollR = vs.scrollR + r; vs.scrollC = vs.scrollC + c;
    G().paint(); ui.refresh();
  };
  A.setPageBreakPreview = function (on) {
    const shx = sh();
    shx.view.pageBreakPreview = !!on;
    if (on) { shx.view.zoom = Math.min(shx.view.zoom || 100, 60); L.print.computeBreaks(shx); ui.msg('欢迎使用分页预览\n\n您可以用鼠标单击并拖动分页符，调整分页位置。', { icon: 'info' }); }
    else { shx.view.zoom = 100; }
    G().paint(); ui.refresh();
  };
  A.zoomToSelection = function () {
    const rg = G().range();
    const g = LY.geo(sh());
    const fr = G().frame();
    const main = fr.panes[fr.panes.length - 1];
    const w = g.cols.pos(Math.min(rg.c2 + 1, MAXC)) - g.cols.pos(rg.c1), hh = g.rows.pos(Math.min(rg.r2 + 1, MAXR)) - g.rows.pos(rg.r1);
    const z = Math.max(10, Math.min(400, Math.floor(Math.min((main.w + main.x - fr.hdrW) / w, (main.h + main.y - fr.hdrH) / hh) * 100)));
    G().setZoom(z);
    G().vs().scrollR = rg.r1; G().vs().scrollC = rg.c1;
    G().paint();
  };
  A.toggleTask = function (on) { const tp = L.$('#taskpane'); if (on == null) on = tp.hidden; if (on && L.panes) { tp.hidden = false; if (!L.panes.task.current) L.panes.task.show('getting-started'); } else tp.hidden = !on; A.opts.taskOpen = !tp.hidden; A.saveOpts(); G().paint(); };
  A.addOfficeClip = (clip) => { if (L.panes && L.panes.addClip) L.panes.addClip(clip); };
  A.toggleTaskPane = function () { const tp = L.$('#taskpane'); if (tp.hidden) L.panes.task.show(L.panes.task.current || 'getting-started'); else tp.hidden = true; A.opts.taskOpen = !tp.hidden; A.saveOpts(); G().paint(); };
  A.toggleFullScreen = function () { L.$('#app').classList.toggle('fullscreen'); G().paint(); ui.refresh(); };
  A.calculateNow = function () { const w = wb(); O.recalc(w, null, true); w.needsCalc = false; G().paint(); updateStatus(); };

  /* print area & page breaks */
  A.printAreaOf = (s) => wb().names.find((n) => n.name.toLowerCase() === '_xlnm.print_area' && n.scope === wb().sheets.indexOf(s)) || null;
  A.setPrintArea = function (on) {
    const w = wb(), s = sh(), idx = w.sheets.indexOf(s);
    const names = w.names.filter((n) => !(n.name.toLowerCase() === '_xlnm.print_area' && n.scope === idx));
    if (on) names.push({ name: '_xlnm.Print_Area', ref: G().ranges().map((rg) => F.quoteSheet(s.name) + '!' + F.absRangeName(rg)).join(','), scope: idx });
    O.tx(w, on ? '设置打印区域' : '清除打印区域', () => O.setNames(w, names));
    L.print.computeBreaks(s);
    G().paint();
  };
  A.pageBreakAtSel = function () { const s = G().sel(); const p = sh().print; return (p.rowBreaks || []).includes(s.r) || (p.colBreaks || []).includes(s.c); };
  A.togglePageBreak = function () {
    const shx = sh(), s = G().sel();
    const p = JSON.parse(JSON.stringify(shx.print));
    const rg = s.ranges[s.active];
    const has = A.pageBreakAtSel();
    if (has) { p.rowBreaks = (p.rowBreaks || []).filter((b) => b !== s.r); p.colBreaks = (p.colBreaks || []).filter((b) => b !== s.c); }
    else { if (!G().isWholeCols(rg) && s.r > 0) p.rowBreaks = (p.rowBreaks || []).concat([s.r]).sort((a, b) => a - b); if (!G().isWholeRows(rg) && s.c > 0) p.colBreaks = (p.colBreaks || []).concat([s.c]).sort((a, b) => a - b); }
    O.tx(wb(), has ? '删除分页符' : '插入分页符', () => O.setPrint(shx, p));
    L.print.computeBreaks(shx);
    G().paint(); ui.refresh();
  };
  A.resetPageBreaks = function () { const shx = sh(); const p = JSON.parse(JSON.stringify(shx.print)); p.rowBreaks = []; p.colBreaks = []; O.tx(wb(), '重置分页符', () => O.setPrint(shx, p)); L.print.computeBreaks(shx); G().paint(); };

  /* comments */
  A.editComment = function (r, c) {
    const s = G().sel();
    if (r == null) { r = s.r; c = s.c; }
    const shx = sh();
    if (G().guardProtect('objects')) return;
    const cm = shx.comments.get(M.key(r, c));
    L.dlg.comment(r, c, cm);
  };

  /* pictures */
  A.insertPicture = async function () {
    const files = await L.pickFiles('image/*,.emf,.wmf', true);
    for (const f of files) await A.insertPictureFile(f);
  };
  A.insertPictureFile = async function (file) {
    if (G().guardProtect('objects')) return;
    const w = wb(), shx = sh(), s = G().sel();
    const bytes = new Uint8Array(await L.readAsArrayBuffer(file));
    let ext = (file.name.split('.').pop() || 'png').toLowerCase();
    let data = bytes, type = file.type || L.extToMime(ext);
    /* SVG becomes a PNG, which every Excel version can show */
    if (ext === 'svg' || /svg/.test(type)) {
      const png = await new Promise((res) => {
        const url = URL.createObjectURL(new Blob([bytes], { type: 'image/svg+xml' }));
        const im = new Image();
        im.onload = () => { const k = 2, cv = document.createElement('canvas'); cv.width = Math.max(1, (im.naturalWidth || 300) * k); cv.height = Math.max(1, (im.naturalHeight || 150) * k); cv.getContext('2d').drawImage(im, 0, 0, cv.width, cv.height); URL.revokeObjectURL(url); cv.toBlob((b) => (b ? b.arrayBuffer().then((ab) => res(new Uint8Array(ab))) : res(null)), 'image/png'); };
        im.onerror = () => { URL.revokeObjectURL(url); res(null); };
        im.src = url;
      });
      if (!png) { ui.msg('Ledger 无法读取此图片。', { icon: 'warn' }); return; }
      data = png; ext = 'png'; type = 'image/png';
    }
    if (!w.media) w.media = new Map();
    const id = 'img' + (w.media.size + 1) + '_' + Date.now().toString(36);
    w.media.set(id, { bytes: data, ext, type, name: file.name.replace(/\.svg$/i, '.png') });
    const dims = await new Promise((res) => { const url = URL.createObjectURL(new Blob([bytes], { type: file.type || L.extToMime(ext) })); const im = new Image(); im.onload = () => { res({ w: im.naturalWidth, h: im.naturalHeight }); URL.revokeObjectURL(url); }; im.onerror = () => res({ w: 200, h: 150 }); im.src = url; });
    const scale = Math.min(1, 480 / dims.w, 360 / dims.h);
    const d = { kind: 'image', media: id, name: 'Picture ' + (shx.drawings.length + 1), id: shx.drawings.length + 2, anchor: { type: 'one', from: { r: s.r, c: s.c, rOff: 0, cOff: 0 }, w: dims.w * scale * 0.75, h: dims.h * scale * 0.75 } };
    O.tx(w, '插入图片', () => O.setDrawings(shx, shx.drawings.concat([d])));
    G().selectObject(d);
  };
  A.deleteObject = function (d) {
    const shx = sh();
    O.tx(wb(), L.N('删除'), () => O.setDrawings(shx, shx.drawings.filter((x) => x !== d)));
    G().deselectObject(); G().syncObjects(true);
  };
  A.sheetBackground = async function () {
    const shx = sh();
    if (shx.background) { O.tx(wb(), '删除背景', () => O.sheetProp(wb(), shx, 'background', null)); G().paint(); return; }
    const files = await L.pickFiles('image/*');
    if (!files.length) return;
    const f = files[0];
    const bytes = new Uint8Array(await L.readAsArrayBuffer(f));
    const id = 'bg' + Date.now().toString(36);
    wb().media.set(id, { bytes, ext: (f.name.split('.').pop() || 'png').toLowerCase(), type: f.type, name: f.name });
    O.tx(wb(), L.N('背景'), () => O.sheetProp(wb(), shx, 'background', id));
    G().paint();
  };
  A.pasteAsHyperlink = function () {
    const c = L.clip.get();
    if (!c) return;
    const s = G().sel();
    const loc = F.quoteSheet(c.sh.name) + '!' + F.rangeName(c.range);
    O.tx(wb(), '粘贴为超链接', () => {
      O.setLinks(sh(), sh().links.filter((l) => !M.rangeContains(l.ref, s.r, s.c)).concat([{ ref: { r1: s.r, c1: s.c, r2: s.r, c2: s.c }, location: loc, display: String(c.snap.rows[0][0] && c.snap.rows[0][0].v || loc) }]));
      if (sh().get(s.r, s.c) == null) O.setValue(sh(), s.r, s.c, String(c.snap.rows[0][0] && c.snap.rows[0][0].v || loc));
    });
  };

  /* formula auditing arrows */
  A.trace = function (kind) {
    const st = G().state;
    if (!kind) { st.arrows = null; G().paint(); return; }
    const shx = sh(), s = G().sel();
    const cell = shx.get(s.r, s.c);
    st.arrows = st.arrows || [];
    if (kind === 'prec' || kind === 'err') {
      if (!cell || cell.f == null) { ui.msg(kind === 'err' ? '活动单元格不包含错误值。' : '活动单元格中的公式不包含有效引用。', { icon: 'warn' }); return; }
      const p = C.precedents(wb(), shx, C.astOf(cell), s.r, s.c);
      for (const a of p.refs) st.arrows.push({ from: a, to: { sheet: shx, r1: s.r, c1: s.c, r2: s.r, c2: s.c }, err: kind === 'err' });
    } else {
      const deps = [];
      wb().graph.dependents(shx, s.r, s.c, deps);
      if (!deps.length) { ui.msg('Ledger 找不到任何从属单元格。', { icon: 'warn' }); return; }
      for (const n of deps) st.arrows.push({ from: { sheet: shx, r1: s.r, c1: s.c, r2: s.r, c2: s.c }, to: { sheet: n.sh, r1: n.r, c1: n.c, r2: n.r, c2: n.c } });
    }
    G().paint();
  };
  /** Circle Invalid Data: red ovals around cells whose values break their validation rule (at most 255, as Excel) */
  A.circleInvalid = function (on) {
    const st = G().state;
    if (!on) { st.circles = null; G().paint(); return; }
    const shx = sh();
    const list = [];
    let more = false;
    for (const dv of shx.dv) for (const rg of dv.ranges) {
      const r2 = Math.min(rg.r2, shx.maxR), c2 = Math.min(rg.c2, shx.maxC);
      for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) {
        const v = C.cellValue(shx, r, c);
        if (v == null || v === '') continue;
        if (L.cf.dvCheck(shx, r, c, v)) { if (list.length < 255) list.push({ r, c }); else more = true; }
      }
    }
    st.circles = { sheet: shx, list };
    G().paint();
    if (more) ui.msg('此工作表包含超过 255 个无效单元格。仅标记其中的 255 个。', { icon: 'warn' });
    else if (!list.length) ui.toast('此工作表中没有无效数据。');
  };
  L.bus.on('grid-paint', () => drawArrows());
  function drawArrows() {
    const layer = L.$('#glayer');
    let svg = layer.querySelector('svg.trace-svg');
    const st = G().state;
    /* a corrected cell loses its circle */
    if (st.circles) st.circles.list = st.circles.list.filter((p) => { const v = C.cellValue(st.circles.sheet, p.r, p.c); return v != null && v !== '' && L.cf.dvCheck(st.circles.sheet, p.r, p.c, v); });
    const circles = st.circles && st.circles.sheet === sh() ? st.circles.list : [];
    if ((!st.arrows || !st.arrows.length) && !circles.length) { if (svg) svg.remove(); return; }
    if (!svg) { svg = document.createElementNS(L.svgNS, 'svg'); svg.setAttribute('class', 'trace-svg'); layer.appendChild(svg); }
    const W = L.$('#gridwrap').clientWidth, H = L.$('#gridwrap').clientHeight;
    svg.setAttribute('width', W); svg.setAttribute('height', H);
    const shx = sh();
    let out = '<defs><marker id="tarr" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#0000FF"/></marker><marker id="tarre" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#FF0000"/></marker></defs>';
    for (const p of circles) {
      const q = G().cellRect(p.r, p.c);
      if (q.w <= 0 || q.h <= 0) continue;
      out += `<ellipse cx="${q.x + q.w / 2}" cy="${q.y + q.h / 2}" rx="${q.w / 2 + 4}" ry="${q.h / 2 + 3}" fill="none" stroke="#FF0000" stroke-width="1.5"/>`;
    }
    for (const a of st.arrows || []) {
      if (a.to.sheet !== shx) continue;
      const q2 = G().cellRect(a.to.r1, a.to.c1);
      const x2 = q2.x + q2.w / 2, y2 = q2.y + q2.h / 2;
      const col = a.err ? '#FF0000' : '#0000FF';
      if (a.from.sheet !== shx) { out += `<path d="M${x2 - 40} ${y2 - 30} L${x2} ${y2}" stroke="#000" stroke-dasharray="4 2" fill="none" marker-end="url(#tarr)"/><rect x="${x2 - 52}" y="${y2 - 40}" width="14" height="11" fill="#fff" stroke="#000"/>`; continue; }
      const q1 = G().cellRect(a.from.r1, a.from.c1);
      const multi = a.from.r1 !== a.from.r2 || a.from.c1 !== a.from.c2;
      if (multi) { const q3 = G().cellRect(a.from.r2, a.from.c2); out += `<rect x="${q1.x}" y="${q1.y}" width="${q3.x + q3.w - q1.x}" height="${q3.y + q3.h - q1.y}" fill="none" stroke="${col}" stroke-width="1.5"/>`; }
      const x1 = q1.x + Math.min(q1.w / 2, 12), y1 = q1.y + q1.h / 2;
      out += `<circle cx="${x1}" cy="${y1}" r="3" fill="${col}"/><path d="M${x1} ${y1} L${x2} ${y2}" stroke="${col}" stroke-width="1.5" fill="none" marker-end="url(#${a.err ? 'tarre' : 'tarr'})"/>`;
    }
    svg.innerHTML = out;
  }

  /* ------------------------------------------------------------ drag & drop of files */
  function setupDrop() {
    let hint = null, depth = 0;
    const app = L.$('#app');
    app.addEventListener('dragenter', (e) => { if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return; depth++; if (!hint) { hint = h('div', { class: 'drop-hint', text: '将工作簿拖到此处可打开，将图片拖到此处可插入' }); document.body.appendChild(hint); } });
    app.addEventListener('dragleave', () => { depth--; if (depth <= 0 && hint) { hint.remove(); hint = null; depth = 0; } });
    app.addEventListener('dragover', (e) => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) e.preventDefault(); });
    app.addEventListener('drop', (e) => {
      if (hint) { hint.remove(); hint = null; } depth = 0;
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
      if (!files.length) return;
      e.preventDefault();
      const books = files.filter((f) => /\.(xlsx|xlsm|xltx|xltm|csv|txt|tsv|prn|xml|xls)$/i.test(f.name));
      if (books.length) A.openFiles(books);
      const imgs = files.filter((f) => /^image\//.test(f.type));
      for (const f of imgs) A.insertPictureFile(f);
    });
  }

  /* ------------------------------------------------------------ sample workbook */
  A.sampleWorkbook = function () {
    const w = A.blankWorkbook();
    const s = w.sheets[0];
    const put = (a, v) => { const p = F.parseCell(a); const cell = s.cell(p.r, p.c); if (typeof v === 'string' && v[0] === '=') { cell.f = v.slice(1); cell.dirty = true; } else cell.v = v; };
    const st = (a, d) => { const p = F.parseCell(a); const cell = s.cell(p.r, p.c); cell.s = w.styles.derive(cell.s || 0, d); };
    put('A1', 'Quarterly Sales'); st('A1', { font: { b: true, sz: 14, color: M.rgb('#1F3F7F') } });
    const heads = ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'];
    heads.forEach((t, i) => { put(F.cellName(2, i), t); st(F.cellName(2, i), { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#C0D9F0') }, border: { b: { style: 'thin' } }, align: i ? { h: 'center' } : null }); });
    const data = [['North', 15200, 16850, 14900, 19750], ['South', 11800, 12100, 13400, 15020], ['East', 9650, 10900, 12250, 13100], ['West', 13300, 12950, 14800, 17600]];
    data.forEach((row, i) => { row.forEach((v, j) => { put(F.cellName(3 + i, j), v); if (j) st(F.cellName(3 + i, j), { nf: '#,##0' }); }); put(F.cellName(3 + i, 5), `=SUM(B${4 + i}:E${4 + i})`); st(F.cellName(3 + i, 5), { nf: '#,##0', font: { b: true } }); });
    put('A8', 'Total'); st('A8', { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } } });
    for (let j = 1; j <= 5; j++) { const col = F.colName(j); put(col + '8', `=SUM(${col}4:${col}7)`); st(col + '8', { nf: '#,##0', font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } } }); }
    put('A10', 'Share of year'); st('A10', { font: { i: true } });
    for (let j = 1; j <= 4; j++) { const col = F.colName(j); put(col + '10', `=${col}8/$F$8`); st(col + '10', { nf: '0.0%' }); }
    put('A12', 'Best quarter'); put('B12', '=INDEX(B3:E3,MATCH(MAX(B8:E8),B8:E8,0))');
    put('A13', 'Growth Q1→Q4'); put('B13', '=E8/B8-1'); st('B13', { nf: '0.0%' });
    put('A15', 'Type into any cell, use the formula bar, or open a workbook with File ▸ Open.'); st('A15', { font: { color: M.rgb('#666666'), i: true } });
    s.cols[0] = { w: 16, custom: true };
    for (let j = 1; j <= 5; j++) s.cols[j] = { w: 11, custom: true };
    s.view.sel = { r: 3, c: 1, ranges: [{ r1: 3, c1: 1, r2: 3, c2: 1 }] };
    return w;
  };

  /* ------------------------------------------------------------ init */
  A.init = function () {
    A.loadOpts();
    if (window.VO) VO.setIcon(L.icons.app(32));
    /* opened from the Start Center: a template, or blank (a new workbook, or to open a file into); else the sample */
    const launch = window.VO && VO.launch;
    const tpl = launch && launch.template && L.panes && L.panes.TEMPLATES[launch.template];
    const w = tpl ? L.panes.buildTemplate(launch.template) : launch ? A.blankWorkbook() : A.sampleWorkbook();
    untitled = 1;
    A.prepare(w);
    G().init(w);
    L.editor.init();
    buildMenus();
    buildShortcuts();
    buildToolbars();
    L.ribbonUI.init();
    buildStatus();
    buildTabNav();
    buildNameBox();
    if (L.panes) L.panes.task.mount(L.$('#taskpane'));
    /* the open task pane follows the selected object to its tool set (common/ribbon.js) */
    if (L.panes) L.ribbonUI.follow({ task: L.panes.task, sets: { chart: true, picture: true, drawing: true }, on: () => A.opts.paneFollows !== false });
    A.addBook(w, tpl ? tpl.name.replace(/\s+/g, '') + '1' : 'Book1', { type: 'xlsx', untitled: true });
    setupDrop();
    L.bus.on('selection', () => { updateStatus(); ui.refresh(); A.updateToolbars(); });
    L.bus.on('changed', () => { updateStatus(); ui.refresh(); renderTabsSoon(); });
    L.bus.on('mode', () => updateStatus());
    L.bus.on('zoom', () => ui.refresh());
    window.addEventListener('beforeunload', (e) => { if (A.books.some((b) => b.dirty)) { e.preventDefault(); e.returnValue = ''; } });
    const narrow = window.innerWidth < 820;
    if (A.opts.taskOpen && !narrow && L.panes) L.panes.task.show('getting-started');
    A.applyOpts();
    ui.refresh();
    G().focus();
    L.bus.emit('app-ready');
    if (window.VO) VO.attach('ledger', {
      async open(f, o) {
        await A.openFiles([f]);
        /* an unsaved version: the changes are not saved yet, and a never-saved workbook still needs Save As */
        const b = book();
        if (o?.draft && b) L.opc.recoverLosses(b.wb, o.lossState);
        if (o && o.draft && b) { b.dirty = true; b.name = o.saved ? o.name : o.name.replace(/\.(xlsx|xlsm|xltx|xltm)$/i, ''); b.untitled = !o.saved; b.saved = !!o.saved; updateTitle(); A.status(L.t('已恢复未保存的更改。请保存工作簿以保留这些更改。')); ui.refresh(); }
      },
      snapshot: (b) => L.xlsxWrite.write(b.wb, { type: L.opc.formats[b.type] ? b.type : b.wb.ooxmlFormat || 'xlsx', password: b.password }),
      docInfo(b) {
        const type = L.opc.formats[b.type] ? b.type : b.wb.ooxmlFormat || 'xlsx';
        const name = /\.\w+$/.test(b.name) ? b.name : b.name + '.' + type;
        return { name, draftName: name.replace(/\.(csv|txt|tsv|prn|xml|html?)$/i, '.' + type), lossState: L.opc.lossState(b.wb) };
      },
      isDirty: (b) => !!(b && b.dirty),
      current: () => book(),
      showRecovery() { A.opts.taskOpen = true; L.$('#taskpane').hidden = false; L.panes.task.show('recovery'); },
      thumb: gridThumb,
      message: (t) => ui.msg(t, { icon: 'warn' }),
      templates: () => Object.entries(L.panes.TEMPLATES).map(([id, t]) => ({ id, name: t.name })),
    });
  };
  /** Recent Files picture: the top left of the sheet as it is drawn, 4:3 */
  function gridThumb() {
    const c = L.$('canvas.gcv');
    if (!c || !c.width || !c.height) return null;
    const w = Math.min(600, c.width, Math.round(c.height * 4 / 3)), hh = Math.round(w * 3 / 4);   // about A1:G25 at 100 %
    const out = h('canvas', { width: w, height: hh });
    out.getContext('2d').drawImage(c, 0, 0, w, hh, 0, 0, w, hh);
    return out;
  }
  const renderTabsSoon = L.debounce(() => renderTabs(), 30);
  if (typeof document !== 'undefined') {
    /* Embedded copy (tools-box): no i18n runtime — the text is already Chinese (common/text.js) */
    const boot = () => A.init();
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
    else setTimeout(boot, 0);
  }
})(typeof window !== 'undefined' ? window : globalThis);
