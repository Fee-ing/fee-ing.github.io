/* Lectern — panes and alternate views. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = L.ui;
  const M = L.model;
  const E = () => L.ed;
  const P = (L.panes = {});

  /* ================= Slides / Outline pane ================= */
  const LP = (P.left = { tab: 'slides' });
  LP.mount = function (el) {
    LP.el = el;
    LP.tabs = h('div', { class: 'lp-tabs', role: 'tablist' });
    LP.tabOutline = h('div', { class: 'lp-tab', role: 'tab', tabindex: '0', text: '大纲' });
    LP.tabSlides = h('div', { class: 'lp-tab', role: 'tab', tabindex: '0', text: '幻灯片' });
    const close = h('button', { class: 'lp-close', type: 'button', 'aria-label': '关闭窗格', html: L.icons.get('close', 12) });
    close.addEventListener('click', () => L.app.toggleLeft(false));
    LP.tabs.append(LP.tabOutline, LP.tabSlides, close);
    LP.body = h('div', { class: 'lp-body', tabindex: '-1' });
    el.append(LP.tabs, LP.body);
    LP.tabOutline.addEventListener('click', () => LP.setTab('outline'));
    LP.tabSlides.addEventListener('click', () => LP.setTab('slides'));
    LP.body.addEventListener('focus', () => { if (L.app) L.app.focusArea = 'slides'; }, true);
    LP.body.addEventListener('blur', () => { if (L.app) L.app.focusArea = null; }, true);
    LP.setTab('slides');
  };
  LP.setTab = function (t) {
    if (LP.tab === 'outline' && t !== 'outline') LP.commitOutline();
    LP.tab = t;
    LP.tabOutline.classList.toggle('on', t === 'outline');
    LP.tabSlides.classList.toggle('on', t === 'slides');
    LP.render();
    L.ui.refresh();
  };
  LP.render = function () {
    if (!LP.body) return;
    if (E().view === 'master') { renderMasterThumbs(); return; }
    if (LP.tab === 'outline') renderOutline();
    else renderThumbs();
  };
  /* --- thumbnails --- */
  function thumbWidth() { return Math.max(60, (LP.body.clientWidth || 190) - 48); }
  function renderThumbs() {
    const pres = L.pres;
    const box = h('div', { class: 'thumbs', role: 'listbox', 'aria-label': '幻灯片' });
    const w = thumbWidth();
    pres.slides.forEach((s, i) => box.appendChild(thumbItem(s, i, w)));
    keepFocus(() => L.clear(LP.body).appendChild(box));
    LP.list = box;
    scrollCur();
  }
  /** the pane drawn again (a slide deleted, moved, added) keeps the keyboard when it had it: the focused thumbnail
      goes with the old list, and the next Delete or arrow must still work on slides (P4) */
  function keepFocus(fn) {
    const had = LP.body.contains(document.activeElement) || (L.app && L.app.focusArea === 'slides');
    fn();
    if (had && !LP.body.contains(document.activeElement)) { try { LP.body.focus({ preventScroll: true }); } catch (e) { /* ignore */ } L.app.focusArea = 'slides'; }
  }
  function thumbItem(s, i, w) {
    const sel = L.app.slideSel;
    const it = h('div', { class: 'th' + (i === E().idx ? ' cur' : '') + (sel.has(i) ? ' sel' : '') + (s.hidden ? ' hiddenslide' : ''), 'data-i': i, role: 'option', 'aria-selected': String(i === E().idx) });
    const meta = h('div', { class: 'th-meta' }, h('span', { class: 'th-num', text: String(i + (L.pres.firstNum || 1)) }));
    if ((s.trans && s.trans.type && s.trans.type !== 'none') || (s.anims && s.anims.length)) {
      const star = h('span', { class: 'th-star', 'data-tip': '播放动画', html: '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333"/></svg>' });
      star.addEventListener('click', (e) => { e.stopPropagation(); const fr = (e.currentTarget.closest('.th') || it).querySelector('.th-frame');   /* the thumbnail shown: refreshThumb moves the star into it */ L.show.previewTransition(fr, L.pres, s, L.pres.slides[i - 1]); });
      meta.appendChild(star);
    }
    const frame = h('button', { class: 'th-frame', type: 'button', 'aria-label': L.t('幻灯片 {0}', i + 1) });
    frame.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
    it.append(meta, frame);
    it.addEventListener('pointerdown', (e) => onThumbDown(e, i));
    it.addEventListener('contextmenu', (e) => { e.preventDefault(); if (!L.app.slideSel.has(i)) { L.app.slideSel = new Set([i]); E().goto(i); } L.app.contextThumb(e); });
    it.addEventListener('dblclick', () => { if (L.app.view !== 'normal') L.app.setView('normal'); });
    return it;
  }
  function scrollCur() {
    const cur = LP.body.querySelector('.th.cur');
    if (cur) cur.scrollIntoView({ block: 'nearest' });
  }
  LP.refreshThumb = function (i) {
    if (LP.tab !== 'slides' || !LP.list || E().view === 'master') return;
    const s = L.pres.slides[i];
    const old = LP.list.querySelector(`.th[data-i="${i}"]`);
    if (!s || !old) { renderThumbs(); return; }
    /* in place: the thumbnail's elements stay, so a press on it in the meantime still gets its click (P4) */
    const fresh = thumbItem(s, i, thumbWidth());
    old.className = fresh.className;
    old.querySelector('.th-meta').replaceWith(fresh.querySelector('.th-meta'));
    old.querySelector('.th-frame').replaceChildren(...fresh.querySelector('.th-frame').childNodes);
  };
  LP.markCurrent = function () {
    if (!LP.list) return;
    L.$$('.th', LP.list).forEach((el) => { const i = +el.dataset.i; el.classList.toggle('cur', i === E().idx); el.classList.toggle('sel', L.app.slideSel.has(i)); el.setAttribute('aria-selected', String(i === E().idx)); });
    scrollCur();
  };
  function onThumbDown(e, i) {
    if (e.button !== 0) return;
    /* the Slides pane takes the keyboard: text editing on the slide ends, even on the current slide's thumbnail (P3) */
    if (E().editing()) L.te.end();
    LP.body.focus({ preventScroll: true });
    L.app.focusArea = 'slides';
    const sel = L.app.slideSel;
    if (e.shiftKey) { const a = Math.min(E().idx, i), b = Math.max(E().idx, i); L.app.slideSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); LP.markCurrent(); return; }
    if (e.ctrlKey || e.metaKey) { if (sel.has(i) && sel.size > 1) sel.delete(i); else sel.add(i); LP.markCurrent(); return; }
    if (!sel.has(i)) L.app.slideSel = new Set([i]);
    if (E().idx !== i || L.app.view !== 'normal') E().goto(i);
    LP.markCurrent();
    /* drag to reorder (common/ui.js ui.drag: Esc, a second button or a release the page never saw cancel it) */
    let line = null, target = null;
    const mv = (ev) => {
      if (!line) { line = h('div', { class: 'drop-line' }); LP.body.appendChild(line); }
      const items = L.$$('.th', LP.list);
      const br = LP.body.getBoundingClientRect();
      /* let go outside the pane (over the slide), the slides stay where they are */
      if (ev.clientX < br.left || ev.clientX > br.right) { line.style.display = 'none'; target = null; return; }
      line.style.display = '';
      target = items.length;
      for (let k = 0; k < items.length; k++) { const r = items[k].getBoundingClientRect(); if (ev.clientY < r.top + r.height / 2) { target = k; break; } }
      const ref = items[target] || items[items.length - 1];
      const rr = ref.getBoundingClientRect();
      line.style.top = (target < items.length ? rr.top - 5 : rr.bottom + 4) - br.top + LP.body.scrollTop + 'px';
      if (ev.clientY < br.top + 20) LP.body.scrollTop -= 8; else if (ev.clientY > br.bottom - 20) LP.body.scrollTop += 8;
    };
    L.ui.drag(e, {
      threshold: 6,
      move: mv,
      up() { if (line) { line.remove(); if (target != null) L.app.moveSlides(Array.from(L.app.slideSel).sort((a, b) => a - b), target); } },
      cancel() { if (line) line.remove(); },
    });
  }
  function renderMasterThumbs() {
    const box = h('div', { class: 'thumbs' });
    const d = E().design();
    const kinds = [['slide', '幻灯片母版']].concat(d.titleDeco ? [['title', '标题母版']] : []);
    kinds.forEach(([k, label], i) => {
      E().masterKind = E().masterKind === 'title' && !d.titleDeco ? 'slide' : E().masterKind;
      const it = h('div', { class: 'th' + (E().masterKind === k ? ' cur' : '') });
      const frame = h('button', { class: 'th-frame', type: 'button', 'aria-label': label, 'data-tip': L.t(k === 'title' ? '{0} 标题母版' : '{0} 幻灯片母版', M.designName(d)) });
      const saveKind = E().masterKind;
      E().masterKind = k; E().buildMaster();
      frame.appendChild(L.render.thumb(L.pres, E().masterSlide, thumbWidth(), { prompts: false }));
      E().masterKind = saveKind; E().buildMaster();
      frame.addEventListener('click', () => { E().masterKind = k; E().buildMaster(); E().sel = []; E().render(); renderMasterThumbs(); });
      it.append(h('div', { class: 'th-meta' }, h('span', { class: 'th-num', text: String(i + 1) })), frame);
      box.appendChild(it);
    });
    L.clear(LP.body).appendChild(box);
    LP.list = null;
  }

  /* --- outline --- */
  let olDirty = false, olHistPushed = false;
  function renderOutline() {
    const ol = h('div', { class: 'outline' + (L.app.opts.outlinePlain ? ' plain' : ''), contenteditable: 'true', spellcheck: String(!!L.app.opts.spell), role: 'textbox', 'aria-multiline': 'true', 'aria-label': '大纲' });
    L.pres.slides.forEach((s, i) => {
      const title = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle'));
      const row = h('div', { class: 'ol-row title', 'data-slide': s.id });
      row.appendChild(document.createTextNode(title ? L.txt.plain(title.tx).replace(/\n/g, ' ') : ''));
      if (!row.textContent) row.appendChild(h('br'));
      ol.appendChild(row);
      const body = bodyShape(s);
      if (body) for (const p of body.tx.ps) {
        const t = L.txt.paraText(p).replace(/\n/g, ' ');
        const r = h('div', { class: 'ol-row body', 'data-lvl': p.lvl || 0 });
        r.appendChild(document.createTextNode(t));
        if (!t) r.appendChild(h('br'));
        ol.appendChild(r);
      }
      void i;
    });
    ol.addEventListener('keydown', onOutlineKey);
    ol.addEventListener('input', () => { if (!olHistPushed) { L.hist.push('编辑大纲'); olHistPushed = true; } olDirty = true; commitSoon(); });
    ol.addEventListener('paste', (e) => { e.preventDefault(); const t = (e.clipboardData && e.clipboardData.getData('text/plain')) || ''; document.execCommand('insertText', false, t.replace(/\r?\n/g, ' ')); });
    ol.addEventListener('focusout', () => LP.commitOutline());
    ol.addEventListener('keyup', syncOutlineSlide);
    ol.addEventListener('click', syncOutlineSlide);
    L.clear(LP.body).appendChild(ol);
    LP.outline = ol;
    styleOutline();
  }
  function styleOutline() {
    if (!LP.outline) return;
    let n = L.pres.firstNum || 1;
    for (const r of LP.outline.children) {
      if (!r.classList) continue;
      if (r.classList.contains('title')) { r.dataset.num = n++; r.style.cssText = 'font-weight:bold;margin-top:6px;padding-left:36px;position:relative;min-height:1.35em'; }
      else { const lv = +r.dataset.lvl || 0; r.style.cssText = `padding-left:${48 + lv * 18}px;position:relative;min-height:1.35em;font-weight:${L.app.opts.outlinePlain ? 'normal' : 'normal'}`; }
    }
    let style = document.getElementById('ol-style');
    if (!style) { style = h('style', { id: 'ol-style' }); document.head.appendChild(style); }
    style.textContent = '.outline .ol-row.title::before{content:attr(data-num);position:absolute;left:0;width:16px;text-align:right;font-weight:normal}.outline .ol-row.title::after{content:"";position:absolute;left:20px;top:3px;width:12px;height:9px;border:1px solid #555;background:linear-gradient(#cfdcf3 0 35%,#fff 35%)}.outline .ol-row.body::before{content:"•";position:absolute;margin-left:-12px}';
  }
  const commitSoon = L.debounce(() => LP.commitOutline(true), 700);
  function rowCaret() {
    const s = window.getSelection();
    if (!s.rangeCount) return null;
    let n = s.anchorNode;
    while (n && n.parentNode !== LP.outline) n = n.parentNode;
    return n;
  }
  function syncOutlineSlide() {
    const row = rowCaret();
    if (!row) return;
    let idx = -1;
    for (const r of LP.outline.children) { if (r.classList.contains('title')) idx++; if (r === row) break; }
    idx = Math.max(0, idx);
    if (idx !== E().idx && idx < L.pres.slides.length) { E().idx = idx; E().render(); L.bus.emit('slide-changed-silent'); L.app.updateStatus(); }
  }
  function onOutlineKey(e) {
    const row = rowCaret();
    if (!row) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      if (!olHistPushed) { L.hist.push('编辑大纲'); olHistPushed = true; }
      const s = window.getSelection();
      const r = s.getRangeAt(0);
      r.deleteContents();
      const after = document.createRange();
      after.setStart(r.endContainer, r.endOffset);
      after.setEndAfter(row.lastChild || row);
      const tail = after.extractContents();
      const nr = h('div', { class: row.className, 'data-lvl': row.dataset.lvl || '0' });
      nr.appendChild(tail);
      if (!nr.textContent) { L.clear(nr); nr.appendChild(h('br')); }
      if (!row.textContent) { L.clear(row); row.appendChild(h('br')); }
      if (e.ctrlKey && row.classList.contains('title')) { nr.className = 'ol-row body'; nr.dataset.lvl = '0'; }
      row.after(nr);
      const c = document.createRange(); c.setStart(nr, 0); c.collapse(true); s.removeAllRanges(); s.addRange(c);
      olDirty = true; styleOutline(); commitSoon();
      return;
    }
    if (e.key === 'Tab' || (e.altKey && e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft'))) {
      e.preventDefault();
      const demote = e.key === 'Tab' ? !e.shiftKey : e.key === 'ArrowRight';
      if (!olHistPushed) { L.hist.push('编辑大纲'); olHistPushed = true; }
      if (demote) { if (row.classList.contains('title')) { if (row.previousElementSibling) { row.className = 'ol-row body'; row.dataset.lvl = '0'; } } else row.dataset.lvl = Math.min(4, (+row.dataset.lvl || 0) + 1); }
      else if (row.classList.contains('body')) { const lv = +row.dataset.lvl || 0; if (lv === 0) row.className = 'ol-row title'; else row.dataset.lvl = lv - 1; }
      olDirty = true; styleOutline(); commitSoon();
    }
  }
  function bodyShape(s) { return s.shapes.find((x) => x.ph && x.type === 'text' && ['body', 'obj', 'subTitle'].includes(x.ph.type)); }
  LP.commitOutline = function (keepFocus) {
    if (!olDirty || !LP.outline) return;
    olDirty = false;
    const rows = Array.from(LP.outline.children).map((r) => ({ title: r.classList.contains('title'), lvl: +r.dataset.lvl || 0, text: r.textContent.replace(/ /g, ' ') }));
    const groups = [];
    for (const r of rows) {
      if (r.title || !groups.length) groups.push({ title: r.title ? r.text : '', body: r.title ? [] : [r] });
      else groups[groups.length - 1].body.push(r);
    }
    const pres = L.pres;
    groups.forEach((g, i) => {
      let s = pres.slides[i];
      if (!s) { s = M.newSlide(pres, 'text', (pres.slides[i - 1] || pres.slides[0] || {}).design); M.insertSlides(pres, pres.slides.length, [s]); }
      let t = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle'));
      if (!t && g.title) { M.applyLayout(pres, s, s.layout === 'blank' || s.layout === 'contentOnly' ? 'titleOnly' : s.layout); t = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle')); }
      if (t && L.txt.plain(t.tx).replace(/\n/g, ' ') !== g.title) L.txt.setPlain(t.tx, g.title);
      let b = bodyShape(s);
      if (!b && g.body.some((x) => x.text)) { M.applyLayout(pres, s, s.layout === 'titleOnly' || s.layout === 'blank' ? 'text' : s.layout); b = bodyShape(s); }
      if (b) {
        const old = b.tx.ps;
        b.tx.ps = (g.body.length ? g.body : [{ lvl: 0, text: '' }]).map((r, k) => {
          const src = old[k] || old[old.length - 1] || L.txt.para('');
          if (old[k] && L.txt.paraText(old[k]).replace(/\n/g, ' ') === r.text && (old[k].lvl || 0) === r.lvl) return old[k];
          const rp = src.rs[0] ? L.txt.runProps(src.rs[0]) : src.end || {};
          return { lvl: r.lvl, pp: L.clone(src.pp || {}), rs: r.text ? [Object.assign({ t: r.text }, rp)] : [], end: L.clone(rp) };
        });
      }
    });
    /* titles removed: drop slides that only held text */
    for (let i = pres.slides.length - 1; i >= groups.length; i--) {
      const s = pres.slides[i];
      const onlyText = s.shapes.every((x) => x.ph && x.type === 'text');
      if (onlyText && pres.slides.length > 1) pres.slides.splice(i, 1);
    }
    if (E().idx >= pres.slides.length) E().idx = pres.slides.length - 1;
    olHistPushed = false;
    E().render();
    L.bus.emit('slides-changed', { keepOutline: true });
    if (!keepFocus) renderOutline();
    else styleOutline();
  };
  L.app = L.app || {};

  /* ================= notes pane ================= */
  const NP = (P.notes = {});
  NP.mount = function (el) {
    NP.el = el;
    NP.ed = h('div', { class: 'notes-in', contenteditable: 'true', spellcheck: 'true', role: 'textbox', 'aria-multiline': 'true', 'aria-label': '演讲者备注', 'data-ph': L.t('单击此处添加备注') });
    el.appendChild(NP.ed);
    NP.ed.addEventListener('input', () => {
      const s = E().slide();
      if (!s || s.isMaster) return;
      L.hist.push(L.N('备注'), 'notes:' + s.id);
      s.notes = NP.ed.innerText.replace(/\n$/, '');
      if (!NP.ed.textContent) NP.ed.innerHTML = '';
    });
    NP.ed.addEventListener('paste', (e) => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData && e.clipboardData.getData('text/plain')) || ''); });
    NP.ed.addEventListener('focus', () => { if (L.te.active()) L.te.end(); });
  };
  NP.render = function () {
    if (!NP.ed) return;
    const s = E().slide();
    NP.ed.contentEditable = s && !s.isMaster ? 'true' : 'false';
    NP.ed.textContent = s && !s.isMaster ? s.notes || '' : '';
  };

  /* ================= Slide Sorter ================= */
  const SO = (P.sorter = { zoom: 0.25 });
  SO.mount = function (parent) {
    SO.el = h('div', { class: 'sorter', tabindex: '-1', 'aria-label': '幻灯片浏览' });
    parent.appendChild(SO.el);
    SO.el.addEventListener('pointerdown', (e) => { if (e.target === SO.el || e.target.classList.contains('sorter-grid')) { L.app.slideSel = new Set([E().idx]); SO.mark(); } });
    SO.el.addEventListener('focus', () => { L.app.focusArea = 'slides'; });
    SO.el.addEventListener('blur', () => { L.app.focusArea = null; });
  };
  SO.render = function () {
    if (!SO.el) return;
    const grid = h('div', { class: 'sorter-grid' });
    const w = Math.max(90, Math.round(SO.el.clientWidth ? Math.min(260, (SO.el.clientWidth - 60) / 4) : 180));
    L.pres.slides.forEach((s, i) => {
      const it = h('div', { class: 'so-it' + (L.app.slideSel.has(i) ? ' sel' : '') + (s.hidden ? ' hiddenslide' : ''), 'data-i': i });
      const fr = h('button', { class: 'so-frame', type: 'button', 'aria-label': L.t('幻灯片 {0}', i + 1) });
      fr.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
      const foot = h('div', { class: 'so-foot' });
      if (s.trans && s.trans.type && s.trans.type !== 'none') {
        const star = h('span', { class: 'th-star', 'data-tip': L.anim.transName(s.trans), html: '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333"/></svg>' });
        star.addEventListener('click', (e) => { e.stopPropagation(); L.show.previewTransition(fr, L.pres, s, L.pres.slides[i - 1]); });
        foot.appendChild(star);
      }
      if (s.trans && s.trans.after != null) foot.appendChild(h('span', { text: fmtSecs(s.trans.after) }));
      foot.appendChild(h('span', { class: 'so-num' + (s.hidden ? ' hid' : ''), text: String(i + (L.pres.firstNum || 1)) }));
      it.append(fr, foot);
      it.addEventListener('pointerdown', (e) => onSorterDown(e, i, it));
      it.addEventListener('dblclick', () => { E().goto(i); L.app.setView('normal'); });
      it.addEventListener('contextmenu', (e) => { e.preventDefault(); if (!L.app.slideSel.has(i)) { L.app.slideSel = new Set([i]); E().idx = i; SO.mark(); } L.app.contextThumb(e); });
      grid.appendChild(it);
    });
    L.clear(SO.el).appendChild(grid);
  };
  const fmtSecs = (ms) => { const s = Math.round(ms / 1000); return `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; };
  SO.mark = function () { L.$$('.so-it', SO.el).forEach((el) => el.classList.toggle('sel', L.app.slideSel.has(+el.dataset.i))); L.app.updateStatus(); L.ui.refresh(); };
  function onSorterDown(e, i, it) {
    if (e.button !== 0) return;
    SO.el.focus({ preventScroll: true });
    const sel = L.app.slideSel;
    if (e.shiftKey) { const a = Math.min(E().idx, i), b = Math.max(E().idx, i); L.app.slideSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); SO.mark(); return; }
    if (e.ctrlKey || e.metaKey) { if (sel.has(i) && sel.size > 1) sel.delete(i); else sel.add(i); SO.mark(); return; }
    if (!sel.has(i)) L.app.slideSel = new Set([i]);
    E().idx = i;
    SO.mark();
    /* drag to reorder (common/ui.js ui.drag) */
    let marker = null, target = null;
    const mv = (ev) => {
      if (!marker) { marker = h('div', { style: 'position:absolute;width:3px;background:#000;pointer-events:none;z-index:5' }); SO.el.appendChild(marker); }
      const items = L.$$('.so-it', SO.el);
      target = items.length;
      let ref = null, before = true;
      for (let k = 0; k < items.length; k++) {
        const r = items[k].getBoundingClientRect();
        if (ev.clientY >= r.top - 12 && ev.clientY <= r.bottom + 12 && ev.clientX < r.left + r.width / 2) { target = k; ref = items[k]; before = true; break; }
        if (ev.clientY >= r.top - 12 && ev.clientY <= r.bottom + 12) { ref = items[k]; before = false; target = k + 1; }
      }
      if (!ref) return;
      const br = SO.el.getBoundingClientRect(), rr = ref.getBoundingClientRect();
      marker.style.left = (before ? rr.left - 12 : rr.right + 10) - br.left + SO.el.scrollLeft + 'px';
      marker.style.top = rr.top - br.top + SO.el.scrollTop + 'px';
      marker.style.height = rr.height + 'px';
    };
    L.ui.drag(e, {
      threshold: 6,
      move: mv,
      up() { if (marker) { marker.remove(); if (target != null) L.app.moveSlides(Array.from(L.app.slideSel).sort((a, b) => a - b), target); } },
      cancel() { if (marker) marker.remove(); },
    });
    void it;
  }

  /* ================= Notes Page view ================= */
  const NV = (P.notesView = {});
  NV.mount = function (parent) { NV.el = h('div', { class: 'notespage', 'aria-label': '备注页' }); parent.appendChild(NV.el); };
  NV.render = function () {
    if (!NV.el) return;
    const s = E().slide();
    L.clear(NV.el);
    if (!s) return;
    const avail = Math.max(300, NV.el.clientWidth - 60);
    const pw = Math.min(avail, 640), ph = pw * (10 / 7.5);
    const page = h('div', { class: 'np-page', style: `width:${pw}px;height:${ph}px` });
    const sw = pw * 0.67, shh = (sw * L.pres.H) / L.pres.W;
    const sb = h('div', { class: 'np-slide', style: `left:${(pw - sw) / 2}px;top:${ph * 0.075}px;width:${sw}px;height:${shh}px` });
    sb.appendChild(L.render.thumb(L.pres, s, sw, { index: E().idx }));
    const ny = ph * 0.075 + shh + ph * 0.05;
    const notes = h('div', { class: 'np-notes', contenteditable: 'true', spellcheck: 'true', style: `left:${pw * 0.1}px;top:${ny}px;width:${pw * 0.8}px;height:${ph - ny - ph * 0.08}px;font-size:${Math.max(11, pw / 48)}px`, role: 'textbox', 'aria-label': '备注' });
    notes.textContent = s.notes || '';
    notes.addEventListener('input', () => { L.hist.push(L.N('备注'), 'notes:' + s.id); s.notes = notes.innerText.replace(/\n$/, ''); });
    notes.addEventListener('paste', (e) => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData && e.clipboardData.getData('text/plain')) || ''); });
    page.append(sb, notes, h('div', { class: 'pv-hdr', style: `right:${pw * 0.05}px;bottom:${ph * 0.03}px`, text: String(E().idx + (L.pres.firstNum || 1)) }));
    NV.el.appendChild(page);
  };

  /* ================= Print Preview ================= */
  const PV = (P.preview = {});
  PV.mount = function (parent) {
    PV.el = h('div', { class: 'preview', 'aria-label': '打印预览' });
    PV.bar = h('div', { class: 'pv-bar' });
    PV.scroll = h('div', { class: 'pv-scroll' });
    PV.el.append(PV.bar, PV.scroll);
    parent.appendChild(PV.el);
  };
  PV.render = function () {
    if (!PV.el) return;
    const o = L.app.printOpts;
    L.clear(PV.bar);
    const canPrint = window.self === window.top;
    if (canPrint) PV.bar.appendChild(ui.button('打印(&P)...', () => { document.body.classList.add('printing'); window.print(); setTimeout(() => document.body.classList.remove('printing'), 500); }));
    PV.bar.append(
      h('label', { text: '打印内容：' }),
      ui.select([['slides', '幻灯片'], ['handouts', '讲义'], ['notes', '备注页'], ['outline', '大纲视图']], o.what, (v) => { o.what = v; PV.render(); }),
      h('label', { text: '每页：' }),
      ui.select([[1, '1'], [2, '2'], [3, '3'], [4, '4'], [6, '6'], [9, '9']], o.per, (v) => { o.per = +v; PV.render(); }, { disabled: o.what !== 'handouts' }),
      ui.select([['color', '颜色'], ['gray', '灰度'], ['bw', '纯黑白']], o.color, (v) => { o.color = v; PV.render(); }),
      ui.check('幻灯片加框', o.frame, (v) => { o.frame = v; PV.render(); }),
      ui.button('另存为 PDF(&D)...', () => L.app.previewPDF()),
      ui.button('关闭(&C)', () => L.app.setView(L.app.prevView || 'normal')));
    if (!canPrint) PV.bar.appendChild(h('span', { class: 'tp-note', text: '要打印，请将这些页面保存为 PDF，再打印该 PDF。' }));
    L.clear(PV.scroll);
    const slides = L.pres.slides.map((s, i) => [s, i]).filter(([s, i]) => (o.hidden || !s.hidden) && (!o.range || o.range.includes(i)));
    const landscape = o.what === 'slides';
    const pw = Math.min(PV.scroll.clientWidth - 40 || 700, landscape ? 760 : 560), ph = landscape ? pw * (8.5 / 11) : pw * (11 / 8.5);
    const filt = o.color === 'gray' ? 'grayscale(1)' : o.color === 'bw' ? 'grayscale(1) contrast(30)' : '';
    const page = () => { const p = h('div', { class: 'pv-page', style: `width:${pw}px;height:${ph}px` }); if (filt) p.style.filter = filt; PV.scroll.appendChild(p); return p; };
    const cell = (p, s, i, x, y, w) => {
      const hh = (w * L.pres.H) / L.pres.W;
      const c = h('div', { class: 'pv-cell', style: `left:${x}px;top:${y}px;width:${w}px;height:${hh}px;${o.frame || o.what !== 'slides' ? '' : 'border-color:transparent'}` });
      c.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
      p.appendChild(c);
      return hh;
    };
    const footer = (p, n) => { p.appendChild(h('div', { class: 'pv-hdr', style: `right:${pw * 0.05}px;bottom:${ph * 0.025}px`, text: String(n) })); };
    if (o.what === 'slides') {
      slides.forEach(([s, i]) => { const p = page(); const w = Math.min(pw * 0.9, ((ph * 0.9) * L.pres.W) / L.pres.H); const hh = (w * L.pres.H) / L.pres.W; cell(p, s, i, (pw - w) / 2, (ph - hh) / 2, w); });
    } else if (o.what === 'notes') {
      slides.forEach(([s, i], n) => {
        const p = page();
        const w = pw * 0.67; const hh = cell(p, s, i, (pw - w) / 2, ph * 0.07, w);
        p.appendChild(h('div', { class: 'pv-text', style: `left:${pw * 0.1}px;top:${ph * 0.07 + hh + ph * 0.04}px;width:${pw * 0.8}px;height:${ph * 0.4}px;font-size:${pw / 50}px`, text: s.notes || '', translate: 'no' }));
        footer(p, n + 1);
      });
    } else if (o.what === 'outline') {
      let p = page(), y = ph * 0.06, pn = 1;
      const lh = pw / 40;
      const line = (t, indent, bold) => { if (y > ph * 0.92) { footer(p, pn++); p = page(); y = ph * 0.06; } p.appendChild(h('div', { class: 'pv-text', style: `left:${pw * 0.08 + indent}px;top:${y}px;width:${pw * 0.84 - indent}px;font-size:${lh * 0.85}px;${bold ? 'font-weight:bold' : ''}`, text: t, translate: 'no' })); y += lh * 1.3; };
      slides.forEach(([s, i]) => {
        line(`${i + 1}  ${M.slideTitle(s)}`, 0, true);
        const b = bodyShape(s);
        if (b) for (const para of b.tx.ps) { const t = L.txt.paraText(para); if (t) line('• ' + t, 24 + (para.lvl || 0) * 18, false); }
      });
      footer(p, pn);
    } else {
      const per = o.per;
      const layouts = { 1: [[0.5, 0.5, 0.8]], 2: [[0.5, 0.28, 0.6], [0.5, 0.72, 0.6]], 3: [[0.3, 0.2, 0.42], [0.3, 0.5, 0.42], [0.3, 0.8, 0.42]], 4: [[0.28, 0.3, 0.42], [0.72, 0.3, 0.42], [0.28, 0.7, 0.42], [0.72, 0.7, 0.42]], 6: [[0.28, 0.2, 0.4], [0.72, 0.2, 0.4], [0.28, 0.5, 0.4], [0.72, 0.5, 0.4], [0.28, 0.8, 0.4], [0.72, 0.8, 0.4]], 9: [[0.2, 0.2, 0.28], [0.5, 0.2, 0.28], [0.8, 0.2, 0.28], [0.2, 0.5, 0.28], [0.5, 0.5, 0.28], [0.8, 0.5, 0.28], [0.2, 0.8, 0.28], [0.5, 0.8, 0.28], [0.8, 0.8, 0.28]] }[per];
      for (let k = 0, pn = 1; k < slides.length; k += per, pn++) {
        const p = page();
        layouts.forEach(([cx, cy, wf], j) => {
          const it = slides[k + j];
          if (!it) return;
          const w = pw * wf, hh = (w * L.pres.H) / L.pres.W;
          cell(p, it[0], it[1], pw * cx - w / 2, ph * cy - hh / 2, w);
          if (per === 3) { const lines = h('div', { class: 'pv-lines', style: `left:${pw * 0.56}px;top:${ph * cy - hh / 2}px;width:${pw * 0.36}px;height:${hh}px` }); for (let q = 0; q < 7; q++) lines.appendChild(h('span')); p.appendChild(lines); }
        });
        footer(p, pn);
      }
    }
    if (!slides.length) PV.scroll.appendChild(h('div', { text: '没有可打印的内容。', style: 'color:#fff' }));
  };

  /* ================= task panes ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    TP.back = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '后退', html: L.icons.get('back') });
    TP.fwd = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '向前', html: L.icons.get('forward') });
    TP.home = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '主页', html: L.icons.get('home') });
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '关闭任务窗格', html: L.icons.get('close', 12) });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => L.app.toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id]) return;
    L.app.toggleTask(true);
    if (TP.stack[TP.pos] !== id) { TP.stack = TP.stack.slice(0, TP.pos + 1); TP.stack.push(id); TP.pos = TP.stack.length - 1; }
    TP.render();
  };
  TP.go = function (pos) { if (pos < 0 || pos >= TP.stack.length) return; TP.pos = pos; TP.render(); };
  TP.render = function () {
    const id = TP.stack[TP.pos];
    TP.current = id;
    const p = TP.registry[id];
    if (!p) return;
    TP.title.firstChild.textContent = L.t(p.title);
    TP.back.disabled = TP.pos <= 0; TP.fwd.disabled = TP.pos >= TP.stack.length - 1;
    L.clear(TP.body);
    p.render(TP.body);
  };
  TP.refresh = function (reason) {
    const p = TP.registry[TP.current];
    if (!p || !TP.el || TP.el.hidden) return;
    if (p.refresh) p.refresh(TP.body, reason);
    else if (p.live) { L.clear(TP.body); p.render(TP.body); }
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn) => h('button', { class: 'tp-link', type: 'button', onclick: fn }, icon ? h('span', { html: L.icons.get(icon) }) : null, h('span', { text: label }));

  /* --- Getting Started --- */
  /* ---------- Document Recovery: unsaved versions kept by common/suite.js ---------- */
  TP.register({ id: 'recovery', title: '文档恢复', render(b) { if (window.VO) VO.recoveryPane(b); } });
  TP.register({
    id: 'getting-started', title: '入门',
    render(b) {
      const q = h('input', { type: 'text', id: 'gs-search', placeholder: '搜索帮助', style: 'flex:1;min-width:0' });
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { P.helpQuery = q.value; TP.show('help'); } });
      b.append(
        h('div', { class: 'tp-row' }, h('span', { html: L.icons.app(32) }), h('b', { text: 'Lectern 网络版', style: 'color:var(--pane-ink)' })),
        sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec(L.N('打开'), window.VO ? VO.recentLinks(4) : null, link('open', L.N('打开演示文稿 (.pptx)...'), () => ui.exec('open')), link('new', '新建演示文稿...', () => TP.show('new')), link('slideshow', '查看示例演示文稿', () => L.app.loadSample())),
        sec(L.N('学习'), link('help', '键盘快捷键和帮助', () => TP.show('help')), link('design', '应用设计模板', () => TP.show('design')), link('animation', '为幻灯片添加动画', () => TP.show('customAnim'))));
    },
  });

  /* --- New Presentation --- */
  TP.register({
    id: 'new', title: '新建演示文稿',
    render(b) {
      b.append(
        sec(L.N('新建'), link('new', '空白演示文稿', () => L.app.newPresentation()), link('design', '根据设计模板', () => { L.app.newPresentation(); TP.show('design'); }), link('properties', '根据内容提示向导...', () => L.app.autoContent()), link('open', '根据现有演示文稿...', () => ui.exec('open')), link('photoAlbum', '相册...', () => ui.exec('photoAlbum'))),
        sec(L.N('模板'), link('slideshow', '示例：咖啡馆扩展计划', () => L.app.loadSample())));
    },
  });

  /* --- Help --- */
  const HELP = [
    [L.N('浏览与定位'), L.N('使用左侧“幻灯片”选项卡切换幻灯片，拖动缩略图可重新排序。“大纲”选项卡可输入标题和项目符号，按 Tab 降级，Shift+Tab 升级。在标题行按 Enter 可新建幻灯片。')],
    [L.N('添加文本'), L.N('单击占位符（“单击此处添加标题”）并键入内容。使用“绘图”工具栏中的“文本框”工具可在任意位置添加文字。单击文本框即可开始键入，拖动其边框可移动文本框。')],
    [L.N('设置文字格式'), L.N('选择文字后，使用“格式”工具栏或“格式 ▸ 字体”。项目符号和编号位于“格式 ▸ 项目符号和编号”。在项目符号开头按 Tab 可将其降级。')],
    [L.N('绘制形状'), L.N('在“绘图”工具栏中选择“自选图形”，然后在幻灯片上拖动。按住 Shift 可绘制正方形、正圆或以 15° 为单位的角度；按住 Ctrl 可从中心绘制。黄色菱形用于调整形状比例，绿色圆点用于旋转。')],
    [L.N('图片、表格和图表'), L.N('使用“插入”菜单或内容占位符中的图标。双击图表可编辑其数据表，双击图片可裁剪或重新着色。')],
    [L.N('设计模板'), L.N('“格式 ▸ 幻灯片设计”会显示设计库。该窗格顶部提供“配色方案”和“动画方案”链接。')],
    [L.N('动画'), L.N('“幻灯片放映 ▸ 自定义动画”：选择对象，选择“添加效果”，再设置“开始”、“方向”和“速度”。使用“按第一级段落”（“效果选项 ▸ 文本动画”）可逐条显示项目符号。')],
    [L.N('幻灯片放映'), L.N('按 F5 从第一张幻灯片开始放映，或按 Shift+F5 从当前幻灯片开始。单击、空格或 → 可前进，← 可后退，B 可显示黑屏，Ctrl+P 可切换到笔，Esc 可结束放映。')],
    [L.N('文件'), L.N('“文件 ▸ 打开”可读取 .pptx、.ppsx 和 .potx 文件。“文件 ▸ 保存”可生成能在 PowerPoint、Keynote、Google Slides 和 LibreOffice 中打开的 .pptx。“另存为”还提供 .ppsx、.potx、单文件网页、当前幻灯片的 PNG 图片以及文本大纲。')],
  ];
  const KEYS = [['Ctrl+N / Ctrl+O / Ctrl+S', L.N('新建、打开、保存')], ['Ctrl+M', L.N('新建幻灯片')], ['Ctrl+D', L.N('创建副本')], ['Ctrl+Z / Ctrl+Y', L.N('撤消/重做')], ['Ctrl+X / C / V', L.N('剪切、复制、粘贴')], ['Ctrl+A', L.N('全选')], ['Ctrl+G / Ctrl+Shift+G', L.N('组合/取消组合')], ['Ctrl+B / I / U', L.N('加粗、倾斜、下划线')], ['Ctrl+E / L / R / J', L.N('居中、左对齐、右对齐、两端对齐')], ['Ctrl+Shift+> / <', L.N('增大/减小字号')], ['Ctrl+T', L.N('字体对话框')], ['Shift+F3', L.N('更改大小写')], ['Ctrl+K', L.N('超链接')], ['Ctrl+F / Ctrl+H', L.N('查找/替换')], ['Tab / Shift+Tab', L.N('降级/升级；下一个对象')], [L.N('方向键'), L.N('微移（按住 Ctrl 可精细微移）')], ['F2', L.N('在文字选择和对象选择之间切换')], ['F5 / Shift+F5', L.N('从头/从当前幻灯片放映')], ['Page Up / Page Down', L.N('上一张/下一张幻灯片')], ['Ctrl+F1', L.N('任务窗格')]];
  TP.register({
    id: 'help', title: '帮助',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: '搜索', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        L.clear(res);
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        /* searched as shown: the topics in the UI's language, the keys with their names there (Strg+M) */
        const topics = HELP.map(([t, d]) => [L.t(t), L.t(d)]).filter(([t, d]) => !words.length || hit(t + ' ' + d));
        for (const [t, d] of topics) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        const keyName = (k) => k.split(' / ').map((x) => (/ [a-z]/.test(x) ? L.t(x) : ui.keyText(x))).join(' / ');
        const keys = KEYS.map(([k, d]) => [keyName(k), L.t(d)]).filter(([k, d]) => !words.length || hit(k + ' ' + d));
        if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: '键盘快捷键' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { text: k, translate: 'no' }), h('td', { text: d }))))));
        if (!res.children.length) res.appendChild(h('div', { class: 'tp-note', text: '没有匹配的主题。请尝试其他词语。' }));
      };
      q.addEventListener('input', draw);
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
    },
  });

  /* --- Clip Art --- */
  TP.register({
    id: 'clipart', title: '剪贴画',
    render(b) {
      const q = h('input', { type: 'text', id: 'ca-q', placeholder: '例如：人物、成长、时间', style: 'flex:1;min-width:0' });
      const grid = h('div', { class: 'clip-grid' });
      const draw = () => {
        L.clear(grid);
        const list = L.clipart.search(q.value);
        for (const it of list) {
          const c = h('button', { class: 'clip-it', type: 'button', 'data-tip': it.name }, h('img', { src: L.clipart.url(it), alt: it.name }), h('span', { text: it.name }));
          c.addEventListener('click', () => L.app.insertClipArt(it));
          grid.appendChild(c);
        }
        if (!list.length) grid.appendChild(h('div', { class: 'tp-note', text: '未找到剪贴画。' }));
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') draw(); });
      b.append(sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', draw, { class: 'btn small' }))), sec(L.N('结果'), grid), h('div', { class: 'tp-note', text: '单击剪贴画可将其插入当前幻灯片。' }));
      draw();
    },
  });

  /* --- Clipboard --- */
  TP.register({
    id: 'clipboard', title: '剪贴板', live: true,
    render(b) {
      const items = L.app.clipHistory;
      b.append(h('div', { class: 'tp-btns' }, ui.button('全部粘贴', () => { for (const it of items.slice().reverse()) L.app.pasteItem(it); }, { class: 'btn small' }), ui.button('全部清除', () => { items.length = 0; TP.refresh(); }, { class: 'btn small' })), h('div', { class: 'tp-note', text: L.t('{0}/24 — 单击项目可粘贴。', items.length) }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { class: 'cb-ico', html: L.icons.get(it.kind === 'slides' ? 'sorterView' : it.kind === 'text' ? 'font' : 'autoshapes') }), h('span', { class: 'cb-t', text: it.label, translate: 'no' }));
        el.addEventListener('click', () => L.app.pasteItem(it));
        b.appendChild(el);
      }
    },
  });

  /* --- Slide Layout --- */
  /** a layout's picture, in the presentation's shape */
  function layoutSVG(key, w) {
    const W = L.pres ? L.pres.W : M.DEFAULT_SIZE.w, H = L.pres ? L.pres.H : M.DEFAULT_SIZE.h;
    const d = L.model.buildDesign('default', W, H);
    const frames = M.layoutFrames(key, d);
    const k = w / W, hh = H * k;
    let body = '';
    for (const f of frames) {
      const x = f.x * k, y = f.y * k, fw = f.w * k, fh = f.h * k;
      if (f.type === 'title' || f.type === 'ctrTitle') body += f.vert ? `<rect x="${x + fw * 0.35}" y="${y + 3}" width="${fw * 0.3}" height="${fh - 6}" fill="#7a7a7a"/>` : `<rect x="${x + fw * 0.15}" y="${y + fh * 0.35}" width="${fw * 0.7}" height="${Math.max(3, fh * 0.3)}" fill="#7a7a7a"/>`;
      else if (f.type === 'subTitle') body += `<rect x="${x + fw * 0.25}" y="${y + fh * 0.3}" width="${fw * 0.5}" height="${Math.max(2, fh * 0.12)}" fill="#a8a8a8"/>`;
      else if (f.type === 'body') { if (f.vert) for (let i = 0; i < 4; i++) body += `<rect x="${x + fw * 0.85 - i * fw * 0.15}" y="${y + 4}" width="2" height="${fh * 0.7}" fill="#a8a8a8"/>`; else for (let i = 0; i < 4; i++) body += `<circle cx="${x + 5}" cy="${y + 6 + i * Math.min(fh / 5, 9)}" r="1.2" fill="#555"/><rect x="${x + 9}" y="${y + 5 + i * Math.min(fh / 5, 9)}" width="${fw * 0.75}" height="2" fill="#a8a8a8"/>`; }
      else body += `<rect x="${x + 2}" y="${y + 2}" width="${fw - 4}" height="${fh - 4}" fill="none" stroke="#b0b0b0" stroke-dasharray="2 1"/><g transform="translate(${x + fw / 2 - 9},${y + fh / 2 - 6})">${f.type === 'tbl' ? '<rect width="18" height="12" fill="#fff" stroke="#3a5a8c"/><path d="M0 4h18M0 8h18M6 0v12M12 0v12" stroke="#3a5a8c"/>' : f.type === 'chart' ? '<rect x="1" y="6" width="4" height="6" fill="#9999ff"/><rect x="7" y="2" width="4" height="10" fill="#993366"/><rect x="13" y="4" width="4" height="8" fill="#ffffcc" stroke="#999" stroke-width=".5"/>' : '<rect x="0" y="0" width="5" height="5" fill="#9cb7e3"/><rect x="6.5" y="0" width="5" height="5" fill="#9999ff"/><rect x="13" y="0" width="5" height="5" fill="#9fd39a"/><rect x="0" y="7" width="5" height="5" fill="#bfe1ff"/><rect x="6.5" y="7" width="5" height="5" fill="#cfe0fb"/><rect x="13" y="7" width="5" height="5" fill="#ccc"/>'}</g>`;
    }
    return `<svg width="${w}" height="${hh}" viewBox="0 0 ${w} ${hh}">${body}</svg>`;
  }
  P.layoutSVG = layoutSVG;
  /* the headings of M.LAYOUTS' groups */
  const LAYOUT_GROUPS = { '文字版式': L.N('文字版式：'), '内容版式': L.N('内容版式：'), '文字和内容版式': L.N('文字和内容版式：'), '其他版式': L.N('其他版式：') };
  TP.register({
    id: 'layout', title: '幻灯片版式',
    render(b) {
      b.appendChild(h('div', { class: 'tp-note', text: '应用幻灯片版式：' }));
      const groups = {};
      for (const l of M.LAYOUTS) (groups[l.group] = groups[l.group] || []).push(l);
      const cur = E().slide() ? E().slide().layout : null;
      for (const g in groups) {
        const grid = h('div', { class: 'lay-grid' });
        for (const l of groups[g]) {
          const it = h('button', { class: 'lay-it' + (l.key === cur ? ' on' : ''), type: 'button', 'data-key': l.key, 'data-tip': l.name, 'aria-label': l.name, html: layoutSVG(l.key, 64) });
          it.addEventListener('click', () => L.app.applyLayoutToSel(l.key));
          it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: '应用于选定幻灯片(&S)', run: () => L.app.applyLayoutToSel(l.key) }, { label: '重新应用母版样式(&R)', run: () => L.app.applyLayoutToSel(l.key, true) }, '-', { label: '插入新幻灯片(&N)', run: () => L.app.insertSlide(l.key) }], { x: e.clientX, y: e.clientY }); });
          grid.appendChild(it);
        }
        b.appendChild(sec(LAYOUT_GROUPS[g] || g, grid));
      }
      b.appendChild(ui.check('插入新幻灯片时显示', !!L.app.opts.layoutPaneOnNew, (v) => { L.app.opts.layoutPaneOnNew = v; L.app.saveOpts(); }));
    },
    refresh(b) { L.$$('.lay-it', b).forEach((el) => el.classList.toggle('on', !!E().slide() && el.dataset.key === M.layoutInfo(E().slide().layout).key)); },
  });

  /* --- Slide Design --- */
  const previewCache = {};
  function designPreview(key, w) {
    const ck = key + '|' + w + '|' + L.pres.W + 'x' + L.pres.H;
    if (previewCache[ck]) return previewCache[ck].cloneNode(true);
    const pres = M.newPresentation({ design: key, w: L.pres.W, h: L.pres.H, empty: true });
    const did = Object.keys(pres.designs)[0];
    const s = M.newSlide(pres, 'text', did);
    const t = s.shapes.find((x) => x.ph.type === 'title');
    L.txt.setPlain(t.tx, L.t('标题'));
    const b = s.shapes.find((x) => x.ph.type === 'body');
    b.tx.ps = [L.txt.para(L.t('项目符号文字')), L.txt.para(L.t('第二行'))];
    pres.slides.push(s);
    const el = L.render.thumb(pres, s, w);
    previewCache[ck] = el;
    return el.cloneNode(true);
  }
  TP.register({
    id: 'design', title: '幻灯片设计',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', '设计模板', () => TP.show('design')), link('colorPic', '配色方案', () => TP.show('colors')), link('animation', '动画方案', () => TP.show('animSchemes'))));
      const cur = E().design();
      const used = Array.from(new Set(L.pres.slides.map((s) => s.design))).map((id) => L.pres.designs[id]).filter(Boolean);
      const usedGrid = h('div', { class: 'dsn-grid' });
      for (const d of used) {
        const it = h('button', { class: 'dsn-it' + (d === cur ? ' on' : ''), type: 'button', 'data-tip': d.name });
        const tmp = L.pres.slides.find((s) => s.design === d.id);
        it.append(L.render.thumb(L.pres, Object.assign({}, tmp, { shapes: tmp.shapes.filter((x) => x.ph && x.type === 'text') }), 100, { prompts: false }), h('span', { class: 'dsn-name', text: d.name }));
        it.addEventListener('click', () => { const sel = L.app.selectedSlides(); L.hist.push('应用设计模板'); for (const s of sel) s.design = d.id; E().render(); L.bus.emit('slides-changed'); });
        usedGrid.appendChild(it);
      }
      b.appendChild(sec('此演示文稿中使用的', usedGrid));
      const grid = h('div', { class: 'dsn-grid' });
      for (const t of M.DESIGNS) {
        const it = h('button', { class: 'dsn-it', type: 'button', 'data-tip': t.name, 'aria-label': t.name });
        it.append(designPreview(t.key, 100), h('span', { class: 'dsn-name', text: t.name }), h('span', { class: 'dsn-dd', html: '<span class="dd-arrow" style="margin:0"></span>' }));
        it.addEventListener('click', (e) => {
          if (e.target.closest('.dsn-dd')) { ui.openMenu([{ label: '应用于所有幻灯片(&A)', run: () => L.app.applyDesign(t.key, true) }, { label: '应用于选定幻灯片(&S)', run: () => L.app.applyDesign(t.key, false) }], { x: e.clientX, y: e.clientY }); return; }
          L.app.applyDesign(t.key, true);
        });
        it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: '应用于所有幻灯片(&A)', run: () => L.app.applyDesign(t.key, true) }, { label: '应用于选定幻灯片(&S)', run: () => L.app.applyDesign(t.key, false) }], { x: e.clientX, y: e.clientY }); });
        grid.appendChild(it);
      }
      b.appendChild(sec('可供使用', grid));
      b.appendChild(link('open', L.N('浏览...（使用其他 .pptx/.potx 文件的设计）'), () => L.app.browseDesign()));
    },
    refresh(b, reason) { if (reason === 'slides' || reason === 'design') { L.clear(b); this.render(b); } },
  });

  /* --- Color Schemes --- */
  TP.register({
    id: 'colors', title: '幻灯片设计 - 配色方案',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', '设计模板', () => TP.show('design')), link('colorPic', '配色方案', () => TP.show('colors')), link('animation', '动画方案', () => TP.show('animSchemes'))));
      const grid = h('div', { class: 'scheme-grid' });
      const d0 = E().design();
      const schemes = [d0.colors].concat(M.COLOR_SCHEMES);
      schemes.forEach((cs, i) => {
        const it = h('button', { class: 'scheme-it', type: 'button', 'aria-label': i === 0 ? L.t('当前方案') : L.t('配色方案 {0}', i), 'data-tip': i === 0 ? L.N('当前配色方案') : '应用此配色方案' });
        for (const k of ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'hlink', 'folHlink']) it.appendChild(h('span', { style: `background:${cs[k]}` }));
        it.addEventListener('click', () => L.app.applyColorScheme(cs, true));
        it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: '应用于所有幻灯片(&A)', run: () => L.app.applyColorScheme(cs, true) }, { label: '应用于选定幻灯片(&S)', run: () => L.app.applyColorScheme(cs, false) }], { x: e.clientX, y: e.clientY }); });
        grid.appendChild(it);
      });
      b.append(sec('应用配色方案：', grid), link('options', '编辑配色方案...', () => L.app.editColorScheme()));
    },
    refresh(b, reason) { if (reason === 'design') { L.clear(b); this.render(b); } },
  });

  /* --- Animation Schemes --- */
  const AS = [
    [L.N('无动画'), null], [L.N('细微型'), 'group'],
    [L.N('出现'), { tr: { type: 'cut' }, title: ['appear'], body: ['appear'] }], [L.N('出现并变暗'), { tr: { type: 'fade' }, title: ['appear'], body: ['appear'] }],
    [L.N('全部淡入'), { tr: { type: 'fade' }, title: ['fade'], body: ['fade', null, 'all'] }], [L.N('逐个淡入'), { tr: { type: 'fade' }, title: ['fade'], body: ['fade'] }],
    [L.N('淡化擦除'), { tr: { type: 'fade' }, title: ['fade'], body: ['wipe', 'b'] }], [L.N('淡化缩放'), { tr: { type: 'fade' }, title: ['fade'], body: ['zoom'] }],
    [L.N('溶解进入'), { tr: { type: 'dissolve' }, title: ['dissolve'], body: ['dissolve'] }], [L.N('随机线条'), { tr: { type: 'randomBar', dir: 'horz' }, title: ['randomBars', 'horz'], body: ['randomBars', 'horz'] }],
    [L.N('擦除'), { tr: { type: 'wipe', dir: 'r' }, title: ['wipe', 'l'], body: ['wipe', 't'] }], [L.N('百叶窗'), { tr: { type: 'blinds', dir: 'horz' }, title: ['blinds', 'horz'], body: ['blinds', 'horz'] }],
    [L.N('温和型'), 'group'],
    [L.N('升起'), { tr: { type: 'fade' }, title: ['ascend'], body: ['ascend'] }], [L.N('下降'), { tr: { type: 'fade' }, title: ['descend'], body: ['descend'] }],
    [L.N('缩放'), { tr: { type: 'zoom', dir: 'out' }, title: ['zoom'], body: ['zoom'] }], [L.N('展开'), { tr: { type: 'fade' }, title: ['expand'], body: ['expand'] }],
    [L.N('旋转器'), { tr: { type: 'fade' }, title: ['spinner'], body: ['spinner'] }], [L.N('缓慢进入'), { tr: { type: 'cover', dir: 'u' }, title: ['peekIn', 'b'], body: ['peekIn', 'b'] }],
    [L.N('华丽型'), 'group'],
    [L.N('飞入'), { tr: { type: 'push', dir: 'u' }, title: ['flyIn', 't'], body: ['flyIn', 'l'] }], [L.N('棋盘'), { tr: { type: 'checker', dir: 'horz' }, title: ['checkerboard', 'across'], body: ['checkerboard', 'across'] }],
    [L.N('轮子'), { tr: { type: 'wheel', spokes: 4 }, title: ['wheel', '4'], body: ['wheel', '1'] }], [L.N('加号'), { tr: { type: 'plus' }, title: ['plus', 'out'], body: ['plus', 'out'] }],
    [L.N('菱形'), { tr: { type: 'diamond' }, title: ['diamond', 'out'], body: ['diamond', 'out'] }], [L.N('盒状展开'), { tr: { type: 'zoom', dir: 'out' }, title: ['box', 'out'], body: ['box', 'out'] }],
  ];
  P.ANIM_SCHEMES = AS;
  TP.register({
    id: 'animSchemes', title: '幻灯片设计 - 动画方案',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', '设计模板', () => TP.show('design')), link('colorPic', '配色方案', () => TP.show('colors')), link('animation', '动画方案', () => TP.show('animSchemes'))));
      const list = h('div', { class: 'list-box trans-list', style: 'max-height:300px' });
      let pick = null;
      AS.forEach(([name, def]) => {
        if (def === 'group') { list.appendChild(h('div', { class: 'li-group', text: name })); return; }
        const li = h('div', { class: 'li', text: name });
        li.addEventListener('click', () => { L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on'); pick = def; L.app.applyAnimScheme(def, false); if (L.app.opts.autoPreview) P.playSlide(); });
        list.appendChild(li);
      });
      b.append(sec('应用于选定幻灯片：', list),
        h('div', { class: 'tp-btns' }, ui.button('应用于所有幻灯片', () => L.app.applyAnimScheme(pick, true), { class: 'btn tp-btn' })),
        h('div', { class: 'tp-btns' }, ui.button('▶ 播放', () => P.playSlide(), { class: 'btn small' }), ui.button('幻灯片放映', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
        ui.check('自动预览', !!L.app.opts.autoPreview, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
    },
  });

  /* --- Custom Animation --- */
  const CA = { selId: null };
  function effectMenu(cls, onPick) {
    return L.anim.list(cls).map(([k, n]) => ({ label: n, run: () => onPick(cls, k) }));
  }
  P.addEffectMenu = (onPick) => [
    { label: '进入(&E)', html: '<span style="color:#2f9a3a;font-size:13px">★</span>', sub: effectMenu('entr', onPick) },
    { label: '强调(&m)', html: '<span style="color:#c9a100;font-size:13px">★</span>', sub: effectMenu('emph', onPick) },
    { label: '退出(&x)', html: '<span style="color:#d33a2c;font-size:13px">★</span>', sub: effectMenu('exit', onPick) },
    { label: '动作路径(&P)', html: '<span style="color:#555;font-size:13px">☆</span>', sub: effectMenu('path', onPick) },
  ];
  TP.register({
    id: 'customAnim', title: '自定义动画',
    render(b) { renderCA(b); },
    refresh(b) { renderCA(b); },
  });
  function renderCA(b) {
    L.clear(b);
    const slide = E().slide();
    if (!slide || slide.isMaster) { b.appendChild(h('div', { class: 'tp-note', text: '自定义动画可在普通视图中使用。' })); return; }
    const anims = slide.anims || (slide.anims = []);
    const shapes = E().selected();
    const add = ui.tbDrop('添加效果', 'animation', (r) => {
      if (!E().selected().length) { ui.msg('请先选择幻灯片上的对象，再选择效果。'); return; }
      ui.openMenu(P.addEffectMenu((cls, eff) => L.app.addAnimation(cls, eff)), r);
    }, '添加效果');
    const sel = anims.find((a) => a.id === CA.selId) || null;
    const remove = ui.button('移除', () => { if (!sel) return; L.hist.push('删除效果'); slide.anims = anims.filter((a) => a.id !== sel.id); CA.selId = null; renderCA(b); L.bus.emit('slide-modified', E().idx); }, { class: 'btn small', disabled: !sel });
    b.append(h('div', { class: 'tp-row' }, add, remove));
    const upd = (k, v) => { if (!sel) return; L.hist.push('修改效果'); sel[k] = v; renderCA(b); if (L.app.opts.autoPreview) P.playSlide(sel.id); };
    const info = sel ? L.anim.info(sel.cls, sel.eff) : null;
    const start = ui.select([['click', '单击时'], ['with', '与上一动画同时'], ['after', '上一动画之后']], sel ? sel.start : 'click', (v) => upd('start', v), { disabled: !sel });
    const dir = ui.select(info && info.dirs ? info.dirs.map((d) => [d[0], d[1]]) : [['', '']], sel && sel.dir, (v) => upd('dir', v), { disabled: !(info && info.dirs) });
    const spd = ui.select(L.anim.SPEEDS.map(([n, ms]) => [ms, n]), sel ? L.anim.SPEEDS.reduce((best, s) => (Math.abs(s[1] - sel.dur) < Math.abs(best - sel.dur) ? s[1] : best), 500) : 1000, (v) => upd('dur', +v), { disabled: !sel || (sel.eff === 'appear' || sel.eff === 'disappear') });
    b.append(sec(sel ? L.t('修改：{0}', L.t(L.anim.name(sel))) : '修改效果', h('div', { class: 'tp-row' }, h('label', { text: '开始：' }), start), h('div', { class: 'tp-row' }, h('label', { text: '方向：' }), dir), h('div', { class: 'tp-row' }, h('label', { text: '速度：' }), spd)));
    const list = h('div', { class: 'list-box', style: 'min-height:150px' });
    let step = 0;
    anims.forEach((a, i) => {
      const sh = M.shapeById(slide, a.sid);
      if (a.start === 'click' || i === 0) step += a.start === 'click' ? 1 : 0;
      const label = (sh ? sh.name : '?') + (sh && sh.tx ? ': ' + L.txt.plain(sh.tx).slice(0, 24) : '');
      const col = { entr: '#2f9a3a', emph: '#c9a100', exit: '#d33a2c', path: '#555' }[a.cls];
      const li = h('div', { class: 'li' + (a.id === CA.selId ? ' on' : ''), 'data-tip': `${L.t(L.anim.name(a))} — ${label}` },
        h('span', { class: 'li-num', text: a.start === 'click' ? String(step) : '' }),
        h('span', { class: 'li-ico', html: a.start === 'click' ? '<svg width="10" height="14" viewBox="0 0 10 14"><rect x="1" y="1" width="8" height="12" rx="4" fill="#fff" stroke="#555"/><path d="M5 1v5" stroke="#555"/></svg>' : a.start === 'after' ? '<svg width="12" height="12" viewBox="0 0 12 12"><circle cx="6" cy="6" r="5" fill="#fff" stroke="#555"/><path d="M6 3v3l2 1" stroke="#555" fill="none"/></svg>' : '' }),
        h('span', { class: 'li-ico', html: `<span style="color:${col}">★</span>` }),
        h('span', { class: 'li-n', text: `${L.t(L.anim.name(a))}${a.by === 'para' ? ' ¶' : ''} — ${label}` }));
      li.addEventListener('click', () => { CA.selId = a.id; if (sh && E().topOf(a.sid)) E().select([E().topOf(a.sid).id]); renderCA(b); });
      li.addEventListener('dblclick', () => openOptions(a));
      li.addEventListener('contextmenu', (e) => {
        e.preventDefault(); CA.selId = a.id; renderCA(b);
        ui.openMenu([
          { label: '单击时开始(&C)', checked: () => a.start === 'click', run: () => upd('start', 'click') },
          { label: '与上一动画同时开始(&W)', checked: () => a.start === 'with', run: () => upd('start', 'with') },
          { label: '上一动画之后开始(&A)', checked: () => a.start === 'after', run: () => upd('start', 'after') },
          '-', { label: '效果选项(&E)...', run: () => openOptions(a) }, { label: '计时(&T)...', run: () => openOptions(a) }, '-', { label: '移除(&R)', run: () => { L.hist.push('删除效果'); slide.anims = slide.anims.filter((x) => x.id !== a.id); CA.selId = null; renderCA(b); } },
        ], { x: e.clientX, y: e.clientY });
      });
      list.appendChild(li);
    });
    if (!anims.length) list.appendChild(h('div', { class: 'tp-note', style: 'padding:8px', text: shapes.length ? '选择幻灯片中的元素，然后单击“添加效果”以添加动画。' : '选择幻灯片中的元素，然后单击“添加效果”以添加动画。' }));
    const move = (d) => { if (!sel) return; const i = anims.indexOf(sel), j = i + d; if (j < 0 || j >= anims.length) return; L.hist.push('重新排列效果'); [anims[i], anims[j]] = [anims[j], anims[i]]; renderCA(b); };
    b.append(list,
      h('div', { class: 'tp-row' }, h('span', { text: '重新排序' }), ui.button('↑', () => move(-1), { class: 'btn small', 'aria-label': '上移' }), ui.button('↓', () => move(1), { class: 'btn small', 'aria-label': '下移' })),
      h('div', { class: 'tp-btns' }, ui.button('▶ 播放', () => P.playSlide(), { class: 'btn small' }), ui.button('幻灯片放映', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
      ui.check('自动预览', !!L.app.opts.autoPreview, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
    function openOptions(a) { L.dlg.effectOptions(a, (o) => { L.hist.push('效果选项'); Object.assign(a, o); renderCA(b); if (L.app.opts.autoPreview) P.playSlide(a.id); }); }
  }
  P.selectAnim = (id) => { CA.selId = id; };

  /* --- Slide Transition --- */
  TP.register({
    id: 'transition', title: '幻灯片切换',
    render(b) { renderTR(b); },
    refresh(b, reason) { if (reason !== 'trans') renderTR(b); },
  });
  function renderTR(b) {
    L.clear(b);
    const slide = E().slide();
    if (!slide || slide.isMaster) { b.appendChild(h('div', { class: 'tp-note', text: '可在普通视图或幻灯片浏览视图中设置切换效果。' })); return; }
    const tr = slide.trans || {};
    const curName = L.anim.transName(tr);
    const list = h('div', { class: 'list-box trans-list', style: 'height:220px' });
    for (const t of L.anim.TRANSITIONS) {
      const li = h('div', { class: 'li' + (t.name === curName ? ' on' : ''), text: t.name });
      li.addEventListener('click', () => {
        L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on');
        L.app.setTransition((s) => { const keep = { spd: s.trans.spd, click: s.trans.click, after: s.trans.after }; s.trans = Object.assign({}, keep, { type: t.type, dir: t.dir, thruBlk: t.thruBlk, spokes: t.spokes, orient: t.orient }); for (const k in s.trans) if (s.trans[k] === undefined) delete s.trans[k]; });
        if (L.app.opts.autoPreview !== false) P.previewTransition();
      });
      list.appendChild(li);
    }
    setTimeout(() => { const on = list.querySelector('.li.on'); if (on) on.scrollIntoView({ block: 'nearest' }); }, 0);
    const spd = ui.select([['slow', '慢速'], ['med', '中等'], ['fast', '快速']], tr.spd || 'fast', (v) => L.app.setTransition((s) => { s.trans.spd = v; delete s.trans.dur; }));
    const snd = ui.select([['', '[无声音]']], '', null, { disabled: true });
    const onClick = ui.check('单击鼠标时', tr.click !== false, (v) => L.app.setTransition((s) => { s.trans.click = v; }));
    const secs = ui.spin({ value: tr.after != null ? tr.after / 1000 : 0, min: 0, max: 3600, step: 1, dec: 1, unit: ' s', onChange: (v) => { if (auto.input.checked) L.app.setTransition((s) => { s.trans.after = v * 1000; }); } });
    const auto = ui.check('自动换片间隔', tr.after != null, (v) => L.app.setTransition((s) => { s.trans.after = v ? secs.get() * 1000 : null; }));
    b.append(sec('应用于选定幻灯片：', list),
      sec('修改切换效果', h('div', { class: 'tp-row' }, h('label', { text: '速度：' }), spd), h('div', { class: 'tp-row' }, h('label', { text: '声音：' }), snd)),
      sec('换片方式', onClick, h('div', { class: 'tp-row' }, auto, secs)),
      h('div', { class: 'tp-btns' }, ui.button('应用于所有幻灯片', () => L.app.transitionToAll(), { class: 'btn tp-btn' })),
      h('div', { class: 'tp-btns' }, ui.button('▶ 播放', () => P.previewTransition(), { class: 'btn small' }), ui.button('幻灯片放映', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
      ui.check('自动预览', L.app.opts.autoPreview !== false, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
  }

  /* --- previews over the editor --- */
  P.previewTransition = function () {
    const ed = E();
    if (L.app.view === 'sorter') { const it = document.querySelector(`.so-it[data-i="${ed.idx}"] .so-frame`); if (it) L.show.previewTransition(it, L.pres, ed.slide(), L.pres.slides[ed.idx - 1]); return; }
    const host = ed.host;
    if (!host || !ed.slide()) return;
    const ov = h('div', { style: `position:absolute;left:${host.style.left};top:${host.style.top};width:${host.style.width};height:${host.style.height};overflow:hidden;z-index:4;pointer-events:none` });
    ed.canvas.appendChild(ov);
    L.show.previewTransition(ov, L.pres, ed.slide(), L.pres.slides[ed.idx - 1]);
    setTimeout(() => ov.remove(), L.anim.transDur(ed.slide().trans) + 400);
  };
  P.playSlide = function (onlyId) {
    const ed = E();
    const host = ed.host;
    const slide = ed.slide();
    if (!host || !slide) return;
    const ov = h('div', { style: `position:absolute;left:${host.style.left};top:${host.style.top};width:${host.style.width};height:${host.style.height};overflow:hidden;z-index:4;background:#fff` });
    ed.canvas.appendChild(ov);
    const stg = h('div', { style: `position:absolute;left:0;top:0;width:${L.pres.W}px;height:${L.pres.H}px;transform:scale(${ed.scale});transform-origin:0 0` });
    ov.appendChild(stg);
    L.show.playInline(stg, slide, { only: onlyId, done: () => ov.remove() });
    ov.addEventListener('pointerdown', () => ov.remove());
  };
})();
