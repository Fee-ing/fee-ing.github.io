/* Quire — dialogs, part 3: Format Object (pictures, AutoShapes, text boxes, WordArt), Fill Effects,
 * Tools (Word Count, AutoSummarize, Hyphenation, Envelopes and Labels, Letter Wizard, Macros, Templates,
 * AutoCorrect, Customize, Options) and File (Properties, Print, Save As, New from template, Zoom,
 * Paste Special, About).
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= Format Object ================= */
  G.formatObject = function (tab) {
    const it0 = L.drawing.selectedItem();
    if (!it0) return;
    const it = L.clone(it0);
    const isImg = it.t === 'img', isShape = it.t === 'shape', isTb = !!it.tb, isWA = !!it.wordart, isLine = isShape && (it.geom === 'line' || L.geom.isLine(it.geom));
    const title = isImg ? '设置图片格式' : isTb ? '设置文本框格式' : isWA ? '设置艺术字格式' : it.t === 'chart' ? '设置对象格式' : '设置自选图形格式';
    /* Colors and Lines */
    let fill = it.fill ? L.clone(it.fill) : { t: 'none' };
    const fillBtn = G.color(fill.t === 'solid' ? String(fill.c || '#FFFFFF').replace('#', '') : null, { id: 'fo-fill', mode: 'fill', automatic: false, noneText: fill.t === 'none' ? 'No Fill' : fill.t === 'grad' ? '渐变' : fill.t === 'patt' ? '图案' : fill.t === 'img' ? 'Picture' : 'No Fill', onChange: (v) => { fill = v ? { t: 'solid', c: '#' + v, a: 1 - transp.get() / 100 } : { t: 'none' }; } });
    const fx = ui.button('填充效果(&E)...', async () => { const f = await G.fillEffects(fill); if (f) { fill = f; fillBtn.set(null); fillBtn.querySelector('.cpk-l').textContent = { grad: '渐变', patt: '图案', img: 'Picture', solid: '颜色' }[f.t] || 'Fill'; } }, { class: 'btn small' });
    const transp = ui.spin({ value: Math.round((1 - (fill.a == null ? 1 : fill.a)) * 100), min: 0, max: 100, step: 5, dec: 0, unit: ' %', id: 'fo-tr', onChange: (v) => { if (fill.t === 'solid') fill.a = 1 - v / 100; } });
    const ln = isImg ? (it.border ? { c: '#' + (it.border.color || '000000'), w: it.border.sz || 0.75, dash: 'solid' } : { t: 'none' }) : it.line ? L.clone(it.line) : { t: 'none' };
    const lineBtn = G.color(ln.t === 'none' || !ln.c ? null : String(ln.c).replace('#', ''), { id: 'fo-line', mode: 'line', automatic: false, noneText: 'No Line' });
    const dash = ui.select(Object.keys(L.drawing.DASH).map((k) => [k, k.replace(/([A-Z])/g, ' $1').toLowerCase()]), ln.dash || 'solid', null, { id: 'fo-dash' });
    const weight = ui.spin({ value: ln.w || 0.75, min: 0, max: 1584, step: 0.25, dec: 2, unit: ' pt', id: 'fo-wt' });
    const arrows = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'].map((x) => [x, x === 'none' ? '（无）' : x]);
    const head = ui.select(arrows, (ln.head && ln.head.type) || 'none', null, { id: 'fo-ah' });
    const tail = ui.select(arrows, (ln.tail && ln.tail.type) || 'none', null, { id: 'fo-at' });
    /* Size */
    const hgt = G.len(it.h || 72, { id: 'fo-h', min: 1, onChange: (v) => { if (lock.input.checked && it.h) { wid.set((it.w * v) / it.h); scW.set(Math.round((wid.get() / (it.natW || it.w)) * 100)); } scH.set(Math.round((v / (it.natH || it.h)) * 100)); } });
    const wid = G.len(it.w || 72, { id: 'fo-w', min: 1, onChange: (v) => { if (lock.input.checked && it.w) { hgt.set((it.h * v) / it.w); scH.set(Math.round((hgt.get() / (it.natH || it.h)) * 100)); } scW.set(Math.round((v / (it.natW || it.w)) * 100)); } });
    const rot = ui.spin({ value: it.rot || 0, min: -3600, max: 3600, step: 1, dec: 1, unit: '°', id: 'fo-rot' });
    const scH = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sh', onChange: (v) => { hgt.set(((it.natH || it.h) * v) / 100); if (lock.input.checked) { wid.set(((it.natW || it.w) * v) / 100); scW.set(v); } } });
    const scW = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sw', onChange: (v) => { wid.set(((it.natW || it.w) * v) / 100); if (lock.input.checked) { hgt.set(((it.natH || it.h) * v) / 100); scH.set(v); } } });
    if (it.natW) { scW.set(Math.round((it.w / it.natW) * 100)); scH.set(Math.round((it.h / it.natH) * 100)); }
    const lock = ui.check('锁定纵横比(&A)', isImg || it.t === 'chart' || !!it.lockAspect, null);
    const reset = ui.button('重置(&S)', () => { if (it.natW) { wid.set(it.natW); hgt.set(it.natH); scW.set(100); scH.set(100); } rot.set(0); }, { class: 'btn small' });
    /* Layout */
    let wrap = it.float ? it.float.wrap || 'none' : 'inline';
    const wrapBox = h('div', { class: 'preset-grid' });
    const wrapSvg = { inline: '<svg width="40" height="36"><path d="M2 6h36M2 12h8M30 12h8M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="12" y="8" width="16" height="10" fill="#9cf" stroke="#333"/></svg>', square: '<svg width="40" height="36"><path d="M2 6h36M2 12h10M28 12h10M2 18h10M28 18h10M2 24h10M28 24h10M2 30h36" stroke="#999"/><rect x="13" y="10" width="14" height="16" fill="#9cf" stroke="#333"/></svg>', tight: '<svg width="40" height="36"><path d="M2 6h36M2 12h12M26 12h12M2 18h9M29 18h9M2 24h12M26 24h12M2 30h36" stroke="#999"/><circle cx="20" cy="18" r="8" fill="#9cf" stroke="#333"/></svg>', behind: '<svg width="40" height="36"><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#555"/></svg>', none: '<svg width="40" height="36"><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/></svg>' };
    for (const [k, l] of [['inline', '<u>嵌</u>入型'], ['square', '<u>方</u>形'], ['tight', '<u>紧</u>密'], ['behind', '<u>文</u>字下方'], ['none', '<u>浮</u>于文字上方']]) {
      const b = h('button', { type: 'button', class: wrap === k ? 'on' : '' }, h('span', { html: wrapSvg[k] }), h('span', { html: l }));
      b.addEventListener('click', () => { wrap = k; L.$$('button', wrapBox).forEach((x) => x.classList.toggle('on', x === b)); });
      wrapBox.appendChild(b);
    }
    const hAlign = G.radios('fo-ha', [['left', '&Left'], ['center', '&Center'], ['right', '&Right'], ['other', '其他(&O)']], it.float && it.float.posH && it.float.posH.align ? it.float.posH.align : 'other');
    const adv = ui.button('高级(&V)...', () => advanced(), { class: 'btn small' });
    let posH = it.float ? L.clone(it.float.posH) : { rel: 'column', off: 0 }, posV = it.float ? L.clone(it.float.posV) : { rel: 'paragraph', off: 0 };
    let dist = it.float ? L.clone(it.float.dist || { t: 0, b: 0, l: 9, r: 9 }) : { t: 0, b: 0, l: 9, r: 9 };
    let moveWith = it.float ? it.float.posV.rel === 'paragraph' : true, overlap = it.float ? it.float.allowOverlap !== false : true;
    function advanced() {
      const rels = [['column', '列'], ['margin', '页边距'], ['page', '页'], ['character', '字符']];
      const vrels = [['paragraph', 'Paragraph'], ['margin', '页边距'], ['page', '页'], ['line', 'Line']];
      const hMode = G.radios('fa-hm', [['align', '对齐(&G)'], ['abs', '绝对位置(&P)']], posH.align ? 'align' : 'abs');
      const hAl = ui.select([['left', '左'], ['center', '居中'], ['right', '右']], posH.align || 'left', null, { id: 'fa-ha' });
      const hRel = ui.select(rels, posH.rel || 'column', null, { id: 'fa-hr' });
      const hOff = G.len(posH.off || 0, { id: 'fa-ho' });
      const vMode = G.radios('fa-vm', [['align', '对齐(&N)'], ['abs', '绝对位置(&S)']], posV.align ? 'align' : 'abs');
      const vAl = ui.select([['top', '上'], ['center', '居中'], ['bottom', '底端']], posV.align || 'top', null, { id: 'fa-va' });
      const vRel = ui.select(vrels, posV.rel || 'paragraph', null, { id: 'fa-vr' });
      const vOff = G.len(posV.off || 0, { id: 'fa-vo' });
      const mv = ui.check('对象随文字移动(&M)', moveWith, null), ov = ui.check('允许重叠(&W)', overlap, null);
      const dT = G.len(dist.t || 0, { id: 'fa-dt', min: 0 }), dB = G.len(dist.b || 0, { id: 'fa-db', min: 0 }), dL = G.len(dist.l || 0, { id: 'fa-dl', min: 0 }), dR = G.len(dist.r || 0, { id: 'fa-dr', min: 0 });
      ui.dialog({ title: '高级版式', width: 460, body: ui.tabs([{ label: '图片位置(&P)', body: h('div', { class: 'col' }, ui.group('水平', hMode, G.row(hAl, h('span', { text: '相对于' }), hRel), G.row(hOff, h('span', { text: '位于右侧' }))), ui.group('垂直', vMode, G.row(vAl, h('span', { text: '相对于' }), vRel), G.row(vOff, h('span', { text: '下方' }))), ui.group('选项', mv, ov)) }, { label: '文字环绕(&T)', body: ui.group('距正文', G.row(G.f('上(&P):', dT, 'narrow'), G.f('左(&L):', dL, 'narrow')), G.row(G.f('底端(&O):', dB, 'narrow'), G.f('右(&R):', dR, 'narrow'))) }]), buttons: [{ label: '确定', primary: true, onClick: () => {
        posH = hMode.get() === 'align' ? { rel: hRel.value, align: hAl.value } : { rel: hRel.value, off: hOff.get() };
        posV = vMode.get() === 'align' ? { rel: vRel.value === 'paragraph' ? 'margin' : vRel.value, align: vAl.value } : { rel: vRel.value, off: vOff.get() };
        moveWith = mv.input.checked; overlap = ov.input.checked;
        dist = { t: dT.get(), b: dB.get(), l: dL.get(), r: dR.get() };
        hAlign.set(posH.align || 'other');
        if (wrap === 'inline') wrap = 'square';
      } }, { label: '取消' }] });
    }
    /* Picture */
    const crop = Object.assign({ l: 0, t: 0, r: 0, b: 0 }, it.crop || {});
    const cropL = G.len((crop.l || 0) * (it.natW || it.w), { id: 'fo-cl' }), cropR = G.len((crop.r || 0) * (it.natW || it.w), { id: 'fo-cr' }), cropT = G.len((crop.t || 0) * (it.natH || it.h), { id: 'fo-ct' }), cropB = G.len((crop.b || 0) * (it.natH || it.h), { id: 'fo-cb' });
    const color = ui.select([['', '自动'], ['gray', 'Grayscale'], ['bw', 'Black & White'], ['washout', 'Washout']], it.gray ? 'gray' : it.bw ? 'bw' : it.washout ? 'washout' : '', null, { id: 'fo-col' });
    const bright = ui.spin({ value: Math.round(((it.bright || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-br' });
    const contrast = ui.spin({ value: Math.round(((it.contrast || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-co' });
    /* Text Box */
    const ins = Object.assign({ l: 7.2, t: 3.6, r: 7.2, b: 3.6 }, it.ins || {});
    const iL = G.len(ins.l, { id: 'fo-il', min: 0 }), iR = G.len(ins.r, { id: 'fo-ir', min: 0 }), iT = G.len(ins.t, { id: 'fo-it', min: 0 }), iB = G.len(ins.b, { id: 'fo-ib', min: 0 });
    const anchor = ui.select([['t', '上'], ['ctr', '居中'], ['b', '底端']], it.anchor || 't', null, { id: 'fo-anc' });
    const autofit = ui.check('调整自选图形以适应文字(&F)', !!it.autofit, null);
    /* Web */
    const alt = h('textarea', { rows: 4, id: 'fo-alt', style: 'width:100%' });
    alt.value = it.alt || it.descr || '';
    const tabsDef = [
      { label: '颜色和线条', body: h('div', { class: 'col' }, isImg ? null : ui.group('填充', G.row(G.f('颜色(&O):', fillBtn), fx), G.f('透明度(&T):', transp)), ui.group('线条', G.row(G.f('颜色(&L):', lineBtn), G.f('虚线(&D):', dash)), G.f('粗细(&W):', weight)), isLine || (isShape && !isTb && !isWA) ? ui.group('箭头', G.row(G.f('起始样式(&B):', head), G.f('结束样式(&E):', tail))) : null) },
      { label: '尺寸', body: h('div', { class: 'col' }, ui.group('尺寸和旋转', G.row(G.f('高度(&E):', hgt, 'narrow'), G.f('宽度(&I):', wid, 'narrow')), G.f('旋转(&T):', rot)), ui.group('缩放', G.row(G.f('高度(&H):', scH, 'narrow'), G.f('宽度(&W):', scW, 'narrow')), lock), it.natW ? h('div', { class: 'tp-note', text: `Original size — Height: ${D.fmtLen(it.natH)}, Width: ${D.fmtLen(it.natW)}` }) : null, h('div', { class: 'tp-btns' }, reset)) },
      { label: '版式', body: h('div', { class: 'col' }, ui.group('环绕方式', wrapBox), ui.group('水平对齐', h('div', { class: 'row' }, ...hAlign.querySelectorAll('.ck'))), h('div', { class: 'tp-btns' }, adv)) },
    ];
    if (isImg) tabsDef.push({ label: '图片', body: h('div', { class: 'col' }, ui.group('裁剪', G.row(G.f('左(&L):', cropL, 'narrow'), G.f('上(&O):', cropT, 'narrow')), G.row(G.f('右(&G):', cropR, 'narrow'), G.f('底端(&B):', cropB, 'narrow'))), ui.group('图像控制', G.f('颜色(&C):', color), G.f('亮度(&H):', bright), G.f('对比度(&N):', contrast))) });
    if (isTb) tabsDef.push({ label: '文本框', body: h('div', { class: 'col' }, ui.group('内部边距', G.row(G.f('左(&L):', iL, 'narrow'), G.f('上(&T):', iT, 'narrow')), G.row(G.f('右(&R):', iR, 'narrow'), G.f('底端(&B):', iB, 'narrow'))), G.f('垂直对齐(&V):', anchor), autofit) });
    tabsDef.push({ label: 'Web', body: h('div', { class: 'col' }, h('label', { for: 'fo-alt', html: '<u>替</u>换文字:' }), alt, h('div', { class: 'tp-note', text: '图片加载中或缺失时，浏览器会显示替换文字；屏幕阅读器会朗读它。' })) });
    const tabIdx = tab === 'picture' ? tabsDef.findIndex((t) => t.label === 'Picture') : tab === 'layout' ? 2 : tab === 'size' ? 1 : 0;
    ui.dialog({
      title, width: 470, body: ui.tabs(tabsDef, Math.max(0, tabIdx)),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        L.drawing.modifySelected(title, (x) => {
          if (!isImg) { if (fill.t === 'solid') fill.a = 1 - transp.get() / 100; x.fill = fill; }
          const lc = lineBtn.get();
          if (isImg) x.border = lc ? { val: dash.value === 'solid' ? 'single' : 'dashed', sz: weight.get(), color: lc } : undefined;
          else x.line = lc ? Object.assign({}, x.line && x.line.t !== 'none' ? x.line : {}, { c: '#' + lc, w: weight.get(), dash: dash.value, head: head.value === 'none' ? undefined : { type: head.value, w: 'med', len: 'med' }, tail: tail.value === 'none' ? undefined : { type: tail.value, w: 'med', len: 'med' } }) : { t: 'none' };
          if (x.line && x.line.t !== 'none') delete x.line.t;
          if (x.t === 'group' && (wid.get() !== x.w || hgt.get() !== x.h)) { const kx = wid.get() / x.w, ky = hgt.get() / x.h; const sc = (g) => { for (const k of g.kids || []) { k.x *= kx; k.y *= ky; k.w *= kx; k.h *= ky; sc(k); } }; sc(x); }
          x.w = L.round(wid.get(), 2); x.h = L.round(hgt.get(), 2);
          x.rot = rot.get() ? ((rot.get() % 360) + 360) % 360 : undefined;
          if (lock.input.checked && !isImg) x.lockAspect = true;
          L.drawing.setWrap(x, wrap);
          if (x.float) {
            x.float.posH = hAlign.get() !== 'other' ? { rel: posH.rel === 'character' ? 'column' : posH.rel || 'column', align: hAlign.get() } : posH;
            x.float.posV = posV;
            x.float.dist = dist;
            x.float.allowOverlap = overlap;
          }
          if (isImg) {
            const nw = x.natW || it.w, nh = x.natH || it.h;
            const c = { l: (cropL.get() || 0) / nw, r: (cropR.get() || 0) / nw, t: (cropT.get() || 0) / nh, b: (cropB.get() || 0) / nh };
            x.crop = c.l || c.r || c.t || c.b ? c : undefined;
            x.gray = color.value === 'gray' || undefined; x.bw = color.value === 'bw' || undefined; x.washout = color.value === 'washout' || undefined;
            x.bright = L.round(bright.get() / 50 - 1, 2) || undefined; x.contrast = L.round(contrast.get() / 50 - 1, 2) || undefined;
          }
          if (isTb) { x.ins = { l: iL.get(), r: iR.get(), t: iT.get(), b: iB.get() }; x.anchor = anchor.value === 't' ? undefined : anchor.value; x.autofit = autofit.input.checked || undefined; }
          x.alt = alt.value || undefined;
        });
      } }, { label: '取消' }],
    });
  };
  /** Fill Effects: resolves with a fill object or null */
  G.fillEffects = function (cur) {
    return new Promise((resolve) => {
      cur = cur || {};
      let res = null;
      const c1 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[0].c).replace('#', '') : '3366FF', { id: 'fe-c1', automatic: false });
      const c2 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[cur.stops.length - 1].c).replace('#', '') : 'FFFFFF', { id: 'fe-c2', automatic: false });
      const colors = G.radios('fe-cols', [['one', '单色(&O)'], ['two', '双色(&T)'], ['preset', '预设(&E)']], 'two');
      const PRESETS = { '早霞': ['#000082', '#66008F', '#BA0066', '#FF0000', '#FF8200'], Horizon: ['#DCEBF5', '#83A7C3', '#768FB9', '#83A7C3', '#FFFFFF'], Ocean: ['#03D4A8', '#21D6E0', '#0087E6', '#005CBF'], Fire: ['#FFF200', '#FF7A00', '#FF0300', '#4D0808'], Silver: ['#FFFFFF', '#E6E6E6', '#7D8496', '#E6E6E6'], Gold: ['#E6DCAC', '#E6D78A', '#C7AC4C', '#E6D78A', '#E6DCAC'], Rainbow: ['#A603AB', '#0819FB', '#1A8D48', '#FFFF00', '#EE3F17', '#E81766'] };
      const preset = ui.select(Object.keys(PRESETS).map((k) => [k, k]), '早霞', null, { id: 'fe-pre' });
      const style = G.radios('fe-st', [['h', 'Hori&zontal'], ['v', '&Vertical'], ['du', 'Diagonal &up'], ['dd', 'Dia&gonal down'], ['corner', '&From corner'], ['center', 'From ce&nter']], 'h');
      const pattGrid = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(8,26px)' });
      let patt = cur.t === 'patt' ? cur.prst : 'pct50';
      for (const k of Object.keys(L.drawing.PATTERNS)) { const b = h('button', { type: 'button', class: k === patt ? 'on' : '', 'data-tip': k, 'aria-label': k, html: L.drawing.patternSVG(k, '#000', '#fff', 22) }); b.addEventListener('click', () => { patt = k; L.$$('button', pattGrid).forEach((x) => x.classList.toggle('on', x === b)); }); pattGrid.appendChild(b); }
      const fg = G.color('000000', { id: 'fe-fg', automatic: false }), bg = G.color('FFFFFF', { id: 'fe-bg', automatic: false });
      let pic = cur.t === 'img' ? cur.media : null;
      const picName = h('div', { class: 'tp-note', text: pic ? '已选定图片' : '未选定图片' });
      let which = cur.t === 'patt' ? 2 : cur.t === 'img' ? 3 : 0;
      const tabs = ui.tabs([
        { label: '渐变(&G)', body: h('div', { class: 'col' }, ui.group('颜色', colors, G.row(G.f('颜色 1(&1):', c1), G.f('颜色 2(&2):', c2)), G.f('预设颜色:', preset)), ui.group('底纹样式', style)), onShow: () => { which = 0; } },
        { label: '纹理', body: h('div', { class: 'tp-note', text: '纹理以图案方式绘制；请选择一种图案或图片。' }), onShow: () => { which = 1; } },
        { label: '图案(&P)', body: h('div', { class: 'col' }, pattGrid, G.row(G.f('前景(&F):', fg), G.f('背景(&B):', bg))), onShow: () => { which = 2; } },
        { label: '图片(&I)', body: h('div', { class: 'col' }, ui.button('选择图片(&P)...', async () => { const f = (await L.pickFiles('image/*'))[0]; if (!f) return; pic = L.media.add(new Blob([await L.readAsArrayBuffer(f)], { type: f.type }), f.name); picName.textContent = f.name; }), picName), onShow: () => { which = 3; } },
      ], which);
      ui.dialog({ title: '填充效果', width: 420, body: tabs, buttons: [{ label: '确定', primary: true, onClick: () => {
        if (which === 0) {
          const ang = { h: 90, v: 0, du: 45, dd: 135, corner: 45, center: 0 }[style.get()];
          let stops;
          if (colors.get() === 'preset') { const cs = PRESETS[preset.value]; stops = cs.map((c, i) => ({ p: i / (cs.length - 1), c, a: 1 })); }
          else if (colors.get() === 'one') stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#FFFFFF', a: 1 }];
          else stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#' + c2.get(), a: 1 }];
          res = { t: 'grad', ang, stops, path: style.get() === 'center' ? 'circle' : style.get() === 'corner' ? 'rect' : undefined };
        } else if (which === 2 || which === 1) res = { t: 'patt', prst: patt, fg: '#' + fg.get(), bg: '#' + bg.get() };
        else if (which === 3 && pic) res = { t: 'img', media: pic };
      } }, { label: '取消' }], onClose: () => resolve(res) });
    });
  };
  G.drawingGrid = function () {
    const snap = ui.check('对象与网格对齐(&S)', !!A().opts.snapGrid, null);
    const hs = G.len(A().opts.gridH || 9, { id: 'dg-h', min: 1 }), vs = G.len(A().opts.gridV || 9, { id: 'dg-v', min: 1 });
    ui.dialog({ title: '绘图网格', width: 300, body: h('div', { class: 'col' }, snap, ui.group('网格设置', G.f('水平间距(&H):', hs, 'wide'), G.f('垂直间距(&V):', vs, 'wide'))), buttons: [{ label: '确定', primary: true, onClick: () => { A().opts.snapGrid = snap.input.checked; A().opts.gridH = hs.get(); A().opts.gridV = vs.get(); A().saveOpts(); } }, { label: '取消' }] });
  };


  /* ================= Format ▸ AutoFormat ================= */
  G.autoFormatDoc = function () {
    const mode = G.radios('af-m', [['now', '立即自动套用格式(&A)'], ['review', '自动套用格式并逐项审阅(&R)']], 'now');
    const kind = ui.select([['general', '常规文档'], ['letter', '信函'], ['email', '电子邮件']], 'general', null, { id: 'af-k' });
    const o = A().opts;
    const opts = { headings: ui.check('内置标题样式(&S)', true, null), lists: ui.check('列表样式(&L)', true, null), quotes: ui.check('"直引号" 替换为 “弯引号”(&Q)', o.smartQuotes !== false, null), dashes: ui.check('连字符 (--) 使用破折号 (—)(&H)', o.dashes !== false, null), ordinals: ui.check('序数 (1st) 使用上标(&O)', o.ordinals !== false, null), fractions: ui.check('分数 (1/2) 使用分数字符(&F)', o.fractions !== false, null), urls: ui.check('Internet 路径使用超链接(&I)', o.autoUrl !== false, null), empty: ui.check('删除多余空段落(&E)', true, null) };
    ui.dialog({ title: '自动套用格式', width: 400, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Quire 将自动设置文档格式。' }), mode, G.f('请选择文档类型，以改进自动套用格式的效果:', kind, 'wide'), ui.group('应用 / 替换', ...Object.values(opts))), buttons: [{ label: '确定', primary: true, onClick: () => {
      const cfg = {}; for (const k in opts) cfg[k] = opts[k].input.checked;
      if (kind.value === 'letter') cfg.headings = false;
      const n = autoFormat(cfg);
      if (mode.get() === 'review') ui.dialog({ title: '自动套用格式', width: 380, body: h('div', { class: 'tp-note', text: `Formatting completed: ${n} change${n === 1 ? '' : 's'}. You can accept all of them, or reject all and return the document to its previous state.` }), buttons: [{ label: '全部接受(&A)', primary: true }, { label: '全部拒绝(&R)', onClick: () => { D.undo(); LY.update(); E.restoreDom(true); } }] });
      else A().status(`AutoFormat made ${n} change${n === 1 ? '' : 's'}.`, 6000);
    } }, { label: '取消' }] });
  };
  function autoFormat(cfg) {
    const d = doc();
    let n = 0;
    E.edit('AutoFormat', () => {
      D.reindex(d);
      const paras = d.main.paras.slice();
      const textReplace = (p, re, fn) => {
        const t = D.ptext(p);
        const hits = [];
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(t))) { hits.push([m.index, m[0], fn(m, t)]); if (!re.global) break; }
        for (let i = hits.length - 1; i >= 0; i--) {
          const [at, src, rep] = hits[i];
          if (rep == null || rep === src) continue;
          const rp = O.rPrAt(D.pos(p, at + 1));
          O.deleteRange(D.pos(p, at), D.pos(p, at + src.length));
          if (typeof rep === 'string') O.insertText(D.pos(p, at), rep, rp);
          else { O.insertText(D.pos(p, at), rep.text, Object.assign({}, rp, rep.rPr)); }
          n++;
        }
      };
      for (const p of paras) {
        if (cfg.quotes) {
          textReplace(p, /"/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-]/.test(pr) ? '“' : '”'; });
          textReplace(p, /'/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-“]/.test(pr) ? '‘' : '’'; });
        }
        if (cfg.dashes) { textReplace(p, /(\S) ?-- ?(\S)/g, (m) => m[1] + '—' + m[2]); textReplace(p, /(\w) - (\w)/g, (m) => m[1] + ' – ' + m[2]); }
        if (cfg.fractions) textReplace(p, /\b(1\/2|1\/4|3\/4)\b/g, (m) => ({ '1/2': '½', '1/4': '¼', '3/4': '¾' })[m[1]]);
        if (cfg.ordinals) {
          const t = D.ptext(p);
          const re = /\b(\d+)(st|nd|rd|th)\b/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index + m[1].length, m[2].length]);
          for (const [a, l] of hits) { O.setRunProps(D.pos(p, a), D.pos(p, a + l), { vert: 'superscript' }); n++; }
        }
        if (cfg.urls) {
          const t = D.ptext(p);
          const re = /\b((?:https?:\/\/|www\.)[^\s<>"“”]+[^\s<>"“”.,;:!?)\]])/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index, m[1]]);
          for (const [a, u] of hits) { const r = O.rPrAt(D.pos(p, a + 1)); if (r.link) continue; O.setRunProps(D.pos(p, a), D.pos(p, a + u.length), { link: { url: /^www\./.test(u) ? 'http://' + u : u }, style: 'Hyperlink' }); n++; }
        }
      }
      D.reindex(d);
      const blocks = d.main.blocks;
      const plain = (p) => D.ptext(p).replace(/[-]/g, '');
      if (cfg.lists) {
        let bulletNum = null, numNum = null, lastKind = null, touchedNum = false;
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || (p.pPr.num && p.pPr.num.id)) { lastKind = null; continue; }
          const t = plain(p);
          const mb = /^\s*([*•\-–>])\s+\S/.exec(t), mn = /^\s*(\d+|[a-zA-Z])[.)]\s+\S/.exec(t);
          if (!mb && !mn) { if (t.trim()) lastKind = null; continue; }
          const kind = mb ? 'b' : 'n';
          const cut = (mb || mn)[0].length - 1;
          O.deleteRange(D.pos(p, 0), D.pos(p, cut));
          D.touch(p);
          if (!touchedNum) { D.touchKey(d, 'numbering'); touchedNum = true; }
          if (kind === 'b') { if (!bulletNum || lastKind !== 'b') bulletNum = D.addNum(d, D.makeBulletAbs('•', 'Symbol', '')); p.pPr.num = { id: bulletNum, lvl: 0 }; }
          else { if (!numNum || lastKind !== 'n') numNum = D.addNum(d, D.makeNumberAbs(/[a-z]/.test(mn[1]) ? 'lowerLetter' : /[A-Z]/.test(mn[1]) ? 'upperLetter' : 'decimal', '%1.')); p.pPr.num = { id: numNum, lvl: 0 }; }
          if (!p.pPr.style || p.pPr.style === 'Normal') p.pPr.style = 'ListParagraph' in d.styles ? 'ListParagraph' : p.pPr.style;
          lastKind = kind;
          n++;
        }
      }
      if (cfg.headings) {
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || D.headingLevel(d, p) || (p.pPr.num && p.pPr.num.id) || (p.pPr.style && p.pPr.style !== 'Normal')) continue;
          const t = plain(p).trim();
          const next = blocks[i + 1];
          if (!t || t.length > 70 || /[.,;:!?]$/.test(t) || t.split(/\s+/).length > 10) continue;
          if (!next || next.t !== 'p' || !plain(next).trim() || plain(next).trim().length < t.length) continue;
          const prev = blocks[i - 1];
          if (prev && prev.t === 'p' && plain(prev).trim() && !D.headingLevel(d, prev)) continue;
          D.touch(p);
          p.pPr.style = /^[A-Z0-9 \-:&]+$/.test(t) || i < 3 ? 'Heading1' : 'Heading2';
          if (!d.styles[p.pPr.style]) { D.touchKey(d, 'styles'); d.styles[p.pPr.style] = D.builtinStyles()[p.pPr.style]; D.stylesChanged(); }
          n++;
        }
      }
      if (cfg.empty) {
        const cont = d.main;
        let removed = false;
        for (let i = blocks.length - 2; i > 0; i--) {
          const a = blocks[i], b = blocks[i - 1];
          if (a.t === 'p' && b.t === 'p' && !a.sect && !b.sect && !plain(a).trim() && !plain(b).trim() && a.runs.every((r) => r.t === 'text') && b.runs.every((r) => r.t === 'text')) { if (!removed) { D.touchList(cont); removed = true; } blocks.splice(i, 1); n++; }
        }
        if (removed) d._idxDirty = true;
      }
      D.reindex(d);
      return E.sel && D.byId(d, E.sel.f.p.id) ? E.sel : D.pos(D.firstPara(d.main.blocks), 0);
    });
    return n;
  }
  /* File ▸ New ▸ From existing document: open a copy that is not tied to the original file name */
  L.app.newFromExisting = async function () {
    const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm,.htm,.html,.txt,.rtf'))[0];
    if (!f) return;
    await A().openFile(f);
    const nm = 'Document' + (++A().untitled);
    A().fileName = nm;
    A().fileType = 'docx';
    doc().dirty = true;
    A().updateTitle();
  };

  /* ================= Word Count ================= */
  G.wordCount = function () {
    const d = doc();
    const selOnly = E.sel && !E.collapsed();
    const incl = ui.check('包含脚注和尾注(&F)', !!L.store.get('wcNotes', false), () => { L.store.set('wcNotes', incl.input.checked); draw(); });
    const tbl = h('table', { class: 'wc-table' });
    const draw = () => {
      L.clear(tbl);
      let paras;
      if (selOnly) { const [a, b] = E.range(); paras = D.parasBetween(d, a.p, b.p).map((p) => (p === a.p || p === b.p ? D.para(D.sliceRuns(p, p === a.p ? a.o : 0, p === b.p ? b.o : D.plen(p))) : p)); }
      else { paras = D.allParas(d); if (incl.input.checked) for (const k of ['fn', 'en']) for (const id in d[k]) D.walk(d[k][id], (b) => { if (b.t === 'p') paras.push(b); }); }
      const st = D.stats(paras);
      let lines = 0;
      if (LY.pages.length) for (const pg of LY.pages) for (const f of pg.el.querySelectorAll('.pg-body .p')) { if (f.closest('.balloons')) continue; const lh = parseFloat(getComputedStyle(f.firstChild).lineHeight) || 16; lines += Math.max(1, Math.round(f.firstChild.offsetHeight / lh)); }
      const rows = [['Pages', selOnly ? '' : LY.pages.length || 1], ['Words', st.words], ['Characters (no spaces)', st.charsNoSp], ['Characters (with spaces)', st.chars], ['Paragraphs', st.paras], ['Lines', selOnly ? '' : lines || st.paras]];
      for (const [k, v] of rows) if (v !== '') tbl.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    };
    draw();
    ui.dialog({ title: '字数统计', width: 300, body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: '统计:' + (selOnly ? ' (selection)' : '') }), tbl, incl), buttons: [{ label: '关闭', primary: true }] });
  };

  /* ================= AutoSummarize ================= */
  G.autoSummarize = function () {
    const kind = G.radios('as-k', [['highlight', '突出显示要点(&H)'], ['top', '在文档顶部插入内容提要或摘要(&I)'], ['new', '新建文档并将摘要放入其中(&C)'], ['hide', '在不离开原文档的情况下隐藏除摘要外的所有内容(&E)']], 'highlight');
    const pct = ui.select([[10, '10 sentences'], [20, '20 sentences'], [100, '100 words or less'], [500, '500 words or less'], ['10%', '10%'], ['25%', '25%'], ['50%', '50%'], ['75%', '75%']], '25%', null, { id: 'as-pct' });
    const upd = ui.check('更新文档统计信息(&U)', true, null);
    ui.dialog({ title: '自动编写摘要', width: 460, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Quire 已检查文档，并挑出了与主题最相关的句子。' }), ui.group('摘要类型', kind), ui.group('摘要长度', G.f('原尺寸的百分比(&P):', pct, 'wide')), upd), buttons: [{ label: '确定', primary: true, onClick: () => summarize(kind.get(), pct.value, upd.input.checked) }, { label: '取消' }] });
  };
  function summarize(kind, len, updStats) {
    const d = doc();
    D.reindex(d);
    const sents = [];
    const freq = new Map();
    const STOP = new Set('a an and are as at be but by for from has have he her his i in is it its of on or our she that the their them there they this to was we were what when which who will with you your not no so if than then also can may'.split(' '));
    for (const p of d.main.paras) {
      const t = D.ptext(p);
      const re = /[^.!?]+[.!?]*\s*/g;
      let m;
      while ((m = re.exec(t))) {
        const s = m[0];
        if (!s.trim()) continue;
        const words = (s.toLowerCase().match(/[\p{L}']+/gu) || []).filter((w) => !STOP.has(w) && w.length > 2);
        for (const w of words) freq.set(w, (freq.get(w) || 0) + 1);
        sents.push({ p, a: m.index, b: m.index + s.trimEnd().length, text: s.trim(), words, idx: sents.length });
      }
    }
    if (!sents.length) return;
    for (const s of sents) s.score = s.words.reduce((a, w) => a + (freq.get(w) || 0), 0) / Math.max(4, s.words.length) + (D.headingLevel(d, s.p) ? 2 : 0) + (s.idx === 0 ? 1 : 0);
    let pick;
    const totalWords = sents.reduce((a, s) => a + s.text.split(/\s+/).length, 0);
    const ranked = sents.slice().sort((a, b) => b.score - a.score);
    if (/%$/.test(len)) { const target = (totalWords * parseFloat(len)) / 100; pick = []; let w = 0; for (const s of ranked) { if (w >= target && pick.length) break; pick.push(s); w += s.text.split(/\s+/).length; } }
    else if (+len >= 100) { pick = []; let w = 0; for (const s of ranked) { const n = s.text.split(/\s+/).length; if (w + n > +len && pick.length) break; pick.push(s); w += n; } }
    else pick = ranked.slice(0, +len);
    pick.sort((a, b) => a.idx - b.idx);
    if (kind === 'highlight' || kind === 'hide') {
      E.edit('AutoSummarize', () => {
        if (kind === 'highlight') for (const s of pick) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hl: 'yellow' });
        else { const set = new Set(pick); for (const s of sents) if (!set.has(s)) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hidden: true }); }
        return E.sel;
      });
      A().status(`${pick.length} key sentence${pick.length === 1 ? '' : 's'} ${kind === 'highlight' ? 'highlighted' : 'kept visible; the rest is hidden text'}. Undo to restore.`, 8000);
    } else if (kind === 'top') {
      E.edit('AutoSummarize', () => {
        const cont = D.touchList(d.main);
        const title = D.para([D.text('Summary', { b: true })], {});
        const body = D.para([D.text(pick.map((s) => s.text).join(' '))], {});
        cont.blocks.unshift(title, body, D.para());
        d._idxDirty = true;
        return D.pos(title, 0);
      });
    } else {
      const text = pick.map((s) => s.text).join(' ');
      A().newDocument({ build: (nd) => { nd.main.blocks = [D.para([D.text('Summary of ' + A().fileName, { b: true, sz: 14 })]), D.para([D.text(text)])]; } });
    }
    if (updStats) { d.props.description = (pick.map((s) => s.text).join(' ')).slice(0, 255); }
  }

  /* ================= Hyphenation ================= */
  G.hyphenation = function () {
    const d = doc();
    const auto = ui.check('自动为文档断字(&A)', !!d.settings.autoHyphen, null);
    const caps = ui.check('大写单词断字(&C)', d.settings.hyphenCaps !== false, null);
    const zone = G.len(d.settings.hyphenZone || 18, { id: 'hy-z', min: 0 });
    const limit = G.num(d.settings.consecHyphen || null, { id: 'hy-l', min: 0, autoLabel: '无限制' });
    ui.dialog({ title: '断字', width: 320, body: h('div', { class: 'col' }, auto, caps, G.f('断字区(&Z):', zone, 'wide'), G.f('连续连字符不超过(&L):', limit, 'wide')), buttons: [{ label: '确定', primary: true, onClick: () => { E.edit('Hyphenation', () => { D.touchKey(d, 'settings'); d.settings.autoHyphen = auto.input.checked; d.settings.hyphenCaps = caps.input.checked; d.settings.hyphenZone = zone.get(); d.settings.consecHyphen = limit.get() || undefined; return E.sel; }); LY.root.classList.toggle('hyphens', auto.input.checked); } }, { label: '手动(&M)...', onClick: () => { ui.msg('手动断字：把插入点放在长单词中，按 Ctrl+- 插入可选连字符。', { icon: 'info' }); return false; } }, { label: '取消' }] });
  };

  /* ================= Envelopes and Labels ================= */
  const LABELS = [['5160', 'Avery 5160 – Address', 3, 10, 189, 72, 13.5, 36, 9], ['5161', 'Avery 5161 – Address', 2, 10, 288, 72, 11.25, 36, 13.5], ['5162', 'Avery 5162 – Address', 2, 7, 288, 96, 11.25, 63, 13.5], ['5163', 'Avery 5163 – Shipping', 2, 5, 288, 144, 11.25, 36, 13.5], ['5164', 'Avery 5164 – Shipping', 2, 3, 288, 240, 11.25, 36, 13.5], ['5167', 'Avery 5167 – Return Address', 4, 20, 126, 36, 21.6, 36, 21.6], ['5395', 'Avery 5395 – Name Badge', 2, 4, 243, 168, 49.5, 41.4, 27], ['L7160', 'Avery L7160 (A4) – Address', 3, 7, 179.15, 107.7, 20.7, 43.9, 7.2]];
  G.envelopes = function (tab) {
    const d = doc();
    let sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : '';
    const addr = h('textarea', { rows: 5, id: 'el-addr', style: 'width:100%' });
    addr.value = sel;
    const ret = h('textarea', { rows: 4, id: 'el-ret', style: 'width:100%' });
    ret.value = L.store.get('userAddress', '') || (L.store.get('userName', '') + '\n');
    const omit = ui.check('忽略(&O)', false, null);
    const size = ui.select([['10', 'Size 10 (4 1/8 x 9 1/2 in)'], ['DL', 'DL (110 x 220 mm)'], ['C5', 'C5 (162 x 229 mm)'], ['Monarch', 'Monarch (3 7/8 x 7 1/2 in)']], '10', null, { id: 'el-size' });
    const laddr = h('textarea', { rows: 5, id: 'el-laddr', style: 'width:100%' });
    laddr.value = sel;
    const lprod = ui.select(LABELS.map((x) => [x[0], x[1]]), '5160', null, { id: 'el-prod' });
    const full = G.radios('el-full', [['full', '整页同一种标签(&F)'], ['single', '单一标签(&L)']], 'full');
    const row = G.num(1, { min: 1, max: 30, id: 'el-row' }), colN = G.num(1, { min: 1, max: 4, id: 'el-col' });
    let cur = tab === 'labels' ? 1 : 0;
    const tabs = ui.tabs([
      { label: '信封(&E)', body: h('div', { class: 'col' }, h('label', { for: 'el-addr', html: '<u>收</u>件人地址:' }), addr, G.row(h('label', { for: 'el-ret', html: '<u>寄</u>件人地址:' }), omit), ret, G.f('信封尺寸(&Z):', size)), onShow: () => { cur = 0; } },
      { label: '标签(&L)', body: h('div', { class: 'col' }, h('label', { for: 'el-laddr', html: '<u>地</u>址:' }), laddr, G.f('标签产品(&P):', lprod), ui.group('打印', full, G.row(G.f('行(&T):', row, 'narrow'), G.f('列(&M):', colN, 'narrow')))), onShow: () => { cur = 1; } },
    ], cur);
    ui.dialog({ title: '信封和标签', width: 420, body: tabs, buttons: [{ label: '添加到文档(&A)', primary: true, onClick: () => { if (cur === 0) addEnvelope(); else newLabels(); } }, { label: cur === 0 ? '&Print (PDF)...' : '新建文档(&D)', onClick: () => { if (cur === 0) { addEnvelope(); setTimeout(() => A().printPDF(), 100); } else newLabels(); } }, { label: '取消' }] });
    function addEnvelope() {
      L.store.set('userAddress', ret.value);
      const sz = { 10: [684, 297], DL: [623.6, 311.8], C5: [649.15, 459.2], Monarch: [540, 279] }[size.value];
      E.edit('添加信封', () => {
        const cont = D.touchList(d.main);
        if (!d.styles.EnvelopeAddress) { D.touchKey(d, 'styles'); d.styles.EnvelopeAddress = D.builtinStyles().EnvelopeAddress; d.styles.EnvelopeReturn = D.builtinStyles().EnvelopeReturn; D.stylesChanged(); }
        const blocks = [];
        if (!omit.input.checked) for (const line of ret.value.split('\n')) blocks.push(D.para(line ? [D.text(line)] : [], { style: 'EnvelopeReturn' }));
        const lines = addr.value.split('\n');
        lines.forEach((line, i) => blocks.push(D.para(line ? [D.text(line)] : [], { style: 'EnvelopeAddress', ind: { l: sz[0] * 0.42 }, sp: i === 0 ? { b: Math.max(0, sz[1] * 0.42 - (omit.input.checked ? 0 : 60)) } : undefined })));
        const last = blocks[blocks.length - 1];
        last.sect = Object.assign(D.defaultSect(), { pgW: sz[0], pgH: sz[1], orient: 'landscape', mt: 18, mb: 18, ml: 36, mr: 36, type: 'nextPage' });
        cont.blocks.unshift(...blocks);
        d._idxDirty = true;
        return D.pos(blocks[0], 0);
      });
    }
    function newLabels() {
      const spec = LABELS.find((x) => x[0] === lprod.value);
      const [, , nc, nr, lw, lh, ml, mt, gap] = spec;
      const text = laddr.value;
      A().newDocument({ build: (nd) => {
        nd.sect.ml = ml; nd.sect.mr = 0; nd.sect.mt = mt; nd.sect.mb = 0; nd.sect.hdr = 0; nd.sect.ftr = 0;
        if (/A4/.test(spec[1])) { nd.sect.pgW = 595.3; nd.sect.pgH = 841.9; }
        const cols = [];
        for (let c = 0; c < nc; c++) { cols.push(lw); if (c < nc - 1 && gap) cols.push(gap); }
        const t = D.table([], cols, { style: 'TableNormal', w: { type: 'dxa', v: cols.reduce((a, b) => a + b, 0) }, layout: 'fixed', cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 }, look: {} });
        for (let r = 0; r < nr; r++) {
          const cells = [];
          for (let c = 0; c < cols.length; c++) {
            const isGap = gap && c % 2 === 1;
            const lr = Math.floor(c / (gap ? 2 : 1));
            const fill = !isGap && (full.get() === 'full' || (r === row.get() - 1 && lr === colN.get() - 1));
            const blocks = fill ? text.split('\n').map((ln) => D.para(ln ? [D.text(ln)] : [], { ind: { l: 5 }, sp: { b: 0, a: 0 } })) : [D.para()];
            cells.push(D.cell(blocks, { w: cols[c], vAlign: 'center' }));
          }
          t.rows.push(D.row(cells, { h: lh, hRule: 'exact', cantSplit: true }));
        }
        nd.main.blocks = [t, D.para()];
      } });
    }
  };

  /* ================= Letter Wizard ================= */
  G.letterWizard = function () {
    const date = ui.check('日期行(&D):', true, null);
    const datePic = ui.select(L.fields.DATE_PICTURES.slice(0, 6).map((p) => [p, L.fields.formatDate(new Date(), p)]), 'MMMM d, yyyy', null, { id: 'lw-date' });
    const style = ui.select([['full', '齐头式'], ['modified', '改良齐头式'], ['semi', '半齐头式']], 'full', null, { id: 'lw-style' });
    const rname = h('input', { type: 'text', id: 'lw-rn', style: 'width:100%' });
    const raddr = h('textarea', { rows: 4, id: 'lw-ra', style: 'width:100%' });
    const sal = h('input', { type: 'text', id: 'lw-sal', value: 'Dear ', list: 'lw-sals', style: 'width:100%' });
    const sals = h('datalist', { id: 'lw-sals' }, ...['Dear ', 'Dear Sir or Madam:', 'Ladies and Gentlemen:', 'To Whom It May Concern:'].map((x) => h('option', { value: x })));
    const refLine = h('input', { type: 'text', id: 'lw-ref', style: 'width:100%', placeholder: '回复给:' });
    const subj = h('input', { type: 'text', id: 'lw-sub', style: 'width:100%' });
    const cc = h('input', { type: 'text', id: 'lw-cc', style: 'width:100%' });
    const sname = h('input', { type: 'text', id: 'lw-sn', value: L.store.get('userName', ''), style: 'width:100%' });
    const saddr = h('textarea', { rows: 3, id: 'lw-sa', style: 'width:100%' });
    saddr.value = L.store.get('userAddress', '');
    const closing = h('input', { type: 'text', id: 'lw-cl', value: 'Sincerely,', list: 'lw-cls', style: 'width:100%' });
    const cls = h('datalist', { id: 'lw-cls' }, ...L.autocorrect.AUTOTEXT.Closing.map((x) => h('option', { value: x })));
    const title = h('input', { type: 'text', id: 'lw-ti', style: 'width:100%' });
    const encl = G.num(0, { id: 'lw-enc', min: 0 });
    const tabs = ui.tabs([
      { label: '信函格式', body: h('div', { class: 'col' }, G.row(date, datePic), G.f('页面设计 / 信函样式:', style, 'wide')) },
      { label: '收件人信息', body: h('div', { class: 'col' }, G.f('收件人姓名:', rname, 'wide'), h('label', { for: 'lw-ra', text: '收件人地址:' }), raddr, sals, G.f('称呼:', sal, 'wide')) },
      { label: '其他元素', body: h('div', { class: 'col' }, G.f('引用线:', refLine, 'wide'), G.f('Subject:', subj, 'wide'), G.f('抄送(cc):', cc, 'wide')) },
      { label: '发件人信息', body: h('div', { class: 'col' }, G.f('发件人姓名:', sname, 'wide'), h('label', { for: 'lw-sa', text: '寄件人地址:' }), saddr, cls, G.f('结束语:', closing, 'wide'), G.f('职务:', title, 'wide'), G.f('附件:', encl, 'wide')) },
    ]);
    ui.dialog({ title: '英文信函向导', width: 460, body: tabs, buttons: [{ label: '确定', primary: true, onClick: () => {
      L.store.set('userAddress', saddr.value);
      const indent = style.value === 'full' ? 0 : 252;
      const P = (t, o) => D.para(t ? [D.text(t)] : [], o || {});
      const bl = [];
      for (const line of saddr.value.split('\n').filter(Boolean)) bl.push(P(line, { ind: { l: indent } }));
      if (bl.length) bl.push(P(''));
      if (date.input.checked) bl.push(P(L.fields.formatDate(new Date(), datePic.value), { ind: { l: indent } }), P(''));
      if (refLine.value) bl.push(P(refLine.value), P(''));
      if (rname.value) bl.push(P(rname.value));
      for (const line of raddr.value.split('\n').filter(Boolean)) bl.push(P(line));
      bl.push(P(''));
      if (subj.value) bl.push(D.para([D.text('Subject: ', { b: true }), D.text(subj.value)]), P(''));
      bl.push(P(sal.value.trim() === '尊敬的' || sal.value === 'Dear ' ? 'Dear ' + (rname.value || '') + ':' : sal.value), P(''));
      bl.push(P('在此输入信函内容。', { ind: { fl: style.value === 'semi' ? 36 : 0 } }), P(''));
      bl.push(P(closing.value, { ind: { l: indent } }), P('', { ind: { l: indent } }), P('', { ind: { l: indent } }), P(sname.value, { ind: { l: indent } }));
      if (title.value) bl.push(P(title.value, { ind: { l: indent } }));
      if (encl.get()) bl.push(P(''), P(`Enclosure${encl.get() > 1 ? 's' : ''} (${encl.get()})`));
      if (cc.value) bl.push(P(''), P('cc: ' + cc.value));
      E.pasteBlocks(bl.map((b) => (delete b.partial, b)));
    } }, { label: '取消' }] });
  };

  /* ================= Macros, Templates & Organizer ================= */
  G.macros = function () {
    const lb = G.listBox([], null, { height: 140 });
    ui.dialog({ title: '宏', width: 360, body: h('div', { class: 'col' }, h('label', { text: '宏名:' }), h('input', { type: 'text', style: 'width:100%', disabled: true }), lb, h('div', { class: 'tp-note', text: 'Macros are not run in Quire. VBA projects in .docm files are kept when the document is saved as .docm by Word, but Quire saves the document text only.' })), buttons: [{ label: '关闭', primary: true }] });
  };
  G.templates = function () {
    const d = doc();
    const auto = ui.check('自动更新文档样式(&U)', !!d.settings.attachedTemplateUpdate, null);
    const name = h('input', { type: 'text', value: d.props.template || 'Normal', style: 'width:100%', readonly: true });
    ui.dialog({ title: '模板和加载项', width: 400, body: h('div', { class: 'col' }, ui.group('文档模板', name, auto, h('div', { class: 'tp-btns' }, ui.button('附加(&A)...', async () => { const f = (await L.pickFiles('.dotx,.dotm,.docx'))[0]; if (!f) return; const res = await L.docx.read(await L.readAsArrayBuffer(f)); d.props.template = f.name; name.value = f.name; if (auto.input.checked) A().mergeStylesFrom(res.doc, true); A().status('Template attached: ' + f.name); }, { class: 'btn small' }), ui.button('管理器(&O)...', () => G.organizer(), { class: 'btn small' }))), h('div', { class: 'tp-note', text: 'Attaching a template with "自动更新文档样式" copies its styles into this document.' })), buttons: [{ label: '确定', primary: true, onClick: () => { d.settings.attachedTemplateUpdate = auto.input.checked || undefined; } }, { label: '取消' }] });
  };
  G.organizer = function () {
    const d = doc();
    let other = null, otherName = '';
    const mine = h('select', { size: 12, multiple: true, style: 'width:100%' });
    const theirs = h('select', { size: 12, multiple: true, style: 'width:100%' });
    const fillSel = (sel, dd) => { L.clear(sel); if (!dd) return; for (const k of Object.keys(dd.styles).sort()) sel.appendChild(h('option', { value: k, text: D.styleDisplayName(dd.styles[k]) })); };
    fillSel(mine, d);
    const lblOther = h('div', { class: 'cd-cap', text: '内嵌（无文件）:' });
    ui.dialog({ title: '管理器', width: 560, body: G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'In ' + A().fileName + ':' }), mine), h('div', { class: 'col', style: 'width:110px;justify-content:center' }, ui.button('<- &Copy', () => { if (!other) return; for (const o of theirs.selectedOptions) { E.edit('Copy Style', () => { D.touchKey(d, 'styles'); d.styles[o.value] = L.clone(other.styles[o.value]); D.stylesChanged(); return E.sel; }); } fillSel(mine, d); }), ui.button('删除(&D)', () => { E.edit('Delete Styles', () => { D.touchKey(d, 'styles'); for (const o of mine.selectedOptions) if (!d.styles[o.value].builtin && o.value !== 'Normal') delete d.styles[o.value]; D.stylesChanged(); return E.sel; }); fillSel(mine, d); }), ui.button('重命名(&R)...', async () => { const o = mine.selectedOptions[0]; if (!o) return; const n = await ui.prompt('New name:', D.styleDisplayName(d.styles[o.value]), 'Rename'); if (!n) return; E.edit('Rename Style', () => { D.touchKey(d, 'styles'); d.styles[o.value].name = n; return E.sel; }); fillSel(mine, d); })), h('div', { class: 'col', style: 'flex:1' }, lblOther, theirs, ui.button('打开文件(&F)...', async () => { const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm'))[0]; if (!f) return; other = (await L.docx.read(await L.readAsArrayBuffer(f))).doc; otherName = f.name; lblOther.textContent = 'In ' + otherName + ':'; fillSel(theirs, other); }))), buttons: [{ label: '关闭', primary: true }] });
  };

  /* ================= AutoCorrect ================= */
  G.autocorrect = function (tab) {
    const o = A().opts;
    const ac = L.autocorrect;
    const chk = (k, l, def) => ui.check(l, o[k] !== undefined ? o[k] !== false : def !== false, (v) => { o[k] = v; A().saveOpts(); });
    const from = h('input', { type: 'text', id: 'ac-from', style: 'width:110px' }), to = h('input', { type: 'text', id: 'ac-to', style: 'flex:1;min-width:0' });
    const list = G.listBox([], null, { height: 150, onChange: (v) => { from.value = v; to.value = ac.list().get(v) || ''; } });
    const fill = () => { const m = ac.list(); list.setItems(Array.from(m.keys()).sort((a, b) => a.localeCompare(b)).map((k) => ({ value: k, label: `${k}  →  ${m.get(k)}` }))); };
    fill();
    from.addEventListener('input', () => { const m = ac.list(); if (m.has(from.value)) to.value = m.get(from.value); });
    const exc = () => { const ta = h('textarea', { rows: 10, style: 'width:100%' }); ta.value = ac.exceptions().join('\n'); ui.dialog({ title: '自动更正例外项', width: 300, body: h('div', { class: 'col' }, h('div', { text: "Don't capitalize after:" }), ta), buttons: [{ label: '确定', primary: true, onClick: () => ac.setExceptions(ta.value.split(/\s+/).filter(Boolean)) }, { label: '取消' }] }); };
    const atName = h('input', { type: 'text', id: 'at-name', style: 'width:100%' });
    const atList = G.listBox(L.store.get('autoText', []).map((x) => x.name), null, { height: 120, onChange: (v) => { const e = L.store.get('autoText', []).find((x) => x.name === v); atName.value = v; atPv.textContent = e ? e.text : ''; } });
    const atPv = h('div', { class: 'tp-box', style: 'min-height:40px;white-space:pre-wrap' });
    const tabs = ui.tabs([
      { label: '自动更正(&A)', body: h('div', { class: 'col' }, chk('acButtons', '显示自动更正选项按钮(&C)'), chk('twoCaps', '更正两个连续的大写字母(&C)'), chk('capSentence', '句首字母大写(&S)'), chk('capCells', '表格单元格首字母大写(&C)'), chk('capDays', '日期名称首字母大写(&N)'), chk('capsLock', '更正误按 Caps Lock 键造成的大小写错误(&L)'), ui.button('例外项(&E)...', exc, { class: 'btn small' }), chk('autocorrect', '键入时替换文字(&T)'), G.row(G.f('替换(&R):', from), G.f('与(&W):', to)), list, h('div', { class: 'tp-btns' }, ui.button('添加(&A)', () => { if (!from.value) return; ac.setEntry(from.value, to.value); fill(); }, { class: 'btn small' }), ui.button('删除(&D)', () => { if (!from.value) return; ac.setEntry(from.value, null); fill(); from.value = ''; to.value = ''; }, { class: 'btn small' }))) },
      { label: '键入时自动套用格式(&Y)', body: h('div', { class: 'col' }, ui.group('键入时替换', chk('smartQuotes', '"直引号" 替换为 “弯引号”(&S)'), chk('ordinals', '序数 (1st) 使用上标(&O)'), chk('fractions', '分数 (1/2) 使用分数字符 (½)(&F)'), chk('dashes', '连字符 (--) 使用破折号 (—)(&H)'), chk('autoFormat', '*加粗* 和 _倾斜_ 使用实际格式(&B)'), chk('autoUrl', 'Internet 和网络路径使用超链接(&I)')), ui.group('键入时应用', chk('autoBullets', '自动项目符号列表(&U)'), chk('autoBorders', '边框线(&L)'), chk('autoTables', '表格(&T)'), chk('autoNumbers', '自动编号列表(&N)'), chk('autoHeadings', '内置标题样式(&Y)', false))) },
      { label: '自动图文集(&T)', body: h('div', { class: 'col' }, chk('autoComplete', '显示自动完成建议(&A)'), h('label', { for: 'at-name', text: '在此输入自动图文集词条:' }), atName, atList, h('div', { class: 'cd-cap', text: '预览' }), atPv, h('div', { class: 'tp-btns' }, ui.button('添加(&A)', () => { const t = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : ''; if (!t) { ui.msg('请先选定要存为自动图文集的文本。', { icon: 'info' }); return; } const n = atName.value.trim() || t.slice(0, 32); ac.addAutoText(n, t); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); }, { class: 'btn small' }), ui.button('删除(&D)', () => { ac.deleteAutoText(atName.value); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); atPv.textContent = ''; }, { class: 'btn small' }), ui.button('插入(&I)', () => { const e = L.store.get('autoText', []).find((x) => x.name === atName.value); if (e) E.insertTextAt(e.text, 'AutoText'); }, { class: 'btn small' }))) },
    ], tab === 'autotext' ? 2 : tab === 'format' ? 1 : 0);
    ui.dialog({ title: '自动更正：英语（美国）', width: 460, body: tabs, buttons: [{ label: '确定', primary: true }, { label: '取消' }] });
  };

  /* ================= Customize ================= */
  G.customize = function () {
    const tbs = ['standard', 'formatting', 'drawing', 'tables', 'reviewing', 'picture', 'outlining', 'headerFooter', 'mailMerge', 'wordart'].map((k) => ui.check(ui.cmds['tb_' + k].label, A().tbShown(k), () => A().toggleToolbar(k)));
    const cmds = Object.values(ui.cmds).filter((c) => c.label && !c.id.startsWith('tb_')).sort((a, b) => ui.stripAmp(a.label).localeCompare(ui.stripAmp(b.label)));
    const cmdList = G.listBox(cmds.map((c) => ({ value: c.id, label: ui.stripAmp(c.menuLabel || c.label) + (c.key ? '   ' + c.key : '') })), null, { height: 220 });
    const keys = h('table', { class: 'kbd-table' });
    for (const [k, id] of Object.entries(A().SHORTCUTS).sort((a, b) => a[1].localeCompare(b[1]))) { const c = ui.cmds[id]; if (c) keys.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: ui.stripAmp(c.menuLabel || c.label) }))); }
    const o = A().opts;
    ui.dialog({ title: '自定义', width: 460, body: ui.tabs([
      { label: '工具栏(&B)', body: h('div', { class: 'col' }, ...tbs) },
      { label: '命令(&C)', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Quire 的全部命令。双击一项即可运行。' }), cmdList), onShow: () => {} },
      { label: '选项(&O)', body: h('div', { class: 'col' }, ui.check('分两行显示常用和格式工具栏', true, null), ui.check('始终显示完整菜单', true, null), ui.check('在工具栏上显示屏幕提示(&N)', o.screenTips !== false, (v) => { o.screenTips = v; A().saveOpts(); document.body.classList.toggle('notips', !v); }), ui.check('在屏幕提示中显示快捷键(&K)', true, null)) },
      { label: '键盘(&K)', body: h('div', { style: 'max-height:300px;overflow:auto' }, keys) },
    ]), buttons: [{ label: '关闭', primary: true }] });
    cmdList.querySelector('.listbox').addEventListener('dblclick', () => { const id = cmdList.get(); if (id) ui.exec(id); });
  };

  /* ================= Options ================= */
  G.options = function (tab) {
    const o = A().opts;
    const d = doc();
    const pwInp = h('input', { type: 'password', id: 'op-pw', value: d.openPassword || '', style: 'width:220px', autocomplete: 'new-password' });
    const chk = (label, get, set) => ui.check(label, get(), set);
    const opt = (k, l, def) => chk(l, () => (o[k] !== undefined ? !!o[k] : !!def), (v) => { o[k] = v; });
    const lyo = (k, l) => chk(l, () => !!LY.opts[k], (v) => { LY.opts[k] = v; rerender = true; });
    let rerender = false;
    const fieldShade = ui.select([['never', '从不'], ['always', '始终'], ['selected', '选定内容时']], o.fieldShading || 'selected', (v) => { o.fieldShading = v; }, { id: 'op-fs' });
    const units = ui.select([['in', '英寸'], ['cm', '厘米'], ['mm', '毫米'], ['pt', '磅'], ['pi', '皮卡']], o.unit || 'in', (v) => { o.unit = v; }, { id: 'op-unit' });
    const recentN = G.num(4, { min: 0, max: 9, id: 'op-rec' });
    const uname = h('input', { type: 'text', id: 'op-un', value: o.userName || '', style: 'width:100%' });
    const uinit = h('input', { type: 'text', id: 'op-ui', value: o.userInitials || '', style: 'width:80px' });
    const uaddr = h('textarea', { rows: 3, id: 'op-ua', style: 'width:100%' });
    uaddr.value = L.store.get('userAddress', '');
    const insColor = ui.select([['author', '按作者'], ['red', '红色'], ['blue', '蓝色'], ['green', '绿色']], o.insColor || 'author', (v) => { o.insColor = v; }, { id: 'op-ic' });
    const balloons = ui.select([['always', '始终'], ['never', '从不'], ['comments', '仅用于批注/格式']], LY.opts.balloons === false ? 'never' : 'comments', (v) => { LY.opts.balloons = v !== 'never'; rerender = true; }, { id: 'op-bal' });
    const compat = ui.select([['11', 'Microsoft Word 2003'], ['12', 'Microsoft Word 2007'], ['14', 'Microsoft Word 2010'], ['15', 'Microsoft Word 2013 及更高版本']], String(d.settings.compat || 11), (v) => { d.settings.compat = +v; d.dirty = true; rerender = true; }, { id: 'op-compat' });
    const autoRec = G.num(o.autoSaveMin || 10, { min: 1, max: 120, id: 'op-asr' });
    const look = ui.lookField();
    const tabs = ui.tabs([
      { label: '视图', body: h('div', { class: 'col' }, ui.group('显示', h('div', { class: 'grid2' }, opt('startupPane', '启动任务窗格', true), opt('paneFollows', '任务窗格随选定对象变化', true), opt('highlightShow', '突出显示', true), lyo('showBookmarks', '书签'), opt('statusBar', '状态栏', true), opt('screenTips', '屏幕提示', true), lyo('fieldCodes', '域代码'), opt('smartTags', '智能标记', false), opt('animText', '动画文字', false)), G.f('域底纹:', fieldShade)), ui.group('格式标记', h('div', { class: 'grid2' }, lyo('showMarks', '全部'), lyo('showHidden', '隐藏文字'))), ui.group('打印与 Web 版式选项', h('div', { class: 'grid2' }, lyo('textBoundaries', '文字边界'), opt('vruler', '垂直标尺（仅页面视图）', true), opt('whiteSpace', '页面间的空白（仅页面视图）', true)))) },
      { label: '常规', body: h('div', { class: 'col' }, look && look.el, opt('bgRepag', '后台重新分页', true), opt('blueBg', '蓝底白字', false), opt('sound', '使用声音提供反馈', false), G.row(opt('recentOn', '最近使用的文件列表:', true), recentN, h('span', { text: 'entries' })), G.f('度量单位:', units, 'wide')) },
      { label: '编辑', body: h('div', { class: 'col' }, ui.group('编辑选项', opt('typingReplaces', '输入内容替换选定内容', true), opt('dragDrop', '拖放式文字编辑', true), opt('ctrlClick', '按住 Ctrl 并单击可访问超链接', true), opt('smartPara', '使用智能段落选择', true), chk('改写模式', () => E.overtype, (v) => { E.overtype = v; }), opt('smartCut', '智能剪切和粘贴', true)), ui.group('剪切和粘贴选项', opt('pasteOptions', '显示粘贴选项按钮', true))) },
      { label: '打印', body: h('div', { class: 'col' }, ui.group('打印选项', opt('updateFieldsOnPrint', '更新域', false), opt('printBg', '背景颜色和图像', true)), ui.group('包含在文档中', opt('printProps', '文档属性', false), opt('printCodes', '域代码', false), opt('printHidden', '隐藏文字', false), opt('printDrawings', '图形对象', true))) },
      { label: '保存', body: h('div', { class: 'col' }, opt('promptProps', '提示输入文档属性', false), G.row(opt('autoRecover', '自动恢复信息保存间隔:', true), autoRec, h('span', { text: 'minutes' })), h('div', { class: 'tp-note', text: '自动恢复只在本浏览器里保留副本；Quire 在意外关闭后重新打开时会提示恢复。' }), G.f('Word 文件保存为:', ui.select([['docx', 'Word Document (*.docx)'], ['dotx', 'Word 模板 (*.dotx)']], 'docx', null, { id: 'op-fmt' }), 'wide')) },
      { label: '安全性', body: h('div', { class: 'col' }, ui.group('隐私选项', opt('removePersonal', '保存时从文件属性中删除个人信息', false), opt('warnTracked', '打印、保存或发送含修订或批注的文件前给出警告', false)), ui.group('此文档的文件加密选项', h('div', { class: 'col' }, h('label', { for: 'op-pw', text: '打开权限密码:' }), pwInp, h('div', { class: 'tp-note', text: 'Saved .docx files are encrypted with AES-256, as Word 2010 and later do. Passwords are case-sensitive; if you forget it, the document cannot be recovered.' }))), h('div', { class: 'tp-note', text: '若要改为限制编辑，请使用「工具 ▸ 保护文档」。' })) },
      { label: '修订', body: h('div', { class: 'col' }, ui.group('标记', G.f('插入内容颜色:', insColor, 'wide')), ui.group('批注框', G.f('使用批注框（页面视图和 Web 版式）:', balloons, 'wide'))) },
      { label: '用户信息', body: h('div', { class: 'col' }, G.f('名称:', uname, 'wide'), G.f('缩写:', uinit, 'wide'), h('label', { for: 'op-ua', text: '邮件地址:' }), uaddr) },
      { label: '兼容性', body: h('div', { class: 'col' }, G.f('按以下版本创建时的方式排版此文档:', compat, 'wide'), h('div', { class: 'tp-note', text: 'Word 2003 layout positions tables by their cell margins and keeps legacy spacing rules; later versions align tables with the text.' })) },
    ], { view: 0, general: 1, edit: 2, print: 3, save: 4, security: 5, spelling: 6, track: 7, user: 8, compat: 9 }[tab] || 0);
    tabs.classList.add('options-tabs');
    ui.dialog({ title: '选项', width: 560, body: tabs, buttons: [{ label: '确定', primary: true, onClick: () => {
      const newPw = pwInp.value;
      if (newPw !== (d.openPassword || '')) {
        if (!newPw) { d.openPassword = null; A().status('The password to open was removed; the next save is not encrypted.'); }
        else {
          /* Word asks for the password a second time before using it */
          const again = h('input', { type: 'password', id: 'op-pw2', style: 'width:100%' });
          let val = null;
          ui.dialog({ title: '确认密码', width: 340, body: h('div', { class: 'col' }, h('label', { for: 'op-pw2', text: '请再次输入打开权限密码:' }), again, h('div', { class: 'tp-note', text: '注意: 密码一旦丢失或忘记将无法恢复。' })), buttons: [{ label: '确定', primary: true, onClick: () => { val = again.value; } }, { label: '取消' }] }).done.then(() => {
            if (val == null) return;
            if (val !== newPw) { ui.msg('两次输入的密码不一致。', { icon: 'warn' }); return; }
            d.openPassword = newPw;
            A().dirty && A().dirty();
            A().status('保存文档时将用此密码加密。');
          });
          setTimeout(() => again.focus(), 0);
        }
      }
      o.userName = uname.value.trim() || 'Quire User';
      o.userInitials = uinit.value.trim();
      o.autoSaveMin = autoRec.get();
      L.store.set('userAddress', uaddr.value);
      A().saveOpts();
      A().applyOpts();
      if (look) look.apply();
      if (rerender) { LY.root.classList.toggle('marks', !!LY.opts.showMarks); LY.render(); E.restoreDom(true); }
    } }, { label: '取消' }] });
  };

  /* ================= Properties ================= */
  G.properties = function () {
    const d = doc();
    const p = d.props;
    const st = D.stats(D.allParas(d));
    const fld = (k, l) => { const inp = h('input', { type: 'text', id: 'pr-' + k, value: p[k] || '', style: 'width:100%' }); return { k, inp, el: G.f(l, inp, 'wide') }; };
    const fields = [fld('title', '标题(&T):'), fld('subject', '主题(&S):'), fld('creator', '作者(&A):'), fld('manager', '经理(&M):'), fld('company', '单位(&O):'), fld('category', '类别(&E):'), fld('keywords', '关键词(&K):')];
    const comments = h('textarea', { rows: 3, id: 'pr-desc', style: 'width:100%' });
    comments.value = p.description || '';
    const hb = fld('hyperlinkBase', '超链接基础(&H):');
    const fmt = (s) => (s ? new Date(s).toLocaleString() : '');
    const statT = h('table', { class: 'wc-table' });
    for (const [k, v] of [['Pages:', LY.pages.length || 1], ['段落:', st.paras], ['行:', A().docStats().lines], ['字数:', st.words], ['字符数:', st.charsNoSp], ['字符数（含空格）:', st.chars]]) statT.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    const heads = h('div', { class: 'tp-box', style: 'max-height:200px;overflow:auto' });
    D.reindex(d);
    for (const q of d.main.paras) { const lv = D.headingLevel(d, q); if (lv) heads.appendChild(h('div', { text: D.plainText(q).trim(), style: `padding-left:${(lv - 1) * 12}px` })); }
    if (!heads.childNodes.length) heads.textContent = '(no headings)';
    /* custom */
    const cName = h('input', { type: 'text', id: 'pr-cn', list: 'pr-cnl', style: 'width:100%' });
    const cnl = h('datalist', { id: 'pr-cnl' }, ...['审核人', '客户', '完成日期', '部门', '目标位置', '处理意见', '分隔符', '文档编号', '编辑者', '转发给', '组合', 'Language', 'Mailstop', '事由', 'Office', 'Owner', '项目', '发布者', '用途', '接收自', '录制者', '录制日期', '引用', '来源', '状态', '电话号码', '打字员'].map((x) => h('option', { value: x })));
    const cType = ui.select([['text', '文字'], ['date', 'Date'], ['number', '数字'], ['bool', '是或否']], 'text', null, { id: 'pr-ct' });
    const cVal = h('input', { type: 'text', id: 'pr-cv', style: 'width:100%' });
    const custom = Object.assign({}, d.custom || {});
    const cList = G.listBox([], null, { height: 110, onChange: (v) => { const c = custom[v]; cName.value = v; cVal.value = c && c.v != null ? c.v : c; cType.value = (c && c.type) || 'text'; } });
    const fillC = () => cList.setItems(Object.keys(custom).map((k) => ({ value: k, label: `${k}: ${custom[k] && custom[k].v != null ? custom[k].v : custom[k]}` })));
    fillC();
    const tabs = ui.tabs([
      { label: '常规', body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center' }, h('span', { html: L.icons.get('new') }), h('b', { text: A().fileName + '.docx' })), h('table', { class: 'wc-table' }, ...[['Type:', 'Microsoft Word Document'], ['创建时间:', fmt(p.created)], ['修改时间:', fmt(p.modified)], ['Last printed:', fmt(p.printed)]].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: b }))))) },
      { label: '摘要', body: h('div', { class: 'col' }, ...fields.map((f) => f.el), h('label', { for: 'pr-desc', html: '<u>批</u>注:' }), comments, hb.el, h('div', { class: 'tp-note', text: 'Template: ' + (p.template || 'Normal') })) },
      { label: '统计信息', body: h('div', { class: 'col' }, h('table', { class: 'wc-table' }, ...[['创建时间:', fmt(p.created)], ['修改时间:', fmt(p.modified)], ['上次保存者:', p.lastModifiedBy || ''], ['Revision number:', String(p.revision || 1)], ['Total editing time:', (p.totalTime || 0) + ' Minutes']].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: b })))), h('div', { class: 'cd-cap', text: '统计:' }), statT) },
      { label: '内容', body: h('div', { class: 'col' }, h('div', { text: '文档内容:' }), heads) },
      { label: '自定义', body: h('div', { class: 'col' }, cnl, G.f('名称(&N):', cName, 'wide'), G.f('类型(&T):', cType, 'wide'), G.f('值(&V):', cVal, 'wide'), h('div', { class: 'tp-btns' }, ui.button('添加(&A)', () => { if (!cName.value.trim()) return; custom[cName.value.trim()] = { type: cType.value, v: cType.value === 'number' ? +cVal.value : cType.value === 'bool' ? /^(y|yes|true|1)$/i.test(cVal.value) : cVal.value }; fillC(); }, { class: 'btn small' }), ui.button('删除(&D)', () => { delete custom[cName.value.trim()]; fillC(); }, { class: 'btn small' })), h('label', { text: '属性:' }), cList) },
    ], 1);
    ui.dialog({ title: A().fileName + ' Properties', width: 420, body: tabs, buttons: [{ label: '确定', primary: true, onClick: () => {
      E.edit('Properties', () => { D.touchKey(d, 'props'); D.touchKey(d, 'custom'); for (const f of fields) d.props[f.k] = f.inp.value; d.props.description = comments.value; d.props.hyperlinkBase = hb.inp.value || undefined; d.custom = custom; return E.sel; });
      A().updateTitle();
    } }, { label: '取消' }] });
  };

  /* ================= Zoom ================= */
  G.zoom = function () {
    const cur = Math.round(LY.zoom * 100);
    const r = G.radios('zm', [['200', '&200%'], ['100', '&100%'], ['75', '&75%'], ['pw', '页宽(&P)'], ['tw', '文字宽度(&T)'], ['wp', '整页(&W)'], ['many', '多页(&M)']], [200, 100, 75].includes(cur) ? String(cur) : 'pct', (v) => { const n = { 200: 200, 100: 100, 75: 75 }[v]; if (n) pct.set(n); });
    const pct = ui.spin({ value: cur, min: 10, max: 500, step: 10, dec: 0, unit: '%', id: 'zm-pct' });
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:110px', text: 'AaBbCcDdEeXxYyZz' });
    ui.dialog({ title: '显示比例', width: 360, body: G.row(ui.group('缩放为', r), h('div', { class: 'col' }, G.f('百分比(&T):', pct), h('div', { class: 'cd-cap', text: '预览' }), pv)), buttons: [{ label: '确定', primary: true, onClick: () => { const v = r.get(); if (v === 'pw') A().zoomTo('页宽'); else if (v === 'tw') A().zoomTo('文字宽度'); else if (v === 'wp') A().zoomTo('整页'); else if (v === 'many') A().zoomTo('两页'); else A().zoomTo(pct.get() + '%'); } }, { label: '取消' }] });
  };

  /* ================= Print (to PDF) ================= */
  G.print = function () {
    const n = LY.pages.length || 1;
    const range = G.radios('pr-r', [['all', '&All'], ['current', '当前页(&E)'], ['selection', '所选内容(&S)'], ['pages', '页(&G):']], 'all');
    const pagesIn = h('input', { type: 'text', id: 'pr-pages', style: 'width:150px' });
    pagesIn.addEventListener('focus', () => range.set('pages'));
    const what = ui.select([['doc', '文档'], ['markup', '显示标记的文档'], ['props', '文档属性'], ['list', '标记列表'], ['styles', 'Styles'], ['autotext', '自动图文集词条'], ['keys', '快捷键分配']], 'doc', null, { id: 'pr-what' });
    const which = ui.select([['all', '区域内的所有页'], ['odd', '奇数页'], ['even', '偶数页']], 'all', null, { id: 'pr-which' });
    const perSheet = ui.select([[1, '1 page'], [2, '2 pages'], [4, '4 pages'], [6, '6 pages'], [8, '8 pages'], [16, '16 pages']], 1, null, { id: 'pr-pps' });
    const copies = G.num(1, { min: 1, max: 99, id: 'pr-copies' });
    ui.dialog({
      title: '打印', width: 470,
      body: h('div', { class: 'col' }, ui.group('打印机', G.f('名称(&N):', ui.select([['pdf', 'Save as PDF (Quire)']], 'pdf', null, { id: 'pr-printer' }), 'wide'), h('div', { class: 'tp-note', text: '状态: 就绪 — 页面已保存为 PDF，可在任意阅读器中打印。' })), G.row(ui.group('页码范围', range, pagesIn, h('div', { class: 'tp-note', text: '输入页码或页码范围，用逗号分隔。例如 1,3,5–12' })), ui.group('份数', G.f('份数(&C):', copies))), G.row(G.f('打印内容(&W):', what, 'wide'), G.f('打印(&R):', which)), ui.group('显示比例', G.f('每张纸的页数(&H):', perSheet, 'wide'))),
      buttons: [{ label: '选项(&O)...', onClick: () => { G.options('print'); return false; } }, { label: '确定', primary: true, onClick: () => { doPrint(); } }, { label: '取消' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    async function doPrint() {
      const w = what.value;
      if (w !== 'doc' && w !== 'markup') { printList(w); return; }
      let pages = null;
      const rv = range.get();
      if (rv === 'current') pages = [Math.max(0, LY.pageOf(E.sel.f)) + 1];
      else if (rv === 'selection' && E.sel && !E.collapsed()) { const [a, b] = E.range(); const p0 = LY.pageOf(a), p1 = LY.pageOf(b); pages = []; for (let i = p0; i <= p1; i++) pages.push(i + 1); }
      else if (rv === 'pages') { pages = []; for (const part of pagesIn.value.split(',')) { const m = /^\s*(\d+)\s*(?:[-–]\s*(\d+))?\s*$/.exec(part); if (m) { const a = +m[1], b = m[2] ? +m[2] : a; for (let i = a; i <= b; i++) if (i >= 1 && i <= n) pages.push(i); } } }
      if (which.value !== 'all') pages = (pages || Array.from({ length: n }, (_, i) => i + 1)).filter((i) => (which.value === 'odd' ? i % 2 === 1 : i % 2 === 0));
      const prevMarkup = LY.opts.markup;
      if (w === 'doc' && prevMarkup !== 'final') { LY.opts.markup = 'final'; LY.render(); }
      try {
        const blob = await A().buildPDF(pages);
        L.ui.busy(false);
        doc().props.printed = new Date().toISOString();
        await L.saveFile(`${A().fileName}.pdf`, blob);
      } catch (e) { L.ui.busy(false); ui.msg('无法创建 PDF：' + (e.message || e), { icon: 'error' }); }
      finally { if (LY.opts.markup !== prevMarkup) { LY.opts.markup = prevMarkup; LY.render(); E.restoreDom(true); } }
    }
    function printList(kind) {
      const d = doc();
      const lines = [];
      if (kind === 'props') { for (const [k, v] of Object.entries(d.props)) if (v) lines.push(`${k}: ${v}`); }
      else if (kind === 'styles') { for (const k of Object.keys(d.styles).sort()) { const s = d.styles[k]; lines.push(`${D.styleDisplayName(s)} (${s.type})${s.basedOn ? ' — based on ' + s.basedOn : ''}`); } }
      else if (kind === 'autotext') { for (const e of L.autocorrect.autoTextEntries()) lines.push(`${e.cat}: ${e.name}`); }
      else if (kind === 'keys') { for (const [k, id] of Object.entries(A().SHORTCUTS)) if (ui.cmds[id]) lines.push(`${k}\t${ui.stripAmp(ui.cmds[id].label)}`); }
      else if (kind === 'list') { for (const r of L.review.list(d.main)) lines.push(`${r.type} — ${r.stamp.author || ''}: ${r.text}`); for (const k in d.comments) lines.push(`Comment ${d.comments[k].author}: ${d.comments[k].blocks.map((b) => (b.t === 'p' ? D.plainText(b) : '')).join(' ')}`); }
      const nd = D.newDoc();
      nd.main.blocks = [D.para([D.text(A().fileName + ' — ' + what.selectedOptions[0].text, { b: true, sz: 14 })])].concat(lines.map((l) => D.para([D.text(l)])));
      const keep = D.doc;
      D.doc = nd;
      LY.render();
      A().buildPDF().then(async (blob) => { L.ui.busy(false); D.doc = keep; LY.render(); E.restoreDom(true); await L.saveFile(`${A().fileName} (${kind}).pdf`, blob); }).catch(() => { D.doc = keep; LY.render(); L.ui.busy(false); });
    }
  };

  /* ================= Save As ================= */
  G.saveAs = function (cb) {
    const name = h('input', { type: 'text', id: 'sa-name', value: A().fileName, style: 'width:100%' });
    const type = ui.select([['docx', 'Word Document (*.docx)'], ['docm', 'Word 启用宏的文档 (*.docm)'], ['dotx', 'Word 模板 (*.dotx)'], ['dotm', 'Word 启用宏的模板 (*.dotm)'], ['html', 'Single File Web Page (*.htm)'], ['rtf', 'RTF 格式 (*.rtf)'], ['txt', 'Plain Text (*.txt)'], ['md', 'Markdown (*.md)'], ['mdzip', 'Markdown with pictures (*.zip)'], ['pdf', 'PDF (*.pdf)']], ['docm', 'dotx', 'dotm', 'md', 'mdzip'].includes(A().fileType) ? A().fileType : 'docx', null, { id: 'sa-type' });
    const note = h('div', { class: 'tp-note', text: '浏览器会询问保存位置，或直接保存到「下载」文件夹。' });
    return ui.dialog({ title: '另存为', width: 430, focus: name, body: h('div', { class: 'col' }, G.f('文件名(&N):', name, 'wide'), G.f('保存类型(&T):', type, 'wide'), note), buttons: [{ label: '保存(&S)', primary: true, onClick: () => { const n = name.value.trim().replace(/\.(docx|docm|dotx|dotm|htm|html|rtf|txt|md|zip|pdf)$/i, '') || 'Document'; cb(n, type.value); } }, { label: '取消' }] });
  };
  G.pasteSpecial = function () {
    const lb = G.listBox([{ value: 'rich', label: '带格式文本 (RTF)' }, { value: 'html', label: 'HTML 格式' }, { value: 'text', label: '无格式文本' }, { value: 'utext', label: '无格式 Unicode 文本' }], 'rich', { height: 110 });
    const clip = E.clip;
    ui.dialog({ title: '选择性粘贴', width: 380, body: h('div', { class: 'col' }, h('div', { text: clip ? '来源: Quire 文档' : '来源: （Quire 剪贴板为空 —— 请先复制内容，或按 Ctrl+V）' }), G.radios('ps-m', [['paste', '粘贴(&P):'], ['link', '粘贴链接(&L):']], 'paste'), lb, h('div', { class: 'tp-note', text: '带格式文本会保留字体和格式；无格式文本则采用目标位置的格式。' })), buttons: [{ label: '确定', primary: true, onClick: () => {
      if (!clip) return;
      const v = lb.get();
      if (v === 'text' || v === 'utext') E.pasteText(clip.text);
      else E.pasteBlocks(clip.blocks, { keepFormat: v === 'rich' || v === 'html' });
    } }, { label: '取消' }] });
  };
  G.about = () => ui.about({ name: 'Quire', office: 'Word', text: [
    'A word processor in the style of Microsoft Office Word 2003 that runs entirely in your browser. Documents are read and written as Office Open XML (.docx, .docm), RTF, HTML, text and Markdown; nothing is uploaded.',
    '自动更正可离线使用。'] });
  G.newFromTemplate = function () {
    const cats = L.templates.categories();
    let pick = 'blank';
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:190px' });
    const drawPv = (id) => { L.clear(pv); const page = h('div', { class: 'ps-page', style: 'width:110px;height:140px' }); for (let i = 0; i < (id === 'blank' ? 0 : 14); i++) page.appendChild(h('i', { style: `top:${14 + i * 8}px;${i % 5 === 4 ? 'right:40px' : ''}` })); pv.appendChild(page); };
    const tabs = ui.tabs(cats.map((c) => {
      const grid = h('div', { class: 'preset-grid' });
      const items = c === 'General' ? [{ id: 'blank', name: '空白文档' }, { id: 'web', name: '网页' }, { id: 'mail', name: '电子邮件' }] : L.templates.list().filter((t) => t.cat === c);
      for (const t of items) { const b = h('button', { type: 'button', class: t.id === pick ? 'on' : '' }, h('span', { html: L.icons.get('new', 32) || L.icons.get('new') }), h('span', { text: t.name })); b.addEventListener('click', () => { pick = t.id; L.$$('.preset-grid button', tabs).forEach((x) => x.classList.toggle('on', x === b)); drawPv(t.id); }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
      return { label: c, body: grid };
    }));
    drawPv('blank');
    const dlg = ui.dialog({ title: '模板', width: 520, body: G.row(h('div', { style: 'flex:1;min-width:240px' }, tabs), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: '预览' }), pv)), buttons: [{ label: '确定', primary: true, onClick: () => { if (pick === 'blank') A().newDocument(); else if (pick === 'web') { A().newDocument().then(() => A().setView('web')); } else if (pick === 'mail') A().newDocument({ build: (d) => { d.main.blocks = [D.para([D.text('To: ')]), D.para([D.text('Subject: ')]), D.para()]; } }); else L.templates.create(pick); } }, { label: '取消' }] });
    const ok = dlg.buttons[0];
  };
})();
