/* Ledger — the task pane (Getting Started, New Workbook, Help, Clipboard, Clip Art) and the spreadsheet templates. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, O = L.ops;
  const { h } = L;
  const ui = L.ui;
  const P = (L.panes = {});
  const A = () => L.app;
  const G = () => L.grid;

  /* ================= shell ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    const nav = (icon, label) => h('button', { class: 'tp-nav', type: 'button', 'aria-label': label, 'data-tip': label, html: L.icons.get(icon) || label[0] });
    TP.back = nav('back', L.N('后退')); TP.fwd = nav('forward', L.N('向前')); TP.home = nav('home', L.N('主页'));
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '关闭任务窗格', 'data-tip': '关闭', html: '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M1 1l8 8M9 1L1 9" stroke="#000" stroke-width="1.6"/></svg>' });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => A().toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
    L.bus.on('clipboard', () => TP.refresh('clip'));
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id] || !TP.el) return;
    if (TP.el.hidden) A().toggleTask(true);
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
    TP.body.textContent = '';
    p.render(TP.body);
  };
  TP.refresh = function (reason) {
    const p = TP.registry[TP.current];
    if (!p || !TP.el || TP.el.hidden) return;
    if (p.live) { TP.body.textContent = ''; p.render(TP.body); }
    void reason;
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn, o) => h('button', Object.assign({ class: 'tp-link', type: 'button', onclick: fn }, o || {}), icon ? h('span', { html: L.icons.get(icon) || '' }) : null, h('span', { text: label }));
  P.sec = sec; P.link = link;

  /* ---------- Getting Started ---------- */
  /* ---------- Document Recovery: unsaved versions kept by common/suite.js ---------- */
  TP.register({ id: 'recovery', title: '文档恢复', render(b) { if (window.VO) VO.recoveryPane(b); } });
  TP.register({
    id: 'getting-started', title: '入门',
    render(b) {
      const q = h('input', { type: 'text', id: 'gs-search', placeholder: '搜索帮助', style: 'flex:1;min-width:0' });
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { P.helpQuery = q.value; TP.show('help'); } e.stopPropagation(); });
      b.append(
        h('div', { class: 'tp-row tp-brand' }, h('span', { html: L.icons.app ? L.icons.app(32) : '' }), h('b', { text: 'Ledger 网络版' })),
        sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec(L.N('打开'), window.VO ? VO.recentLinks(4) : null, link('open', '更多...', () => ui.exec('open')), link('new', '新建工作簿...', () => TP.show('new'))),
        sec(L.N('学习'), link('help', 'Ledger 的功能', () => { P.helpQuery = ''; TP.show('help'); }), link('fx', '函数参考', () => L.dlg.insertFunction()), link('help', '键盘快捷键', () => TP.show('keys'))),
        sec(L.N('文件'), h('div', { class: 'tp-note', text: '可打开 .xlsx、.xlsm、.xltx、.csv、.txt 和 XML Spreadsheet 2003 文件；将文件拖到窗口任意位置即可打开。可保存为 .xlsx、.xlsm、.xltx、.csv、.txt、.htm、.xml 和 PDF。' })));
    },
  });
  /* ---------- New Workbook ---------- */
  TP.register({
    id: 'new', title: '新建工作簿',
    render(b) {
      b.append(
        sec(L.N('新建'), link('new', '空白工作簿', () => A().newWorkbook()), link('open', '根据现有工作簿...', () => A().openDialog('.xlsx,.xlsm,.xltx,.xltm'))),
        sec(L.N('模板'), ...Object.keys(TPL).map((k) => link('sheetIcon', TPL[k].name, () => P.newFromTemplate(k)))),
        sec('', h('div', { class: 'tp-note', text: '模板将作为未保存的新工作簿打开。' })));
    },
  });
  /* ---------- Help ---------- */
  const HELP = [
    ['输入数据', L.N('单击单元格并键入内容，然后按 Enter（下移）、Tab（右移）或方向键。F2 可编辑活动单元格；Alt+Enter 可在单元格内换行；Ctrl+Enter 可将键入内容填入所有选定单元格。数字、日期（3/14/2026、14-Mar）、时间（9:30 AM）、百分比（12%）和货币（$1,200）会被自动识别并设置格式。在内容前加单引号可将其保留为文本。')],
    [L.N('公式'), L.N('以 = 开头并使用单元格引用（A1、$A$1、Sheet2!B3、命名区域以及 Sales[Amount] 等表引用）。输入公式时，可在工作表中单击或拖动以插入引用，按 F4 可循环切换绝对引用和相对引用。Ctrl+Shift+Enter 可输入数组公式；FILTER、SORT、UNIQUE 和 SEQUENCE 等动态数组函数会溢出显示结果。“插入 ▸ 函数”(Shift+F3) 列出超过 450 个函数并提供参数帮助。')],
    ['自动填充和序列', L.N('拖动填充柄（选区右下角的小方块）可复制单元格或扩展序列：1、2 → 3、4…；Mon → Tue；Jan → Feb；Q1 → Q2；日期可按日、工作日、月或年递增。双击填充柄可沿相邻数据向下填充。“编辑 ▸ 填充 ▸ 序列”提供完整设置；Ctrl+D 和 Ctrl+R 分别向下和向右填充。')],
    ['设置单元格格式', L.N('“格式 ▸ 单元格”(Ctrl+1) 包含“数字”、“对齐”、“字体”、“边框”、“图案”和“保护”选项卡。自定义数字格式使用 Excel 格式代码，如 #,##0.00;[Red](#,##0.00)、0%、d-mmm-yy、[h]:mm 和 "text"@。“格式”工具栏可设置字体、对齐、合并及居中、货币、百分比、千位分隔样式、小数位数、缩进、边框、填充色及字体颜色。格式刷用于复制格式。')],
    ['行、列和工作表', L.N('拖动行列标题的边框可调整大小，双击可自动调整。右键单击行列标题可插入、删除、隐藏或取消隐藏。“窗口 ▸ 冻结窗格”可让标题保持可见。右键单击工作表标签可插入、删除、重命名、移动或复制工作表，也可设置标签颜色；Ctrl+Page Up / Page Down 可切换工作表。')],
    ['排序和筛选', L.N('“数据 ▸ 排序”支持最多三个关键字，可升序或降序排列，也可指定有无标题行；排序按钮则按活动列排序。“数据 ▸ 筛选 ▸ 自动筛选”会在标题行添加下拉箭头，可选择某个值、“前 10 个...”、“自定义...”、“空白”或“非空白”。高级筛选使用条件区域，并可将不重复的记录复制到其他位置。“数据 ▸ 分类汇总”会在指定列内容变化处插入 SUBTOTAL 汇总行，并建立分级显示。')],
    ['列表（表格）', L.N('“数据 ▸ 列表 ▸ 创建列表”(Ctrl+L) 可将区域转换为包含标题行、自动筛选和可选汇总行的列表。公式可以通过名称引用列表，例如：=SUM(Table1[Amount])。')],
    ['条件格式和数据验证', L.N('“格式 ▸ 条件格式”可为每个区域设置最多三个条件（单元格值介于、大于…或公式），并支持从较新工作簿中读取数据条、色阶和图标集。“数据 ▸ 数据验证”可将输入限制为整数、小数、序列、日期、时间、文本长度或自定义公式，并提供输入信息和出错警告。')],
    ['图表和图片', L.N('“插入 ▸ 图表”(F11) 可打开图表向导，创建柱形图、条形图、折线图、饼图、XY 散点图、面积图和圆环图，并将其嵌入工作表或置于独立图表工作表中。图表会随数据更改而更新。“插入 ▸ 图片”可添加图片和自选图形；拖动可移动，拖动控点可调整大小，双击可设置格式。')],
    ['名称、批注和超链接', L.N('在名称框中键入名称可为选区命名，或使用“插入 ▸ 名称 ▸ 定义”(Ctrl+F3)。“插入 ▸ 批注”(Shift+F2) 可附加注释；带批注的单元格会显示红色三角形。“插入 ▸ 超链接”(Ctrl+K) 可链接到网页、电子邮件地址或工作簿中的位置。')],
    ['公式审核', L.N('“工具 ▸ 公式审核”可用箭头追踪引用单元格和从属单元格，逐步对公式求值，并显示监视窗口。Ctrl+` 可显示公式而非结果。“错误检查”会逐个检查有错误的单元格。')],
    ['打印和 PDF', L.N('“文件 ▸ 页面设置”可设置方向、缩放（调整为指定页数）、纸张大小、页边距、居中、页眉和页脚（&[Page]、&[Pages]、&[Date]、&[File]、&[Tab]）、打印区域、重复打印的行以及网格线。“视图 ▸ 分页预览”显示各页；“插入 ▸ 分页符”添加手动分页符。“文件 ▸ 打印”和“打印预览”会生成可在任何地方打印的 PDF。')],
    ['文件与兼容性', L.N('Ledger 可读写 Office Open XML 工作簿（.xlsx、启用宏的 .xlsm 及模板），保留公式、样式、合并单元格、批注、列表、条件格式、数据验证、图表、图片、定义的名称、页面设置和宏。它还可通过文本导入向导打开 CSV 和文本文件，打开有密码保护的工作簿，并可使用密码保存（“另存为 ▸ 工具 ▸ 常规选项”）。二进制 .xls 文件须先转换为 .xlsx。')],
    [L.N('保护'), L.N('“工具 ▸ 保护 ▸ 保护工作表”会锁定所有勾选了“锁定”（“格式 ▸ 单元格 ▸ 保护”）的单元格；您可指定用户仍能执行的操作，并设置可选密码。“保护工作簿”会锁定工作表结构。')],
    ['模拟分析', L.N('“工具 ▸ 单变量求解”可找到使公式达到目标值的输入值。“数据 ▸ 模拟运算表”可建立单变量或双变量模拟运算表。“工具 ▸ 方案”可存储多组输入值以便比较。')],
  ];
  TP.register({
    id: 'help', title: 'Ledger 帮助',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: '搜索', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        res.textContent = '';
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        for (const [t, d] of HELP.filter(([t, d]) => !words.length || hit(L.t(t) + ' ' + L.t(d)))) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        if (words.length) {
          const fns = L.fninfo ? L.fninfo.all().filter((f) => hit(f.name + ' ' + L.t(f.desc))).slice(0, 12) : [];
          if (fns.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: '函数' }), ...fns.map((f) => { const x = h('div', { class: 'help-fn' }, h('b', { translate: 'no', text: L.fninfo.syntax(f) }), h('div', { text: f.desc })); x.addEventListener('click', () => L.dlg.insertFunction(f.name)); return x; })));
          const keys = KEYS.filter(([k, d]) => hit(L.t(k) + ' ' + L.t(d)));
          if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: '键盘快捷键' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))));
        }
        if (!res.childNodes.length) res.appendChild(h('div', { class: 'tp-note', text: '没有与搜索内容匹配的帮助主题。' }));
      };
      q.addEventListener('input', draw);
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
      P.helpQuery = '';
    },
  });
  const KEYS = [
    [L.N('Ctrl+N / O / S / P'), '新建、打开、保存、打印'], ['F12', '另存为'], [L.N('Ctrl+Z / Ctrl+Y'), '撤消/重做'], ['F4', L.N('重复操作；在公式中循环切换 $ 引用方式')],
    [L.N('Ctrl+X / C / V'), '剪切、复制、粘贴'], [L.N('Ctrl+Alt+V'), '选择性粘贴'], [L.N('Ctrl+F / H'), '查找、替换'], [L.N('F5 / Ctrl+G'), '定位'],
    ['F2', '编辑活动单元格'], [L.N('Esc'), '取消输入'], [L.N('Alt+Enter'), '在单元格内换行'], [L.N('Ctrl+Enter'), '将输入内容填入选区'], [L.N('Ctrl+Shift+Enter'), '输入数组公式'],
    [L.N('Ctrl+D / Ctrl+R'), '向下/向右填充'], [L.N('Ctrl+; / Ctrl+Shift+:'), '插入日期/时间'], [L.N("Ctrl+' / Ctrl+Shift+\""), '复制上方单元格的公式/值'], [L.N('Alt+='), L.N('自动求和')], [L.N('Shift+F3'), '插入函数'], [L.N('Ctrl+A（在公式中）'), '函数参数'], ['F9', '立即计算'],
    [L.N('方向键'), '移动一个单元格'], [L.N('Ctrl+方向键'), '移动到数据区域边缘'], [L.N('Home / Ctrl+Home'), '行首/A1'], [L.N('Ctrl+End'), '最后一个使用过的单元格'], [L.N('End、方向键'), '结束模式跳转'], [L.N('Page Up / Down'), '移动一屏'], [L.N('Alt+Page Up / Down'), '向左/向右移动一屏'], [L.N('Ctrl+Page Up / Down'), '上一张/下一张工作表'],
    [L.N('Shift+方向键'), '扩展选区'], [L.N('Ctrl+Shift+方向键'), '扩展到数据边缘'], [L.N('Ctrl+Space / Shift+Space'), '选择列/行'], [L.N('Ctrl+A'), '先选择当前区域，再选择整个工作表'], [L.N('Ctrl+Shift+*'), '选择当前区域'], [L.N('Ctrl+/'), '选择当前数组'],
    [L.N('Ctrl+1'), '设置单元格格式'], [L.N('Ctrl+B / I / U / 5'), '加粗、倾斜、下划线、删除线'], [L.N('Ctrl+Shift+~ $ % ^ # @ !'), '常规、货币、百分比、科学记数、日期、时间、数字格式'], [L.N('Ctrl+Shift+& / _'), '外边框/清除边框'],
    [L.N('Ctrl+9 / Ctrl+Shift+('), '隐藏/取消隐藏行'], [L.N('Ctrl+0 / Ctrl+Shift+)'), '隐藏/取消隐藏列'], [L.N('Ctrl+- / Ctrl+Shift++'), '删除/插入单元格'], [L.N('Shift+F11'), '插入工作表'], ['F11', L.N('图表')], [L.N('Shift+F2'), '插入或编辑批注'], [L.N('Ctrl+K'), L.N('超链接')], [L.N('Ctrl+F3 / F3'), '定义/粘贴名称'], [L.N('Ctrl+L'), '创建列表'],
    [L.N('Alt+↓'), '从下拉列表中选择'], [L.N('Ctrl+`'), '显示公式'], ['F7', L.N('拼写')], [L.N('Shift+F10'), '快捷菜单'], [L.N('Ctrl+F1'), '任务窗格'], ['F1', L.N('帮助')], [L.N('Ctrl+滚轮'), L.N('缩放')],
  ];
  P.KEYS = KEYS;
  TP.register({
    id: 'keys', title: '键盘快捷键',
    render(b) { b.appendChild(h('table', { class: 'kbd-table' }, ...KEYS.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))); },
  });
  /* ---------- Office Clipboard ---------- */
  P.clips = [];
  P.addClip = function (clip) {
    P.clips.unshift({ clip, text: (clip.text || '').replace(/\t/g, '  ').slice(0, 200), time: Date.now() });
    if (P.clips.length > 24) P.clips.length = 24;
    L.bus.emit('clipboard');
  };
  TP.register({
    id: 'clipboard', title: '剪贴板', live: true,
    render(b) {
      const items = P.clips;
      b.append(h('div', { class: 'tp-btns' }, ui.button('全部粘贴', () => { for (const it of items.slice().reverse()) { L.clip.set(it.clip); L.clip.pasteInternal({}); } }, { class: 'btn small' }), ui.button('全部清除', () => { items.length = 0; TP.refresh('clip'); }, { class: 'btn small' })),
        h('div', { class: 'tp-note', text: L.t('剪贴板：{0}/24。单击项目可粘贴。', items.length) }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { html: L.icons.get('sheetIcon') || '' }), h('span', { class: 'cb-t', translate: 'no', text: it.text || '[单元格]' }));
        el.addEventListener('click', () => { L.clip.set(it.clip); L.clip.pasteInternal({}); G().paint(); });
        b.appendChild(el);
      }
    },
  });

  /* ================= templates ================= */
  /** build a workbook from a compact description */
  function build(spec) {
    const wb = A().blankWorkbook();
    wb.sheets.splice(0);
    for (const s of spec.sheets) {
      const sh = wb.addSheet(s.name);
      const style = (base, d) => wb.styles.derive(base || 0, d);
      for (const [addr, val, st] of s.cells) {
        const p = F.parseCell(addr);
        const cell = sh.cell(p.r, p.c);
        if (typeof val === 'string' && val[0] === '=') { cell.f = F.toStore(val.slice(1)); cell.dirty = true; }
        else if (val !== null && val !== undefined) cell.v = val;
        if (st) cell.s = style(cell.s, st);
      }
      for (const [rg, st] of s.ranges || []) { const a = F.parseRange(rg); for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const cell = sh.cell(r, c); cell.s = style(cell.s, st); } }
      for (const [col, w] of Object.entries(s.cols || {})) sh.cols[F.colIndex(col)] = { w, custom: true };
      for (const m of s.merges || []) sh.merges.push(F.parseRange(m));
      if (s.freeze) sh.view.freeze = { r: s.freeze[0], c: s.freeze[1], top: { r: s.freeze[0], c: s.freeze[1] } };
      if (s.grid === false) sh.view.grid = false;
      if (s.print) Object.assign(sh.print, s.print);
      if (s.dv) for (const d of s.dv) sh.dv.push(Object.assign({ ranges: M.parseSqref(d.ref), allowBlank: true, showDrop: true, showInput: true, showError: true }, d));
      if (s.sel) sh.view.sel = { r: F.parseCell(s.sel).r, c: F.parseCell(s.sel).c, ranges: [F.parseRange(s.sel)] };
    }
    wb.active = 0;
    return wb;
  }
  const B = { font: { b: true } }, HDR = { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#336699') }, align: { h: 'center', v: 'center', wrap: true } };
  const TITLE = { font: { b: true, sz: 16, color: M.rgb('#1F3F7F') } };
  const MONEY = { nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' };
  const BOX = { border: { l: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') }, t: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#808080') } } };
  const INPUT = { fill: { pattern: 'solid', fg: M.rgb('#FFFFCC') }, prot: { locked: false } };
  const TPL = {
    loan: {
      name: '贷款摊销',
      spec: () => {
        const cells = [
          ['A1', '贷款摊销明细表', TITLE],
          ['A3', '贷款金额', B], ['C3', 25000, Object.assign({}, MONEY, INPUT, BOX)],
          ['A4', '年利率', B], ['C4', 0.065, Object.assign({ nf: '0.00%' }, INPUT, BOX)],
          ['A5', '贷款年限', B], ['C5', 5, Object.assign({}, INPUT, BOX)],
          ['A6', '每年还款次数', B], ['C6', 12, Object.assign({}, INPUT, BOX)],
          ['A7', '贷款起始日期', B], ['C7', 46113, Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)],
          ['E3', '每期还款额', B], ['G3', '=IF(C3*C4*C5*C6>0,-PMT(C4/C6,C5*C6,C3),"")', Object.assign({}, MONEY, BOX)],
          ['E4', '还款期数', B], ['G4', '=C5*C6', BOX],
          ['E5', '利息总额', B], ['G5', '=SUM(F11:F370)', Object.assign({}, MONEY, BOX)],
          ['E6', '贷款总成本', B], ['G6', '=C3+G5', Object.assign({}, MONEY, BOX)],
          ['A10', '期次', HDR], ['B10', '还款日期', HDR], ['C10', '期初余额', HDR], ['D10', '还款额', HDR], ['E10', '本金', HDR], ['F10', '利息', HDR], ['G10', '期末余额', HDR], ['H10', '累计利息', HDR],
        ];
        for (let i = 0; i < 60; i++) {
          const r = 11 + i;
          cells.push(['A' + r, `=IF(ROW()-10<=$G$4,ROW()-10,"")`], ['B' + r, `=IF(A${r}="","",EDATE($C$7,(A${r}-1)*12/$C$6))`, { nf: 'm/d/yyyy' }],
            ['C' + r, i === 0 ? `=IF(A${r}="","",$C$3)` : `=IF(A${r}="","",G${r - 1})`, MONEY], ['D' + r, `=IF(A${r}="","",MIN($G$3,C${r}+F${r}))`, MONEY],
            ['E' + r, `=IF(A${r}="","",D${r}-F${r})`, MONEY], ['F' + r, `=IF(A${r}="","",C${r}*$C$4/$C$6)`, MONEY], ['G' + r, `=IF(A${r}="","",C${r}-E${r})`, MONEY], ['H' + r, i === 0 ? `=IF(A${r}="","",F${r})` : `=IF(A${r}="","",H${r - 1}+F${r})`, MONEY]);
        }
        return { sheets: [{ name: 'Loan', cells, cols: { A: 20, B: 13, C: 16, D: 13, E: 18, F: 13, G: 15, H: 14 }, freeze: [10, 0], print: { titleRows: null, fit: true, fitW: 1, fitH: null }, sel: 'C3' }] };
      },
    },
    expense: {
      name: '费用报表',
      spec: () => {
        const cells = [
          ['A1', '费用报销单', TITLE], ['A3', '员工', B], ['B3', '', Object.assign({}, INPUT, BOX)], ['D3', '期间自', B], ['E3', '', Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)], ['F3', '至', B], ['G3', '', Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)],
          ['A5', '日期', HDR], ['B5', '摘要', HDR], ['C5', '类别', HDR], ['D5', '里程', HDR], ['E5', '里程补贴（$0.67/英里）', HDR], ['F5', '金额', HDR], ['G5', '合计', HDR],
          ['A6', 46120, { nf: 'm/d/yyyy' }], ['B6', '拜访客户，斯普林菲尔德'], ['C6', '差旅'], ['D6', 84], ['F6', 18.5, MONEY],
          ['A7', 46121, { nf: 'm/d/yyyy' }], ['B7', '与客户共进午餐'], ['C7', '餐饮'], ['F7', 46.2, MONEY],
          ['A8', 46124, { nf: 'm/d/yyyy' }], ['B8', '会议注册费'], ['C8', '培训'], ['F8', 395, MONEY],
          ['F20', '小计', B], ['G20', '=SUM(G6:G19)', Object.assign({ font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } } }, MONEY)],
          ['F21', '减预支款', B], ['G21', 0, Object.assign({}, MONEY, INPUT)], ['F22', '应付合计', B], ['G22', '=G20-G21', Object.assign({ font: { b: true } }, MONEY)],
          ['A24', '按类别汇总', B],
        ];
        for (let r = 6; r <= 19; r++) { cells.push(['E' + r, `=IF(D${r}="","",D${r}*0.67)`, MONEY], ['G' + r, `=IF(COUNT(E${r}:F${r})=0,"",SUM(E${r}:F${r}))`, MONEY]); if (r > 8) cells.push(['A' + r, null, { nf: 'm/d/yyyy' }], ['F' + r, null, MONEY]); }
        ['差旅', '餐饮', '住宿', '培训', '用品', '其他'].forEach((c, i) => { cells.push(['A' + (25 + i), c], ['B' + (25 + i), `=SUMIF($C$6:$C$19,A${25 + i},$G$6:$G$19)`, MONEY]); });
        return { sheets: [{ name: 'Expenses', cells, cols: { A: 12, B: 30, C: 12, D: 8, E: 12, F: 12, G: 13 }, dv: [{ ref: 'C6:C19', type: 'list', f1: '"差旅,餐饮,住宿,培训,用品,其他"', prompt: '选择类别', promptTitle: '类别' }], sel: 'B3' }] };
      },
    },
    invoice: {
      name: '销售发票',
      spec: () => {
        const cells = [
          ['A1', '发票', { font: { b: true, sz: 20, color: M.rgb('#336699') } }], ['E1', '发票号', B], ['F1', 1001, Object.assign({}, INPUT, BOX)], ['E2', '日期', B], ['F2', '=TODAY()', Object.assign({ nf: 'mmmm d, yyyy' }, BOX)],
          ['A4', '客户：', B], ['A5', '客户名称', INPUT], ['A6', '街道地址', INPUT], ['A7', '城市、省/市、邮政编码', INPUT],
          ['A9', '数量', HDR], ['B9', '摘要', HDR], ['E9', '单价', HDR], ['F9', '小计', HDR],
          ['A10', 4], ['B10', '标准配件'], ['E10', 12.5, MONEY], ['A11', 1], ['B11', '安装服务'], ['E11', 85, MONEY],
          ['E21', '小计', B], ['F21', '=SUM(F10:F20)', MONEY], ['E22', '销售税率', B], ['F22', 0.0725, Object.assign({ nf: '0.00%' }, INPUT)], ['E23', '销售税', B], ['F23', '=ROUND(F21*F22,2)', MONEY],
          ['E24', '总计', { font: { b: true, sz: 12 } }], ['F24', '=F21+F23', Object.assign({ font: { b: true, sz: 12 }, border: { t: { style: 'thin' }, b: { style: 'double' } } }, MONEY)],
          ['A27', '支票请开给您的公司名称。感谢您的惠顾！', { font: { i: true } }],
        ];
        for (let r = 10; r <= 20; r++) cells.push(['F' + r, `=IF(A${r}*E${r}=0,"",A${r}*E${r})`, MONEY]);
        return { sheets: [{ name: 'Invoice', cells, cols: { A: 8, B: 30, C: 8, D: 8, E: 14, F: 15 }, merges: ['B9:D9'], grid: false, sel: 'A5' }] };
      },
    },
    timecard: {
      name: '工时卡',
      spec: () => {
        const days = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
        const T = { nf: 'h:mm AM/PM' };
        const cells = [['A1', '每周工时卡', TITLE], ['A3', '员工', B], ['B3', '', INPUT], ['A4', '本周起始日', B], ['B4', 46118, Object.assign({ nf: 'm/d/yyyy' }, INPUT)], ['D3', '时薪', B], ['E3', 22.5, Object.assign({}, MONEY, INPUT)],
          ['A6', '星期', HDR], ['B6', '日期', HDR], ['C6', '上班', HDR], ['D6', '午休开始', HDR], ['E6', '午休结束', HDR], ['F6', '下班', HDR], ['G6', '工时', HDR]];
        days.forEach((d, i) => {
          const r = 7 + i;
          cells.push(['A' + r, d], ['B' + r, `=$B$4+${i}`, { nf: 'm/d' }], ['C' + r, i < 5 ? 0.375 : null, T], ['D' + r, i < 5 ? 0.5 : null, T], ['E' + r, i < 5 ? 0.53125 : null, T], ['F' + r, i < 5 ? 0.7083333333 : null, T],
            ['G' + r, `=IF(COUNT(C${r}:F${r})<2,0,ROUND(((F${r}-C${r})-(E${r}-D${r}))*24,2))`, { nf: '0.00' }]);
        });
        cells.push(['F15', '总工时', B], ['G15', '=SUM(G7:G13)', { nf: '0.00', font: { b: true } }], ['F16', '正常工时（≤40）', B], ['G16', '=MIN(40,G15)', { nf: '0.00' }], ['F17', '加班', B], ['G17', '=MAX(0,G15-40)', { nf: '0.00' }],
          ['F19', '应付工资', B], ['G19', '=G16*E3+G17*E3*1.5', Object.assign({ font: { b: true } }, MONEY)]);
        return { sheets: [{ name: 'Timecard', cells, cols: { A: 12, B: 11, C: 11, D: 11, E: 11, F: 13, G: 11 }, sel: 'B3' }] };
      },
    },
    budget: {
      name: '个人月度预算',
      spec: () => {
        const months = ['1 月', '2 月', '3 月', '4 月', '5 月', '6 月', '7 月', '8 月', '9 月', '10 月', '11 月', '12 月'];
        const cols = 'BCDEFGHIJKLM';
        const N = { nf: '#,##0' };
        const cells = [['A1', '个人月度预算', TITLE], ['A3', '收入', HDR], ['N3', '年度', HDR]];
        months.forEach((m, i) => cells.push([cols[i] + '3', m, HDR]));
        const income = [['工资', 4200], ['其他收入', 150]];
        const exp = [['房租 / 房贷', 1450], ['水电燃气', 210], ['食品杂货', 520], ['交通', 260], ['保险', 180], ['电话与网络', 95], ['娱乐', 140], ['储蓄', 400]];
        income.forEach(([n, v], k) => { const r = 4 + k; cells.push(['A' + r, n]); months.forEach((m, i) => cells.push([cols[i] + r, v, N])); cells.push(['N' + r, `=SUM(B${r}:M${r})`, N]); });
        cells.push(['A6', '收入合计', B]);
        for (const c of cols + 'N') cells.push([c + '6', `=SUM(${c}4:${c}5)`, Object.assign({ font: { b: true }, border: { t: { style: 'thin' } } }, N)]);
        cells.push(['A8', '支出', HDR]);
        months.forEach((m, i) => cells.push([cols[i] + '8', m, HDR]));
        cells.push(['N8', '年度', HDR]);
        exp.forEach(([n, v], k) => { const r = 9 + k; cells.push(['A' + r, n]); months.forEach((m, i) => cells.push([cols[i] + r, v, N])); cells.push(['N' + r, `=SUM(B${r}:M${r})`, N]); });
        const last = 8 + exp.length;
        cells.push(['A' + (last + 1), '支出合计', B]);
        for (const c of cols + 'N') cells.push([c + (last + 1), `=SUM(${c}9:${c}${last})`, Object.assign({ font: { b: true }, border: { t: { style: 'thin' } } }, N)]);
        cells.push(['A' + (last + 3), '结余（收入 − 支出）', B]);
        for (const c of cols + 'N') cells.push([c + (last + 3), `=${c}6-${c}${last + 1}`, Object.assign({ font: { b: true }, nf: '#,##0;[Red]-#,##0' })]);
        return { sheets: [{ name: 'Budget', cells, cols: { A: 24 }, freeze: [3, 1], sel: 'B4' }] };
      },
    },
    tasks: {
      name: '任务跟踪表',
      spec: () => {
        const DAY = { nf: 'm/d/yyyy' };
        const cells = [
          ['A1', '任务跟踪表', TITLE],
          ['A2', '未完成', B], ['B2', '=COUNTIFS(A5:A104,"<>",E5:E104,"<>已完成")', { font: { b: true } }],
          ['C2', '已逾期', B], ['D2', '=COUNTIFS(A5:A104,"<>",E5:E104,"<>已完成",D5:D104,"<"&TODAY())', { font: { b: true, color: M.rgb('#C00000') } }],
          ['A4', '任务', HDR], ['B4', '负责人', HDR], ['C4', '优先级', HDR], ['D4', '截止日期', HDR], ['E4', '状态', HDR], ['F4', '剩余天数', HDR], ['G4', '备注', HDR],
          ['A5', '预订场地'], ['B5', '山姆'], ['C5', '高'], ['D5', '=TODAY()+2', DAY], ['E5', '进行中'],
          ['A6', '发送邀请'], ['B6', '阿历克斯'], ['C6', '中'], ['D6', '=TODAY()+7', DAY], ['E6', '未开始'],
          ['A7', '订购用品'], ['B7', '山姆'], ['C7', '低'], ['D7', '=TODAY()-1', DAY], ['E7', '等待中'], ['G7', '已询价'],
          ['A8', '起草预算'], ['B8', '阿历克斯'], ['C8', '高'], ['D8', '=TODAY()-3', DAY], ['E8', '已完成'],
        ];
        for (let r = 5; r <= 104; r++) {
          cells.push(['F' + r, `=IF(OR(A${r}="",D${r}="",E${r}="已完成"),"",D${r}-TODAY())`, { nf: '0;[Red]-0', align: { h: 'center' } }]);
          if (r > 8) cells.push(['D' + r, null, DAY]);
        }
        return { sheets: [{ name: 'Tasks', cells, cols: { A: 32, B: 12, C: 10, D: 11, E: 13, F: 10, G: 30 }, freeze: [4, 0],
          dv: [{ ref: 'C5:C104', type: 'list', f1: '"高,中,低"' }, { ref: 'E5:E104', type: 'list', f1: '"未开始,进行中,等待中,已完成"', prompt: '选择状态', promptTitle: '状态' }],
          sel: 'A9' }] };
      },
    },
    shopping: {
      name: '购物清单',
      spec: () => {
        const cells = [
          ['A1', '购物清单', TITLE],
          ['A2', '预算', B], ['B2', 80, Object.assign({}, MONEY, INPUT, BOX)], ['D2', '清单合计', B], ['E2', '=SUM(F5:F44)', MONEY], ['D3', '还需购买', B], ['E3', '=SUMIF(A5:A44,"",F5:F44)', MONEY],
          ['F2', '=IF(B2="","",IF(E2>B2,"超出预算 "&TEXT(E2-B2,"$0.00"),"剩余 "&TEXT(B2-E2,"$0.00")))', { font: { i: true } }],
          ['A4', '✓', HDR], ['B4', '商品', HDR], ['C4', '类别', HDR], ['D4', '数量', HDR], ['E4', '单价', HDR], ['F4', '小计', HDR],
        ];
        const items = [['苹果', '果蔬', 6, 0.5], ['牛奶', '乳制品', 2, 1.29], ['面包', '烘焙', 1, 2.99], ['鸡蛋（一打）', '乳制品', 1, 3.49], ['鸡胸肉', '肉类与水产', 1, 7.5], ['意面', '干货', 2, 1.19], ['咖啡', '干货', 1, 8.99], ['洗洁精', '日用品', 1, 2.49]];
        items.forEach(([n, c, q, p], i) => { const r = 5 + i; cells.push(['B' + r, n], ['C' + r, c], ['D' + r, q], ['E' + r, p, MONEY]); });
        for (let r = 5; r <= 44; r++) {
          cells.push(['F' + r, `=IF(B${r}="","",IF(D${r}="",1,D${r})*E${r})`, MONEY], ['A' + r, null, { align: { h: 'center' }, font: { b: true, color: M.rgb('#1F8A4C') } }]);
          if (r >= 5 + items.length) cells.push(['E' + r, null, MONEY]);
        }
        return { sheets: [{ name: 'Shopping', cells, cols: { A: 5, B: 28, C: 14, D: 7, E: 12, F: 12 }, freeze: [4, 0],
          dv: [{ ref: 'A5:A44', type: 'list', f1: '"✓"', prompt: '勾选已放入购物篮的商品', promptTitle: '已购买' },
            { ref: 'C5:C44', type: 'list', f1: '"果蔬,乳制品,烘焙,肉类与水产,干货,冷冻,饮料,日用品,其他"' }],
          sel: 'B13' }] };
      },
    },
    schedule: {
      name: '每周日程',
      spec: () => {
        const days = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
        const cols = 'BCDEFGH';
        const SLOT = Object.assign({ align: { v: 'center', wrap: true } }, BOX);
        const cells = [['A1', '每周日程', TITLE], ['A2', '本周起始', B], ['B2', '=TODAY()-WEEKDAY(TODAY(),3)', Object.assign({ nf: 'mmmm d, yyyy' }, INPUT)], ['A4', '时间', HDR]];
        days.forEach((d, i) => cells.push([cols[i] + '4', d, HDR], [cols[i] + '5', `=$B$2+${i}`, { nf: 'mmm d', align: { h: 'center' }, font: { color: M.rgb('#595959') } }]));
        for (let k = 0; k <= 15; k++) {
          const r = 6 + k;
          cells.push(['A' + r, (7 + k) / 24, { nf: 'h:mm AM/PM', font: { b: true }, align: { v: 'center' } }]);
          for (const c of cols) cells.push([c + r, null, SLOT]);
        }
        const plan = [['B8', '团队晨会'], ['D8', '团队晨会'], ['F8', '团队晨会'], ['C11', '与山姆午餐'], ['E17', '游泳'], ['G10', '购买食品'], ['H11', '家庭早午餐'], ['B19', '钢琴课']];
        for (const [a, v] of plan) cells.push([a, v, Object.assign({ fill: { pattern: 'solid', fg: M.rgb('#DDEBF7') } }, SLOT)]);
        cells.push(['A23', '备注', B], ['B23', '', INPUT]);
        return { sheets: [{ name: 'Schedule', cells, cols: { A: 10, B: 15, C: 15, D: 15, E: 15, F: 15, G: 15, H: 15 }, freeze: [5, 1], print: { orientation: 'landscape' }, sel: 'C6' }] };
      },
    },
  };
  P.TEMPLATES = TPL;
  P.buildTemplate = (k) => build(TPL[k].spec());
  P.newFromTemplate = function (k) {
    const t = TPL[k];
    if (!t) return;
    const wb = build(t.spec());
    if (L.cf && L.cf.reset) L.cf.reset();
    A().prepare(wb);
    A().addBook(wb, t.name.replace(/\s+/g, '') + '1', { untitled: true });
  };
})(typeof window !== 'undefined' ? window : globalThis);
