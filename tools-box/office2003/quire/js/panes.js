/* Quire — task panes (Getting Started, New Document, Help, Search, Clip Art, Clipboard, Research,
 * Styles and Formatting, Reveal Formatting, Mail Merge, Protect Document), the Document Map,
 * page thumbnails, Print Preview and the Reading Layout toolbar.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const P = (L.panes = {});
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= task pane shell ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    TP.back = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '后退', 'data-tip': '后退', html: L.icons.get('back') });
    TP.fwd = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '向前', 'data-tip': '向前', html: L.icons.get('forward') });
    TP.home = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '主页', 'data-tip': 'Home', html: L.icons.get('home') });
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': '关闭任务窗格', 'data-tip': '关闭', html: L.icons.get('close', 12) });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => A().toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
    L.bus.on('sel', L.debounce(() => TP.refresh('sel'), 150));
    L.bus.on('doc-changed', L.debounce(() => TP.refresh('doc'), 300));
    L.bus.on('doc-loaded', () => TP.refresh('load'));
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id]) return;
    if (L.$('#taskpane').hidden) A().toggleTask(true);
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
    else if (p.live && reason !== 'sel') { L.clear(TP.body); p.render(TP.body); }
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn, o) => h('button', Object.assign({ class: 'tp-link', type: 'button', onclick: fn }, o || {}), icon ? h('span', { html: L.icons.get(icon) }) : null, h('span', { text: label }));
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
        h('div', { class: 'tp-row' }, h('span', { html: L.icons.app ? L.icons.app(32) : '' }), h('b', { text: 'Quire 网络版', style: 'color:var(--pane-ink)' })),
        sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec(L.N('打开'), window.VO ? VO.recentLinks(4) : null, link('open', '更多...', () => ui.exec('open')), link('new', '新建文档...', () => TP.show('new')), link('sample', '打开欢迎文档', () => A().newDocument({ build: (d) => Object.assign(d, A().buildSample()) }))),
        sec(L.N('学习'), link('help', '键盘快捷键和帮助', () => { P.helpTopic = 'keys'; TP.show('help'); }), link('stylesPane', '样式和格式', () => TP.show('styles')), link('mailMerge', '邮件合并向导', () => TP.show('mailmerge'))));
    },
  });
  /* ---------- New Document ---------- */
  TP.register({
    id: 'new', title: '新建文档',
    render(b) {
      const tpl = L.templates ? L.templates.list() : [];
      b.append(
        sec(L.N('新建'), link('new', '空白文档', () => A().newDocument()), link('webPage', '网页', () => A().newDocument({ build: (d) => { d.webPage = true; } }).then(() => A().setView('web'))), link('mail', '电子邮件正文', () => A().newDocument({ build: (d) => { d.main.blocks = [D.para([D.text('To: ')]), D.para([D.text('Subject: ')]), D.para()]; } })), link('open', '根据现有文档...', () => A().newFromExisting())),
        sec(L.N('模板'), link('templates', '本机上的模板...', () => L.dlg.newFromTemplate()), ...tpl.slice(0, 6).map((t) => link('new', t.name, () => L.templates.create(t.id)))));
    },
  });
  /* ---------- Help ---------- */
  const HELP = [
    [L.N('键入和选择'), L.N('单击定位插入点并键入内容。双击选择单词，三击选择段落；在左页边距单击选择一行，双击选择段落，按住 Ctrl 单击选择整个文档。在文本中按住 Ctrl 单击可选择句子。按 F8 可逐步扩展选区（单词、句子、段落、文档），按 Esc 停止扩展。')],
    [L.N('设置文字格式'), L.N('使用“格式”工具栏或“格式 ▸ 字体”(Ctrl+D)。Ctrl+B、Ctrl+I、Ctrl+U 可切换加粗、倾斜和下划线；Ctrl+Space 清除字符格式，Ctrl+Q 清除段落格式。Shift+F3 循环切换大写、小写和每词首字母大写。格式刷用于复制格式：单击使用一次，双击保持启用，然后拖选文字或单击单词。')],
    [L.N('样式'), L.N('在“格式”工具栏左侧的“样式”框中选择样式，或打开“格式 ▸ 样式和格式”查看完整列表，并使用“修改”、“新建样式”和“选择所有实例”。在“样式”框中键入新名称，可根据当前段落创建样式。')],
    [L.N('段落、制表位和缩进'), L.N('“格式 ▸ 段落”可设置对齐、缩进、间距、行距和分页选项（与下段同页、孤行控制、段前分页）。拖动标尺上的标记可更改缩进；单击标尺可添加其左端方框中所示类型的制表位；将制表位拖出标尺可将其移除。')],
    [L.N('项目符号和编号'), L.N('使用工具栏按钮或“格式 ▸ 项目符号和编号”创建项目符号、编号或多级符号列表。在列表项开头按 Tab 或 Shift+Tab 可更改级别。键入“1. ”或“* ”及文字，再按 Enter，可自动创建列表。')],
    [L.N('表格'), L.N('使用“表格 ▸ 插入 ▸ 表格”，或在“常用”工具栏的“插入表格”网格中拖动。Tab 可移至下一单元格（在末尾会添加新行）。拖动单元格边框可调整列宽。“表格和边框”工具栏可绘制、擦除、合并和拆分单元格，设置边框、底纹和对齐方式，以及排序和求和。')],
    [L.N('图片和绘图'), L.N('通过“插入 ▸ 图片”选择“来自文件”、“剪贴画”、“自选图形”、“LettersArt”或“图表”。在“格式 ▸ 图片 ▸ 版式”或“文字环绕”按钮中选择环绕方式。拖动浮动对象可移动，拖动控点可调整大小，绿色控点可旋转形状。')],
    [L.N('页眉、页脚和页码'), L.N('使用“视图 ▸ 页眉和页脚”（或双击页面顶部或底部）。“页眉和页脚”工具栏可插入页码、总页数、日期和时间，在页眉和页脚之间切换，以及链接或取消链接各节。“插入 ▸ 页码”可添加带页码的框架。')],
    [L.N('分节、分栏和页面设置'), L.N('“插入 ▸ 分隔符”可添加分页符、分栏符和分节符。每节可设置独立的页边距、方向、纸张大小、分栏和页眉。“文件 ▸ 页面设置”可为当前节或整个文档设置页边距、纸张及版式（首页不同、奇偶页不同、垂直对齐、行号）。')],
    [L.N('引用'), L.N('“插入 ▸ 引用”提供“脚注”（Ctrl+Alt+F 为脚注，Ctrl+Alt+D 为尾注）、“题注”、“交叉引用”和“索引和目录”（目录、图表目录、索引）。右键单击域并选择“更新域”，或按 F9。Alt+F9 可显示域代码。')],
    [L.N('修订和批注'), L.N('“工具 ▸ 修订”(Ctrl+Shift+E) 标记插入、删除和格式更改。“审阅”工具栏可接受、拒绝修订并在修订之间移动。“插入 ▸ 批注”可在页面视图中添加批注框。“工具 ▸ 比较并合并文档”可标记与另一份 .docx 的差异。')],
    [L.N('查找和替换'), L.N('“编辑 ▸ 查找”(Ctrl+F)、“替换”(Ctrl+H) 和“定位”(Ctrl+G)。更多选项包括：区分大小写、全字匹配、通配符（? 表示任意字符，* 表示任意文本，[abc] 表示字符集合，< 表示词首，> 表示词尾）、同音词、单词的所有形式、格式和特殊字符（^p 段落标记、^t 制表符、^m 手动分页符、^l 换行符、^? 任意字符、^# 任意数字、^$ 任意字母）。')],
    [L.N('邮件合并'), L.N('“工具 ▸ 信函与邮件 ▸ 邮件合并”打开六步向导：选择信函、信封、标签或目录；打开 .csv、.txt 或 .docx 数据源，或键入新列表；插入地址块、问候语和合并域；预览各条记录；最后合并到新文档或 PDF。')],
    [L.N('文件'), L.N('“文件 ▸ 打开”可读取 .docx、.docm、.dotx、.dotm、.htm、.txt、.rtf 和 Markdown 文件（.md，或包含 .md 及其图片的 .zip）。“文件 ▸ 保存”可生成能在 Word、LibreOffice、Pages 和 Google Docs 中打开的 .docx，或将 Markdown 文件继续保存为 Markdown（未编辑的内容保持原样）。“另存为”还提供 Word 模板 (.dotx)、网页、RTF、纯文本、Markdown 和 PDF 格式。“文件 ▸ 打印”会将所有页面生成为 PDF，供您在任何地方打印。')],
    [L.N('视图'), L.N('普通视图将文字显示为连续的一栏，并以虚线表示分页符。Web 版式根据窗口宽度换行。页面视图按实际打印效果显示页面，包括页眉、页脚和批注框。大纲视图显示标题级别，便于升级、降级和重新排列。阅读版式并排显示两页以便阅读。打印预览 (Ctrl+F2) 显示一页或多页的完整页面。')],
  ];
  const KEYS = [['Ctrl+N / O / S / P', L.N('新建、打开、保存、打印')], ['F12 / Shift+F12', L.N('另存为/保存')], ['Ctrl+Z / Ctrl+Y', L.N('撤消/重做（F4 重复操作）')], ['Ctrl+X / C / V', L.N('剪切、复制、粘贴')], ['Ctrl+A', L.N('全选')], ['Ctrl+F / H / G', L.N('查找、替换、定位')], ['Ctrl+B / I / U', L.N('加粗、倾斜、下划线')], ['Ctrl+Shift+D / W', L.N('双下划线/字下划线')], ['Ctrl+= / Ctrl+Shift++', L.N('下标/上标')], ['Ctrl+Shift+A / K', L.N('全部大写/小型大写字母')], ['Ctrl+Shift+> / <', L.N('增大/减小字号')], ['Ctrl+] / [', L.N('字号增减 1 磅')], ['Ctrl+D', L.N('字体对话框')], ['Ctrl+Space / Ctrl+Q', L.N('清除字符/段落格式')], ['Shift+F3', L.N('更改大小写')], ['Ctrl+L / E / R / J', L.N('左对齐、居中、右对齐、两端对齐')], ['Ctrl+1 / 2 / 5', L.N('单倍、双倍、1.5 倍行距')], ['Ctrl+0', L.N('添加/删除段前 12 磅间距')], ['Ctrl+M / Ctrl+Shift+M', L.N('增加/减少缩进')], ['Ctrl+T / Ctrl+Shift+T', L.N('悬挂缩进/减少悬挂缩进')], ['Ctrl+Shift+N', L.N('正文样式')], ['Ctrl+Alt+1 / 2 / 3', L.N('标题 1、2、3')], ['Ctrl+Shift+L', L.N('项目符号列表样式')], ['Ctrl+Shift+S', L.N('样式框')], ['Ctrl+Enter', L.N('分页符')], ['Ctrl+Shift+Enter', L.N('分栏符')], ['Shift+Enter', L.N('换行符')], ['Ctrl+Shift+Space', L.N('不间断空格')], ['Ctrl+- / Ctrl+Shift+-', L.N('可选连字符/不间断连字符')], ['Ctrl+Alt+- ', L.N('长划线')], ['Ctrl+Alt+C / R / T', L.N('©, ®, ™')], ['Ctrl+K', L.N('超链接')], ['Ctrl+Alt+F / D', L.N('脚注/尾注')], ['Ctrl+Alt+M', L.N('批注')], ['Ctrl+Shift+E', L.N('修订')], ['F7', L.N('拼写和语法')], ['Shift+F7', L.N('同义词库')], ['F9 / Alt+F9', L.N('更新域/切换域代码')], ['Ctrl+Shift+F9', L.N('断开域链接')], ['Alt+Shift+D / T / P', L.N('日期、时间、页码域')], ['F8', L.N('扩展选区')], ['Shift+F5', L.N('返回上次编辑位置')], ['Alt+Shift+↑ / ↓', L.N('上移/下移段落')], ['Alt+Shift+← / →', L.N('升级/降级')], ['Ctrl+Alt+P / N / O', L.N('页面视图、普通视图、大纲视图')], ['Ctrl+*', L.N('显示/隐藏 ¶')], ['Ctrl+F1 / Shift+F1', L.N('任务窗格/显示格式')], ['F1', L.N('帮助')], ['Tab / Shift+Tab', L.N('下一个/上一个单元格；列表级别')]];
  TP.register({
    id: 'help', title: 'Quire 帮助',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: '搜索', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        L.clear(res);
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        /* the topics as shown (and in English) */
        if (P.helpTopic !== 'keys' || words.length) for (const [t, d] of HELP.filter(([t, d]) => !words.length || hit(t + ' ' + d + ' ' + L.t(t) + ' ' + L.t(d)))) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        const keys = KEYS.filter(([k, d]) => !words.length || hit(k + ' ' + d + ' ' + L.t(d)));
        /* the keys with their names as shown (Strg in German) */
        const keyNames = (k) => k.replace(/\b(Ctrl|Shift|Alt|Enter|Tab|Space)\b/g, (m) => ui.keyText(m));
        if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: '键盘快捷键' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { translate: 'no' }, document.createTextNode(keyNames(k))), h('td', { text: d }))))));
        if (!res.childNodes.length) res.appendChild(h('div', { class: 'tp-note', text: '没有与搜索内容匹配的帮助主题。' }));
        if (P.helpTopic === 'keys' && !words.length) res.appendChild(link('help', '显示所有帮助主题', () => { P.helpTopic = null; draw(); }));
      };
      q.addEventListener('input', draw);
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
      P.helpQuery = '';
    },
  });
  /* ---------- Search (in this document) ---------- */
  TP.register({
    id: 'search', title: '基本文件搜索',
    render(b) {
      const q = h('input', { type: 'text', id: 'srch-q', value: P.searchQuery || '', style: 'flex:1;min-width:0' });
      const res = h('div');
      const go = () => {
        P.searchQuery = q.value;
        L.clear(res);
        const needle = q.value.trim().toLowerCase();
        if (!needle) return;
        let n = 0;
        D.reindex(doc());
        for (const st of D.stories(doc())) for (const p of st.paras || []) {
          const t = D.ptext(p);
          let i = t.toLowerCase().indexOf(needle);
          while (i >= 0 && n < 200) {
            n++;
            const a = Math.max(0, i - 30), e = Math.min(t.length, i + needle.length + 40);
            const hit = h('div', { class: 'search-hit', translate: 'no' });
            hit.append(document.createTextNode((a > 0 ? '…' : '') + t.slice(a, i)), h('b', null, document.createTextNode(t.slice(i, i + needle.length))), document.createTextNode(t.slice(i + needle.length, e) + (e < t.length ? '…' : '')));
            const off = i;
            hit.addEventListener('click', () => A().gotoPos(D.pos(p, off), D.pos(p, off + needle.length)));
            res.appendChild(hit);
            i = t.toLowerCase().indexOf(needle, i + needle.length);
          }
        }
        res.prepend(h('div', { class: 'tp-note', text: n === 1 ? L.t('{1} 中有 {0} 个匹配项', n, A().fileName) : n ? L.t('{1} 中有 {0} 个匹配项', n, A().fileName) : '在此文档中未找到匹配项。' }));
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); e.stopPropagation(); });
      b.append(sec('搜索内容', h('div', { class: 'tp-row' }, h('label', { for: 'srch-q', text: '搜索文字：' })), h('div', { class: 'tp-row' }, q, ui.button('转到', go, { class: 'btn small' })), h('div', { class: 'tp-note', text: '搜索已打开文档的文字，包括页眉、页脚、注释和批注。' })), res);
      if (q.value) go();
    },
  });
  /* ---------- Clip Art ---------- */
  TP.register({
    id: 'clipart', title: '剪贴画',
    render(b) {
      const q = h('input', { type: 'text', id: 'ca-q', value: P.clipQuery || '', style: 'flex:1;min-width:0' });
      const grid = h('div', { class: 'clip-grid' });
      const draw = () => {
        L.clear(grid);
        const items = L.clipart ? L.clipart.search(q.value) : [];
        for (const it of items) {
          const el = h('button', { class: 'clip-it', type: 'button', 'data-tip': it.name, 'aria-label': it.name }, h('img', { src: L.clipart.url(it), alt: '' }), h('span', { text: it.name }));
          el.addEventListener('click', () => P.insertClip(it));
          grid.appendChild(el);
        }
        if (!items.length) grid.appendChild(h('div', { class: 'tp-note', text: '没有匹配的剪贴画。请尝试搜索“people”、“chart”、“nature”或“office”。' }));
      };
      q.addEventListener('input', () => { P.clipQuery = q.value; draw(); });
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', draw, { class: 'btn small' }))), grid, h('div', { class: 'tp-note', text: '单击剪贴画可将其插入光标位置。' }));
      draw();
    },
  });
  P.insertClip = async function (it) {
    const blob = await L.clipart.toPNG(it, 512);
    const id = L.media.add(blob, it.name.replace(/\W+/g, '_') + '.png');
    E.edit('插入剪贴画', () => O.insertItem(E.deleteSelection(), D.item('img', { media: id, w: 96, h: 96, name: it.name, alt: it.name })));
  };
  /* ---------- Office Clipboard ---------- */
  TP.register({
    id: 'clipboard', title: '剪贴板', live: true,
    render(b) {
      const items = A().clipItems || [];
      b.append(h('div', { class: 'tp-btns' }, ui.button('全部粘贴', () => { for (const it of items.slice().reverse()) E.pasteBlocks(it.blocks); }, { class: 'btn small' }), ui.button('全部清除', () => { items.length = 0; TP.refresh('doc'); }, { class: 'btn small' })), h('div', { class: 'tp-note', text: L.t('剪贴板：{0}/24。单击项目可粘贴。', items.length) }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { html: L.icons.get(it.blocks.some((x) => x.t === 'tbl') ? 'table' : 'font') }), h('span', { class: 'cb-t', translate: 'no' }, document.createTextNode((it.text || '').slice(0, 160) || L.t('[对象]'))));
        el.addEventListener('click', () => E.pasteBlocks(it.blocks));
        b.appendChild(el);
      }
    },
  });
  /* ---------- Research: offline thesaurus (WordNet-based) and searching this document ---------- */
  const TH = (L.thesaurus = { data: null, keys: null, loading: null });
  TH.load = function () {
    if (TH.data) return Promise.resolve(TH.data);
    if (TH.loading) return TH.loading;
    TH.loading = fetch('../common/dict/en.thes').then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }).then((txt) => {
      const map = new Map();
      for (const ln of txt.split('\n')) { if (!ln) continue; const k = ln.slice(0, ln.indexOf('|')); map.set(k.toLowerCase(), ln); }
      TH.data = map;
      TH.keys = Array.from(map.keys()).sort();
      return map;
    }).catch((e) => { TH.loading = null; throw e; });
    return TH.loading;
  };
  const POS = { noun: 'n.', verb: 'v.', adj: 'adj.', adv: 'adv.' };
  /** entry for a word or its base form: {head, senses:[{pos, syns}], ants} */
  TH.lookup = function (word) {
    const m = TH.data;
    if (!m) return null;
    const w = String(word).trim().toLowerCase().replace(/’/g, "'");
    const tries = [w];
    const add = (x) => { if (x && x.length > 1 && !tries.includes(x)) tries.push(x); };
    add(w.replace(/'s$/, ''));
    add(w.replace(/ies$/, 'y')); add(w.replace(/es$/, '')); add(w.replace(/s$/, ''));
    add(w.replace(/ied$/, 'y')); add(w.replace(/ed$/, '')); add(w.replace(/d$/, '')); add(w.replace(/([b-df-hj-np-tv-z])\1ed$/, '$1'));
    add(w.replace(/ing$/, '')); add(w.replace(/ing$/, 'e')); add(w.replace(/([b-df-hj-np-tv-z])\1ing$/, '$1'));
    add(w.replace(/ily$/, 'y')); add(w.replace(/ly$/, ''));
    add(w.replace(/ier$/, 'y')); add(w.replace(/er$/, '')); add(w.replace(/iest$/, 'y')); add(w.replace(/est$/, ''));
    for (const t of tries) {
      const ln = m.get(t);
      if (!ln) continue;
      const parts = ln.split('|');
      return {
        head: parts[0], base: t !== w,
        senses: parts[1].split(';').map((x) => { const k = x.indexOf(':'); return { pos: x.slice(0, k), syns: x.slice(k + 1).split(',').filter(Boolean) }; }),
        ants: parts[2] ? parts[2].split(',').filter(Boolean) : [],
      };
    }
    return null;
  };
  /** nearby headwords for the alphabetical list Word shows when nothing matches */
  TH.near = function (word, n) {
    const keys = TH.keys || [];
    const w = String(word).toLowerCase();
    let lo = 0, hi = keys.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (keys[mid] < w) lo = mid + 1; else hi = mid; }
    return keys.slice(Math.max(0, lo - Math.floor(n / 2)), Math.max(0, lo - Math.floor(n / 2)) + n);
  };
  /* Embedded copy (tools-box): only 'document' is left — the thesaurus needs common/dict/en.thes, which
     this site does not ship (see the note in the sync header of src/pages/Office2003/index.jsx). */
  const SERVICES = [['document', '搜索本文档']];
  TP.register({
    id: 'research', title: '信息检索',
    render(b) {
      const q = h('input', { type: 'text', id: 'rs-q', value: P.researchQuery || '', style: 'flex:1;min-width:0', 'aria-label': '搜索内容' });
      const svc = ui.select(SERVICES, SERVICES.some((x) => x[0] === P.researchService) ? P.researchService : 'document', (v) => { P.researchService = v; }, { id: 'rs-svc', style: 'flex:1;min-width:0' });
      const hist = P.researchHistory || (P.researchHistory = []);
      const out = h('div');
      const go = async (text, fromHistory) => {
        text = (text == null ? q.value : text).trim();
        if (!text) return;
        q.value = text;
        P.researchQuery = text;
        if (!fromHistory && hist[hist.length - 1] !== text) hist.push(text);
        L.clear(out);
        const service = svc.value;
        if (service === 'thesaurus' || service === 'thesaurusUK' || service === 'all') {
          const status = h('div', { class: 'tp-note', text: '正在搜索...' });
          out.appendChild(status);
          try { await TH.load(); }
          catch (e) { status.textContent = L.t('无法加载同义词库（{0}）。请通过 Web 服务器访问 Quire，不要直接打开本地 index.html 文件。', e.message || e); return; }
          status.remove();
          renderThesaurus(out, text, (w) => go(w));
        }
        if (service === 'document' || service === 'all') renderDocMatches(out, text);
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); e.stopPropagation(); });
      const back = ui.button('◄ 返回', () => { if (hist.length > 1) { hist.pop(); go(hist[hist.length - 1], true); } }, { class: 'btn small' });
      b.append(sec('搜索内容：', h('div', { class: 'tp-row' }, q, ui.button('转到', () => go(), { class: 'btn small' })), h('div', { class: 'tp-row' }, svc), h('div', { class: 'tp-row' }, back)), out);
      if (q.value && P.researchAuto !== false) { P.researchAuto = false; setTimeout(() => go(), 0); }
    },
  });
  function wordMenuButton(w, onLookUp) {
    const btn = h('button', { class: 'tp-link th-word', type: 'button', translate: 'no' }, h('span', null, document.createTextNode(w)), h('span', { class: 'dd-arrow' }));
    btn.addEventListener('click', (e) => {
      ui.openMenu([
        { label: '插入(&I)', run: () => { if (E.sel && E.collapsed()) { const wd = O.wordAt(E.sel.f); if (wd) E.setSel(wd[0], wd[1]); } E.insertTextAt(w, 'Insert'); } },
        { label: '复制(&C)', run: () => { try { navigator.clipboard.writeText(w).catch(() => {}); } catch (err) { /* ignore */ } E.clip = { text: w, blocks: [D.para([D.text(w)])], html: w }; } },
        { label: '查阅(&L)', run: () => onLookUp(w) },
      ], btn.getBoundingClientRect());
      e.stopPropagation();
    });
    btn.addEventListener('dblclick', () => onLookUp(w));
    return btn;
  }
  function renderThesaurus(out, text, onLookUp) {
    const r = TH.lookup(text);
    const box = h('div', { class: 'help-topic' }, h('h4', { text: '同义词库：英语' }));
    out.appendChild(box);
    if (!r) {
      box.appendChild(h('p', { class: 'tp-note', text: L.t('未找到“{0}”的结果。按字母顺序列出：', text) }));
      for (const k of TH.near(text, 12)) box.appendChild(wordMenuButton(k, onLookUp));
      return;
    }
    if (r.base) box.appendChild(h('p', { class: 'tp-note', text: L.t('“{0}”的搜索结果。', r.head) }));
    for (const s of r.senses) {
      /* the English thesaurus's own word and part of speech */
      const head = h('div', { class: 'th-sense', translate: 'no' }, h('b', null, document.createTextNode(`${s.syns[0]} (${POS[s.pos] || s.pos})`)));
      box.appendChild(head);
      for (const w of s.syns) box.appendChild(wordMenuButton(w, onLookUp));
    }
    if (r.ants.length) {
      box.appendChild(h('div', { class: 'th-sense' }, h('b', { text: L.t('{0}（反义词）', r.ants[0]) })));
      for (const w of r.ants) box.appendChild(wordMenuButton(w, onLookUp));
    }
  }
  function renderDocMatches(out, text) {
    const d = doc();
    D.reindex(d);
    const box = h('div', { class: 'help-topic' }, h('h4', { text: '本文档' }));
    out.appendChild(box);
    const needle = text.toLowerCase();
    let n = 0;
    for (const st of D.stories(d)) for (const p of st.paras || []) {
      const t = D.ptext(p);
      let i = t.toLowerCase().indexOf(needle);
      while (i >= 0 && n < 50) {
        const a = i, bEnd = i + needle.length;
        const snip = (a > 30 ? '…' : '') + t.slice(Math.max(0, a - 30), a);
        const hit = h('button', { class: 'search-hit', type: 'button', translate: 'no' }, document.createTextNode(snip.replace(/[￼\t]/g, ' ')), h('b', null, document.createTextNode(t.slice(a, bEnd))), document.createTextNode(t.slice(bEnd, bEnd + 40).replace(/[￼\t]/g, ' ') + (t.length > bEnd + 40 ? '…' : '')));
        hit.addEventListener('click', () => A().gotoPos(D.pos(p, a), D.pos(p, bEnd)));
        box.appendChild(hit);
        n++;
        i = t.toLowerCase().indexOf(needle, bEnd);
      }
    }
    if (!n) box.appendChild(h('p', { class: 'tp-note', text: '在此文档中未找到该文字。' }));
    else if (n >= 50) box.appendChild(h('p', { class: 'tp-note', text: '显示前 50 个匹配项。' }));
  }
  /* ---------- Styles and Formatting ---------- */
  TP.register({
    id: 'styles', title: '样式和格式',
    render(b) {
      const d = doc();
      const cur = h('div', { class: 'tp-box', style: 'min-height:22px;font-size:12px' });
      const list = h('div', { class: 'sf-list' });
      const show = ui.select([['avail', '可用格式设置'], ['inuse', '正在使用的格式'], ['styles', '可用样式'], ['all', '所有样式']], P.styleShow || 'avail', (v) => { P.styleShow = v; draw(); }, { id: 'sf-show', style: 'flex:1' });
      const draw = () => {
        L.clear(list);
        const used = D.usedStyles(d);
        const mode = P.styleShow || 'avail';
        let ids = Object.keys(d.styles).filter((k) => ['paragraph', 'character', 'table', 'numbering'].includes(d.styles[k].type) && !d.styles[k].hidden);
        if (mode === 'inuse') ids = ids.filter((k) => used.has(k));
        else if (mode === 'avail' || mode === 'styles') ids = ids.filter((k) => used.has(k) || d.styles[k].q || d.styles[k].custom || /^Heading[1-3]$|^Normal$/.test(k));
        if (mode === 'all') { const b0 = D.builtinStyles(); for (const k in b0) if (!ids.includes(k)) ids.push(k); }
        ids.sort((a, b2) => D.styleDisplayName(d.styles[a] || D.builtinStyles()[a]).localeCompare(D.styleDisplayName(d.styles[b2] || D.builtinStyles()[b2])));
        list.appendChild(h('div', { class: 'sf-item', role: 'button', tabindex: '0', onclick: () => ui.exec('clearFormats') }, h('span', { class: 'sf-n', text: '清除格式' })));
        const curId = E.sel ? E.sel.f.p.pPr.style || 'Normal' : '';
        for (const id of ids) {
          const s = d.styles[id] || D.builtinStyles()[id];
          if (!s) continue;
          const r = s.type === 'paragraph' ? (d.styles[id] ? D.styleProps(d, id).rPr : s.rPr || {}) : s.type === 'character' ? D.mergeRPr(d.defaults.rPr, d.styles[id] ? D.charStyleProps(d, id) : s.rPr) : d.defaults.rPr;
          const css = `font-family:${L.fontStack(r.font)};font-size:${Math.min(16, Math.max(8, (r.sz || 12)))}px;${r.b ? 'font-weight:bold;' : ''}${r.i ? 'font-style:italic;' : ''}${r.color && r.color !== 'auto' ? 'color:#' + r.color + ';' : ''}${r.u && r.u !== 'none' ? 'text-decoration:underline;' : ''}`;
          /* the style's own name (not translated); its description as shown in the tip */
          const it = h('div', { class: 'sf-item' + (id === curId ? ' on' : ''), role: 'button', tabindex: '0', translate: 'no', 'data-style': id }, h('span', { class: 'sf-n', style: css }, document.createTextNode(D.styleDisplayName(s))), h('span', { class: 'sf-k', text: s.type === 'paragraph' ? '¶' : s.type === 'character' ? 'a' : s.type === 'table' ? '⊞' : '≡' }));
          it.dataset.tip = styleDescription(d, id);
          it.addEventListener('click', () => { if (s.type === 'table') { if (L.tables.inTable()) L.tables.applyStyle(s.name || id); return; } if (s.type === 'numbering') return; A().applyStyle(id); E.refocus(); });
          it.addEventListener('contextmenu', (e) => { e.preventDefault(); styleMenu(id, { x: e.clientX, y: e.clientY }); });
          list.appendChild(it);
        }
      };
      P.refreshCur = () => { cur.textContent = A().styleState ? A().styleState() + describeDirect() : ''; };
      cur.setAttribute('translate', 'no');   // the style's name, then its direct formatting as shown
      b.append(
        sec('所选文字的格式', cur, h('div', { class: 'tp-btns' }, ui.button('全选', () => selectAllLike(), { class: 'btn small' }), ui.button('新建样式...', () => L.dlg.style(null), { class: 'btn small' }))),
        sec('选择要应用的格式', list),
        h('div', { class: 'tp-row' }, h('label', { for: 'sf-show', text: '显示：' }), show));
      draw();
      P.refreshCur();
      this._draw = draw;
    },
    refresh(b, reason) { if (P.refreshCur) P.refreshCur(); if (reason !== 'sel' && this._draw) this._draw(); else if (reason === 'sel') { const curId = E.sel ? E.sel.f.p.pPr.style || 'Normal' : ''; L.$$('.sf-item', b).forEach((x) => x.classList.toggle('on', x.dataset.style === curId)); } },
  });
  function describeDirect() {
    if (!E.sel) return '';
    const it = D.itemAfter(E.sel.f.p, E.sel.f.o) || D.itemBefore(E.sel.f.p, E.sel.f.o);
    const r = it && it.rPr ? it.rPr : {};
    const parts = [];
    if (r.font) parts.push(r.font);
    if (r.sz) parts.push(r.sz + ' pt');
    if (r.b) parts.push(L.t('加粗'));
    if (r.i) parts.push(L.t('倾斜'));
    if (r.u && r.u !== 'none') parts.push(L.t('下划线'));
    if (r.color && r.color !== 'auto') parts.push(L.t('字体颜色：#{0}', r.color));
    const pp = E.sel.f.p.pPr;
    if (pp.jc) parts.push(L.t(JC[pp.jc] || pp.jc));
    if (pp.ind && (pp.ind.l || pp.ind.fl)) parts.push(L.t('缩进'));
    return parts.length ? ' + ' + parts.join(', ') : '';
  }
  /* paragraph alignments as shown */
  const JC = { center: L.N('居中'), right: L.N('右'), both: L.N('两端对齐'), left: L.N('左'), distribute: L.N('分散对齐') };
  function styleDescription(d, id) {
    const s = d.styles[id] || D.builtinStyles()[id];
    if (!s) return id;
    const r = s.rPr || {};
    const p = s.pPr || {};
    const parts = [];
    if (s.basedOn) parts.push((d.styles[s.basedOn] ? D.styleDisplayName(d.styles[s.basedOn]) : s.basedOn) + ' +');
    if (r.font) parts.push(r.font);
    if (r.sz) parts.push(r.sz + ' pt');
    if (r.b) parts.push(L.t('加粗'));
    if (r.i) parts.push(L.t('倾斜'));
    if (p.jc) parts.push(L.t(JC[p.jc] || p.jc));
    if (p.sp && p.sp.b) parts.push(L.t('段前：{0} 磅', p.sp.b));
    if (p.sp && p.sp.a) parts.push(L.t('段后：{0} 磅', p.sp.a));
    if (p.keepNext) parts.push(L.t('与下段同页'));
    if (p.outline != null && p.outline < 9) parts.push(L.t('级别 {0}', p.outline + 1));
    return parts.join(', ') || D.styleDisplayName(s);
  }
  function styleMenu(id, at) {
    const d = doc();
    const s = d.styles[id];
    ui.openMenu([
      { label: L.t('选择全部 {0} 处(&S)', countUses(id)), enabled: () => countUses(id) > 0, run: () => selectAllOfStyle(id) },
      { label: L.t('清除 {0} 处的格式(&C)', countUses(id)), enabled: () => countUses(id) > 0, run: () => { selectAllOfStyle(id); ui.exec('clearFormats'); } },
      '-',
      { label: '更新以匹配所选内容(&U)', enabled: () => !!s && !!E.sel, run: () => updateStyleFromSel(id) },
      { label: '修改(&M)...', run: () => L.dlg.style(id) },
      { label: '删除(&D)...', enabled: () => !!s && !s.builtin && id !== 'Normal', run: async () => { const r = await ui.msg(L.t('要删除样式 {0} 吗？', D.styleDisplayName(s)), { icon: 'question', buttons: ['是(&Y)', '否(&N)'] }); if (r === 0) deleteStyle(id); } },
    ], at);
  }
  function countUses(id) {
    let n = 0;
    D.walk(doc().main, (b) => { if (b.t === 'p') { if ((b.pPr.style || 'Normal') === id) n++; for (const it of b.runs) if (it.rPr && it.rPr.style === id) { n++; break; } } });
    return n;
  }
  function selectAllOfStyle(id) {
    const d = doc();
    D.reindex(d);
    const ps = d.main.paras.filter((p) => (p.pPr.style || 'Normal') === id || p.runs.some((it) => it.rPr && it.rPr.style === id));
    if (!ps.length) return;
    E.setSel(D.pos(ps[0], 0), D.pos(ps[ps.length - 1], D.plen(ps[ps.length - 1])));
    A().status(L.t('{0} 个段落使用 {1}；选区涵盖所有这些段落。', ps.length, D.styleDisplayName(d.styles[id])));
  }
  function selectAllLike() { if (!E.sel) return; selectAllOfStyle(E.sel.f.p.pPr.style || 'Normal'); }
  function updateStyleFromSel(id) {
    const d = doc();
    const p = E.sel.f.p;
    const it = D.itemAfter(p, E.sel.f.o) || D.itemBefore(p, E.sel.f.o);
    E.edit('更新样式', () => {
      D.touchKey(d, 'styles');
      const s = d.styles[id];
      if (s.type === 'paragraph') {
        const pp = L.clone(p.pPr); delete pp.style; delete pp.num; delete pp.sect; delete pp.chg;
        s.pPr = D.mergePPr(s.pPr || {}, pp);
        const rr = it && it.rPr ? L.clone(it.rPr) : {}; delete rr.style; delete rr.link; delete rr.ins; delete rr.del; delete rr.chg;
        s.rPr = D.mergeRPr(s.rPr || {}, rr);
        /* direct formatting now lives in the style */
        D.walk(d.main, (b) => { if (b.t === 'p' && (b.pPr.style || 'Normal') === id) { D.touch(b); for (const k in pp) delete b.pPr[k]; for (const x of b.runs) if (x.rPr) for (const k in rr) if (JSON.stringify(x.rPr[k]) === JSON.stringify(rr[k])) delete x.rPr[k]; D.normalize(b); } });
      } else if (s.type === 'character' && it && it.rPr) { const rr = L.clone(it.rPr); delete rr.style; s.rPr = D.mergeRPr(s.rPr || {}, rr); }
      D.stylesChanged();
      return E.sel;
    });
  }
  function deleteStyle(id) {
    const d = doc();
    E.edit('删除样式', () => {
      D.touchKey(d, 'styles');
      const s = d.styles[id];
      delete d.styles[id];
      for (const k in d.styles) if (d.styles[k].basedOn === id) d.styles[k].basedOn = s.basedOn || (s.type === 'paragraph' ? 'Normal' : null);
      for (const st of D.stories(d)) D.walk(st, (b) => { if (b.t === 'p') { if (b.pPr.style === id) { D.touch(b); delete b.pPr.style; } for (const it of b.runs) if (it.rPr && it.rPr.style === id) { D.touch(b); delete it.rPr.style; } } if (b.t === 'tbl' && b.tblPr.style === id) { D.touchTbl(b); delete b.tblPr.style; } });
      D.stylesChanged();
      return E.sel;
    });
  }
  P.styleMenu = styleMenu;
  /* ---------- Reveal Formatting ---------- */
  TP.register({
    id: 'reveal', title: '显示格式',
    render(b) {
      const box = h('div', { class: 'rf-tree' });
      b.append(sec('所选文字', h('div', { class: 'tp-box', translate: 'no' }, document.createTextNode(E.sel ? (O.textRange(...E.range(), ' ').slice(0, 120) || L.t('示例文字')) : ''))), sec('所选文字的格式', box), ui.check('区分样式来源', !!P.distinguish, (v) => { P.distinguish = v; TP.render(); }), ui.check('显示所有格式标记', !!LY.opts.showMarks, () => ui.exec('showMarks')));
      if (!E.sel) return;
      const d = doc();
      const p = E.sel.f.p;
      const r = E.curRun();
      const pp = E.curPara();
      const sect = D.sectFor(d, p);
      /* k: the label as shown, with its colon; v: the value as shown (translated here, or the document's own: names, fonts) */
      const row = (k, v, dlgFn) => { const el = h('div', null, h('span', { text: k }), ' ', h('span', { translate: 'no' }, document.createTextNode(v))); if (dlgFn) { el.style.cursor = 'pointer'; el.addEventListener('click', dlgFn); } return el; };
      const fl = (v) => D.fmtLen(v || 0);
      const head = (t, fn) => { const el = h('h5', null, h('a', { href: '#', text: t, onclick: (e) => { e.preventDefault(); fn(); } })); return el; };
      const ind = pp.ind || {}, sp = pp.sp || {};
      const lineSp = !sp.line ? L.t('单线') : sp.rule === 'auto' || !sp.rule ? (sp.line === 1 ? L.t('单线') : sp.line === 2 ? L.t('双线') : L.t('{0} 行', sp.line)) : sp.rule === 'exact' ? L.t('固定值 {0} 磅', sp.line) : L.t('最小值 {0} 磅', sp.line);
      const ul = (L.dlg.UNDERLINES || []).find((x) => x[0] === r.u);
      const START = { nextPage: L.N('新页'), continuous: L.N('连续'), evenPage: L.N('偶数页'), oddPage: L.N('奇数页'), nextColumn: L.N('新建列') };
      box.append(...[head(L.N('字体'), () => L.dlg.font()), row(L.N('字体：'), [`${r.font || ''} ${r.sz || 12} pt`, r.b && L.t('加粗'), r.i && L.t('倾斜')].filter(Boolean).join(', ')), r.color && r.color !== 'auto' ? row('字体颜色：', '#' + r.color) : null, r.u && r.u !== 'none' ? row('下划线：', L.t(ul ? ul[1] : r.u)) : null, r.hl ? row('突出显示：', L.t(ui.colorName(L.R.HIGHLIGHT[r.hl] || r.hl))) : null, r.caps ? row('效果：', L.t('全部大写')) : null, r.smallCaps ? row('效果：', L.t('小型大写字母')) : null, r.vert ? row('位置：', L.t({ superscript: L.N('上标'), subscript: L.N('下标') }[r.vert] || r.vert)) : null, row('语言：', L.t(r.lang ? A().langName(r.lang) : '英语（美国）')),
        head(L.N('段落'), () => L.dlg.paragraph()), row('段落样式：', D.styleDisplayName(d.styles[p.pPr.style || 'Normal'])), row('对齐方式：', L.t(JC[pp.jc || 'left'] || pp.jc)), row('缩进：', !ind.fl ? L.t('左：{0}，右：{1}', fl(ind.l), fl(ind.r)) : ind.fl < 0 ? L.t('左：{0}，右：{1}，悬挂：{2}', fl(ind.l), fl(ind.r), fl(-ind.fl)) : L.t('左：{0}，右：{1}，首行：{2}', fl(ind.l), fl(ind.r), fl(ind.fl))),
        row('间距：', L.t('段前：{0} 磅，段后：{1} 磅，行距：{2}', L.round(sp.b || 0, 1), L.round(sp.a || 0, 1), lineSp)),
        pp.keepNext || pp.keepLines || pp.pageBreakBefore ? row('换行和分页：', [pp.keepNext && L.t('与下段同页'), pp.keepLines && L.t('段中不分页'), pp.pageBreakBefore && L.t('段前分页')].filter(Boolean).join(', ')) : null,
        L.lists && L.lists.info(p) ? row('项目符号和编号：', L.lists.info(p).fmt === 'bullet' ? L.t('项目符号列表') : L.t('编号列表')) : null,
        head(L.N('节'), () => L.dlg.pageSetup()), row('页边距：', L.t('上：{0}，下：{1}，左：{2}，右：{3}', fl(sect.mt), fl(sect.mb), fl(sect.ml), fl(sect.mr))), row('布局：', L.t('节的起始位置：{0}', L.t(START[sect.type || 'nextPage']))), row('纸张：', `${L.t(D.paperName(sect))} (${fl(sect.pgW)} × ${fl(sect.pgH)}), ${sect.orient === 'landscape' ? L.t('横向') : L.t('纵向')}`), sect.cols && sect.cols.n > 1 ? row('列：', String(sect.cols.n)) : null].filter(Boolean));
      if (P.distinguish) box.appendChild(h('div', { class: 'tp-note', text: L.t('段落样式：{0}', styleDescription(d, p.pPr.style || 'Normal')) }));
    },
    refresh(b, reason) { if (reason === 'sel' || reason === 'doc') { L.clear(b); this.render(b); } },
  });
  /* ---------- Protect Document ---------- */
  TP.register({
    id: 'protect', title: '保护文档',
    render(b) {
      const d = doc();
      const pr = d.settings.protect || {};
      const fmt = ui.check('限制对选定的样式设置格式', !!pr.formatting, (v) => { pr.formatting = v; });
      const ed = ui.check('仅允许在文档中进行此类型的编辑：', pr.kind && pr.kind !== 'none', (v) => { kind.disabled = !v; });
      const kind = ui.select([['readOnly', '不允许任何更改（只读）'], ['tracked', '修订'], ['comments', '批注'], ['forms', '填写窗体']], pr.kind || 'readOnly', null, { id: 'pr-kind', style: 'width:100%' });
      kind.disabled = !(pr.kind && pr.kind !== 'none');
      const on = !!pr.on;
      const start = ui.button(on ? '停止保护' : '是，启动强制保护', async () => {
        if (on) {
          if (pr.hash) { const pw = await ui.prompt('密码：', '', '取消文档保护'); if (pw == null) return; if (await hash(pw) !== pr.hash) { ui.msg('密码不正确。', { icon: 'warn' }); return; } }
          d.settings.protect = Object.assign({}, pr, { on: false });
        } else {
          if (!ed.input.checked && !fmt.input.checked) { ui.msg('请先选择限制类型。', { icon: 'info' }); return; }
          const pw = await ui.prompt('输入新密码（可选）：', '', '启动强制保护');
          if (pw == null) return;
          d.settings.protect = { on: true, kind: ed.input.checked ? kind.value : 'none', formatting: fmt.input.checked, hash: pw ? await hash(pw) : null };
          if (d.settings.protect.kind === 'tracked') d.settings.track = true;
        }
        d.dirty = true;
        E.readOnly = false;
        TP.render();
        A().updateStatus();
        ui.refresh();
      }, { class: 'btn', style: 'width:100%' });
      b.append(sec(L.N('1. 格式设置限制'), fmt), sec(L.N('2. 编辑限制'), ed, kind), sec(L.N('3. 启动强制保护'), h('div', { class: 'tp-note', text: on ? '此文档已受保护，以防止意外编辑。' : '是否应用这些设置？（以后仍可关闭）' }), start));
    },
  });
  async function hash(s) { try { const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return Array.from(new Uint8Array(buf)).map((x) => x.toString(16).padStart(2, '0')).join(''); } catch (e) { return String(s.length); } }
  /* ---------- Mail Merge wizard ---------- */
  TP.register({ id: 'mailmerge', title: '邮件合并', render(b) { if (L.mailmerge) L.mailmerge.renderWizard(b); else b.appendChild(h('div', { class: 'tp-note', text: '邮件合并不可用。' })); }, refresh(b, reason) { if (reason === 'load') { L.clear(b); this.render(b); } } });

  /* ================= Document Map & thumbnails ================= */
  const DM = (P.docmap = { collapsed: new Set() });
  DM.mount = function (el) { DM.el = el; L.bus.on('doc-changed', L.debounce(() => { if (!el.hidden) (A().opts.thumbs ? P.thumbs.render() : DM.render()); }, 400)); L.bus.on('sel', L.debounce(() => { if (!el.hidden && !A().opts.thumbs) DM.mark(); }, 120)); L.bus.on('doc-loaded', () => { if (!el.hidden) DM.render(); }); };
  DM.render = function () {
    const el = DM.el;
    L.clear(el);
    const d = doc();
    D.reindex(d);
    const heads = [];
    for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) heads.push({ p, lv }); }
    if (!heads.length) { el.appendChild(h('div', { class: 'dm-item', style: 'color:#666', text: '（无标题）' })); return; }
    let hideBelow = 0;
    heads.forEach((x, i) => {
      if (hideBelow && x.lv > hideBelow) return;
      hideBelow = 0;
      const hasKids = heads[i + 1] && heads[i + 1].lv > x.lv;
      const coll = DM.collapsed.has(x.p.id);
      const tog = h('span', { class: 'dm-tog', text: hasKids ? (coll ? '+' : '−') : '' });
      /* the heading's own text, not translated */
      const it = h('div', { class: 'dm-item', 'data-pid': x.p.id, style: `padding-left:${4 + (x.lv - 1) * 12}px`, translate: 'no' }, tog, h('span', null, document.createTextNode(D.plainText(x.p).trim() || L.t('（空白）'))));
      it.title = D.plainText(x.p);
      tog.addEventListener('click', (e) => { e.stopPropagation(); if (coll) DM.collapsed.delete(x.p.id); else DM.collapsed.add(x.p.id); DM.render(); });
      it.addEventListener('click', () => { A().gotoPos(D.pos(x.p, 0)); const f = LY.fragFor(x.p, 0); if (f) f.scrollIntoView({ block: 'start' }); });
      el.appendChild(it);
      if (coll) hideBelow = x.lv;
    });
    DM.mark();
  };
  DM.mark = function () {
    if (!E.sel || !DM.el) return;
    const d = doc();
    let p = E.sel.f.p;
    if (D.storyOf(d, p) !== d.main) return;
    while (p && !D.headingLevel(d, p)) p = D.prevPara(d, p);
    L.$$('.dm-item', DM.el).forEach((x) => x.classList.toggle('on', !!p && +x.dataset.pid === p.id));
  };
  P.thumbs = {
    render() {
      const el = DM.el;
      L.clear(el);
      const box = h('div', { class: 'thumbs' });
      const w = Math.max(60, el.clientWidth - 40);
      LY.pages.forEach((pg, i) => {
        const k = w / pg.el.offsetWidth;
        const frame = h('div', { class: 'thumb', style: `width:${w}px;height:${pg.el.offsetHeight * k}px` });
        const clone = pg.el.cloneNode(true);
        clone.removeAttribute('data-pg');
        for (const e2 of clone.querySelectorAll('[contenteditable]')) e2.removeAttribute('contenteditable');
        clone.style.cssText += `;transform:scale(${k});transform-origin:0 0;margin:0;pointer-events:none;box-shadow:none`;
        frame.appendChild(clone);
        frame.addEventListener('click', () => { pg.el.scrollIntoView({ block: 'start' }); L.$$('.thumb', box).forEach((t, j) => t.classList.toggle('on', j === i)); });
        box.append(frame, h('div', { class: 'thumb-n', text: String(i + 1) }));
      });
      if (!LY.pages.length) box.appendChild(h('div', { class: 'tp-note', text: '缩略图在页面视图和阅读版式中显示页面。' }));
      el.appendChild(box);
    },
  };

  /* ================= Select Browse Object ================= */
  P.browseTo = function (kind, dir) {
    const d = doc();
    if (!E.sel) return;
    D.reindex(d);
    const targets = [];
    if (kind === 'section') { for (const s of D.sections(d)) { const b0 = d.main.blocks[s.from]; const p = b0 && (b0.t === 'p' ? b0 : D.firstPara([b0])); if (p) targets.push(D.pos(p, 0)); } }
    else if (kind === 'endnote' || kind === 'footnote' || kind === 'comment' || kind === 'field' || kind === 'graphic' || kind === 'edit') {
      for (const p of d.main.paras) {
        let o = 0;
        for (const it of p.runs) {
          const ok = (kind === 'footnote' && it.t === 'fn' && !it.self) || (kind === 'endnote' && it.t === 'en' && !it.self) || (kind === 'comment' && it.t === 'cs') || (kind === 'field' && it.t === 'fb') || (kind === 'graphic' && (it.t === 'img' || it.t === 'shape' || it.t === 'group' || it.t === 'chart')) || (kind === 'edit' && it.rPr && (it.rPr.ins || it.rPr.del));
          if (ok) targets.push(D.pos(p, o));
          o += D.ilen(it);
        }
      }
    } else if (kind === 'heading') { for (const p of d.main.paras) if (D.headingLevel(d, p)) targets.push(D.pos(p, 0)); }
    else if (kind === 'table') { D.walk(d.main, (b, c, i, st, parents) => { if (b.t === 'tbl' && !parents.length) { const p = D.firstPara([b]); if (p) targets.push(D.pos(p, 0)); } }); }
    if (!targets.length) { A().status('此文档中没有此类项目。'); return; }
    const cur = E.sel.f;
    const t = dir > 0 ? targets.find((x) => D.cmp(d, x, cur) > 0) : targets.slice().reverse().find((x) => D.cmp(d, x, cur) < 0);
    if (!t) { A().status(dir > 0 ? 'Quire 已到达文档结尾。' : 'Quire 已到达文档开头。'); return; }
    A().gotoPos(t);
  };

  /* ================= Print Preview ================= */
  const PV = (P.preview = { pages: 1 });
  PV.enter = function () {
    E.readOnly = true;
    LY.root.contentEditable = 'false';
    const bar = h('div', { class: 'toolbar preview-bar', id: 'tb-preview', role: 'toolbar', 'aria-label': 'Print Preview' });
    const btn = (icon, label, fn, tip, on) => { const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : '') + (on ? ' on' : ''), type: 'button', 'data-tip': tip || label, 'aria-label': tip || label, html: (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${L.esc(L.t(label))}</span>` : '') }); b.addEventListener('pointerdown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    const zoom = ui.combo({ id: 'pv-zoom', width: 74, tip: '缩放', options: () => ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', '页宽', '文字宽度', '整页', '双页'], value: () => Math.round(LY.zoom * 100) + '%', onChange: (v) => A().zoomTo(v) });
    const magBtn = btn('zoom', '', () => { PV.mag = !PV.mag; magBtn.classList.toggle('on', PV.mag); E.readOnly = PV.mag; LY.root.contentEditable = PV.mag ? 'false' : 'true'; }, L.N('放大镜'), true);
    PV.mag = true;
    bar.append(h('span', { class: 'tb-grip' }), btn('print', '', () => A().printPDF(), L.N('打印')), magBtn, btn('onePage', '', () => { LY.setMultiPage(0); LY.setZoom(LY.zoomFor('wholePage')); }, '单页'), btn('multiPage', '', (e) => PV.multiMenu(e.currentTarget), '多页'), zoom, btn('ruler', '', () => ui.exec('ruler'), '显示标尺'), btn('shrink', '', () => PV.shrink(), '缩小字体填充'), btn('fullScreen', '', () => ui.exec('fullScreen'), '全屏显示'), btn('', '关闭预览', () => A().setView(A().prevView || 'print'), '关闭预览'), btn('help', '', () => ui.exec('help'), L.N('帮助')));
    L.$('#toolbars').appendChild(bar);
    PV.bar = bar;
    PV.prevZoom = LY.zoom;
    LY.setZoom(LY.zoomFor('wholePage'));
    PV.click = (e) => {
      if (!PV.mag) return;
      const pg = e.target.closest('.pg');
      if (!pg) return;
      if (LY.zoom < 0.95) { LY.setMultiPage(0); LY.setZoom(1); pg.scrollIntoView({ block: 'center' }); L.$('#scroller').classList.add('zoomed'); }
      else { LY.setZoom(LY.zoomFor('wholePage')); pg.scrollIntoView({ block: 'start' }); L.$('#scroller').classList.remove('zoomed'); }
    };
    LY.scroller.addEventListener('click', PV.click);
    ui.refresh();
  };
  PV.exit = function () {
    if (PV.bar) PV.bar.remove();
    PV.bar = null;
    LY.scroller.removeEventListener('click', PV.click);
    L.$('#scroller').classList.remove('zoomed');
    E.readOnly = false;
    LY.root.contentEditable = 'true';
    LY.setMultiPage(0);
    if (PV.prevZoom) LY.setZoom(PV.prevZoom);
  };
  PV.multiMenu = function (anchor) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        const cells = h('div', { class: 'tgp-cells', style: 'grid-template-columns:repeat(3,24px)' });
        const lbl = h('div', { class: 'tgp-label', text: '取消' });
        for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) {
          const s = h('span', { style: 'height:30px;width:24px' });
          s.addEventListener('pointerenter', () => { L.$$('span', cells).forEach((x, k) => x.classList.toggle('on', Math.floor(k / 3) <= r && k % 3 <= c)); lbl.textContent = L.t('{0} × {1} 页', r + 1, c + 1); });
          s.addEventListener('click', () => { close(); LY.setMultiPage(c + 1); const n = c + 1; const pg = LY.pages[0]; const avail = LY.scroller.clientWidth - 40, availH = LY.scroller.clientHeight - 30; if (pg) LY.setZoom(Math.min((avail - 16 * n) / (pg.W * n), (availH - 14 * (r + 1)) / (pg.H * (r + 1)))); });
          cells.appendChild(s);
        }
        box.append(cells, lbl);
        return box;
      },
    }], anchor.getBoundingClientRect(), { cls: 'grid-menu' });
  };
  /** Shrink One Page: reduce font sizes until the last page disappears */
  PV.shrink = function () {
    const d = doc();
    const n0 = LY.pages.length;
    if (n0 < 2) { A().status('此文档已经只有一页。'); return; }
    let ok = false;
    E.readOnly = false;
    E.edit('减少一页', () => {
      D.touchAll();
      for (let step = 0; step < 6 && !ok; step++) {
        const scale = (rp) => { if (rp && rp.sz) rp.sz = Math.max(1, rp.sz - 0.5); };
        for (const k in d.styles) { const s = d.styles[k]; if (s.rPr && s.rPr.sz) scale(s.rPr); }
        d.defaults.rPr.sz = Math.max(1, (d.defaults.rPr.sz || 12) - 0.5);
        D.walk(d.main, (b) => { if (b.t === 'p') { for (const it of b.runs) scale(it.rPr); scale(b.rPr); } });
        D.stylesChanged();
        LY.render();
        ok = LY.pages.length < n0;
      }
      return E.sel;
    });
    E.readOnly = true;
    A().status(ok ? '文档已减少一页。' : 'Quire 无法将此文档减少一页。');
  };

  /* ================= Reading Layout toolbar ================= */
  const RL = (P.reading = {});
  RL.enter = function () {
    const bar = h('div', { class: 'readbar', role: 'toolbar', 'aria-label': '阅读版式' });
    const btn = (icon, label, fn, tip) => { const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : ''), type: 'button', 'data-tip': tip || label, 'aria-label': tip || label, html: (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${L.esc(L.t(label))}</span>` : '') }); b.addEventListener('pointerdown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    bar.append(btn('docMap', '文档结构图', () => A().toggleDocMap()), btn('thumbnails', L.N('缩略图'), () => A().toggleThumbs()), btn('find', L.N('查找'), () => ui.exec('find')), btn('research', L.N('信息检索'), () => ui.exec('research')), h('span', { class: 'tb-sep' }), btn('growFont', '', () => { LY.setZoom(LY.zoom * 1.1); }, '增大文字'), btn('shrinkFont', '', () => { LY.setZoom(LY.zoom / 1.1); }, '减小文字'), btn('onePage', '', () => { LY.setMultiPage(0); LY.setZoom(LY.zoomFor('wholePage')); }, '实际页面'), btn('multiPage', '', () => { LY.setMultiPage(2); LY.setZoom(LY.zoomFor('twoPages')); }, '允许多页'), h('span', { class: 'tb-sep' }), btn('', L.N('关闭'), () => A().setView(A().prevView && A().prevView !== 'reading' ? A().prevView : 'print'), L.N('关闭')));
    L.$('#toolbars').appendChild(bar);
    RL.bar = bar;
  };
  RL.exit = function () { if (RL.bar) RL.bar.remove(); RL.bar = null; };
  L.bus.on('view', (v) => { if (v === 'reading') { if (!RL.bar) RL.enter(); } else RL.exit(); });
})();
