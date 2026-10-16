/* Quire — built-in document templates (File ▸ New ▸ On my computer…). Placeholders are MACROBUTTON
 * fields: clicking one selects it so that typing replaces it, as in Word's templates.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const T = (L.templates = {});
  let fidN = 0;
  const fid = () => 'tf' + (++fidN) + '_' + D.nid();
  const P = (runs, pPr, rPr) => D.para(runs, pPr || {}, rPr || {});
  const t = (s, r) => D.text(s, r || {});
  /** [Click here and type …] placeholder */
  const ph = (label, r) => { const f = fid(); return [D.item('fb', { fid: f, instr: ' MACROBUTTON  DoFieldClick [' + label + '] ' }, r || {}), D.item('fs', { fid: f }, r || {}), D.text('[' + label + ']', r || {}), D.item('fe', { fid: f }, r || {})]; };
  const field = (instr, result, r) => { const f = fid(); return [D.item('fb', { fid: f, instr: ' ' + instr + ' ' }, r || {}), D.item('fs', { fid: f }, r || {}), D.text(result, r || {}), D.item('fe', { fid: f }, r || {})]; };
  const today = () => L.fields ? L.fields.formatDate(new Date(), 'MMMM d, yyyy') : new Date().toDateString();
  /** DATE field in the given format, showing today */
  const dateField = (fmt, r) => field(`DATE \\@ "${fmt}"`, L.fields ? L.fields.formatDate(new Date(), fmt) : new Date().toDateString(), r);
  /** a table with a shaded header row; '' cells are left empty to fill in */
  const table = (rows, widths, fill, pStyle) => {
    const tb = D.simpleTable(rows.length, widths.length, widths.reduce((a, b) => a + b, 0));
    rows.forEach((r, ri) => r.forEach((v, ci) => {
      const p = tb.rows[ri].cells[ci].blocks[0];
      p.runs = v ? [t(v, ri === 0 ? { b: true, color: 'FFFFFF' } : {})] : [];
      if (pStyle) p.pPr.style = pStyle;
      if (ri === 0) tb.rows[ri].cells[ci].tcPr.shd = { val: 'clear', fill };
    }));
    tb.rows[0].trPr.header = true;
    tb.grid = widths;
    L.tables && L.tables.fixWidths(tb);
    return tb;
  };
  const style = (d, id, name, o) => { d.styles[id] = Object.assign({ id, name, type: 'paragraph', basedOn: 'Normal', next: null, pPr: {}, rPr: {}, custom: true, q: true }, o); };
  const LIST = [
    { id: 'letter', name: '现代型信函', cat: '信函和传真' },
    { id: 'memo', name: '专业型备忘录', cat: '备忘录' },
    { id: 'fax', name: '典雅型传真', cat: '信函和传真' },
    { id: 'resume', name: '现代型简历', cat: '其他文档' },
    { id: 'report', name: '专业型报告', cat: '报告' },
    { id: 'agenda', name: '会议议程', cat: '其他文档' },
    { id: 'minutes', name: '会议记录', cat: '其他文档' },
    { id: 'todo', name: '待办清单', cat: '其他文档' },
    { id: 'flyer', name: '活动传单', cat: '其他文档' },
    { id: 'recipe', name: '食谱', cat: '其他文档' },
  ];
  T.list = () => LIST;
  T.categories = () => ['常规', '信函和传真', '备忘录', '报告', '其他文档'];
  const B = {
    letter(d) {
      d.sect.mt = 72; d.sect.ml = d.sect.mr = 90;
      style(d, 'CompanyName', '公司名称', { pPr: { sp: { a: 24 }, borders: { bottom: { val: 'single', sz: 0.75, color: '808080', space: 4 } } }, rPr: { font: 'Arial Black', sz: 16, spacing: -0.5 } });
      style(d, 'ReturnAddress', '寄信人地址', { pPr: { jc: 'right' }, rPr: { font: 'Arial', sz: 8 } });
      style(d, 'InsideAddress', '收信人地址', { rPr: { font: 'Arial', sz: 10 } });
      style(d, 'BodyTextLetter', '信函正文', { pPr: { sp: { a: 12, line: 1.15, rule: 'auto' }, jc: 'both' }, rPr: { font: 'Arial', sz: 10 } });
      d.main.blocks = [
        P(ph('键入寄信人地址'), { style: 'ReturnAddress' }),
        P(ph('在此键入公司名称', { font: 'Arial Black' }), { style: 'CompanyName' }),
        P(field('DATE \\@ "MMMM d, yyyy"', today()), { style: 'InsideAddress', sp: { a: 24 } }),
        P(ph('键入收信人地址'), { style: 'InsideAddress', sp: { a: 24 } }),
        P([t('尊敬的 '), ...ph('收信人'), t('：')], { style: 'BodyTextLetter' }),
        P([t('在这里写正文。若要一次更改所有段落的外观，请打开“格式 ▸ 样式和格式”，修改样式，而不是逐处手动设置格式。')], { style: 'BodyTextLetter' }),
        P([t('继续键入即可返回主视图。单击上面的占位符并键入内容，它就会消失。')], { style: 'BodyTextLetter' }),
        P([t('此致')], { style: 'BodyTextLetter', sp: { a: 36 } }),
        P(ph('键入您的姓名'), { style: 'InsideAddress' }),
        P(ph('键入职务'), { style: 'InsideAddress' }),
      ];
    },
    memo(d) {
      style(d, 'MemoTitle', '备忘录标题', { pPr: { sp: { a: 24 }, borders: { bottom: { val: 'single', sz: 3, color: '000000', space: 6 } } }, rPr: { font: 'Arial Black', sz: 36, spacing: -1.5 } });
      style(d, 'MessageHeader', '信头', { pPr: { sp: { a: 6 }, ind: { l: 1080, fl: -1080 }, tabs: [{ pos: 72, al: 'left' }] }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'MemoBody', '备忘录正文', { pPr: { sp: { a: 12 } }, rPr: { font: 'Arial', sz: 10 } });
      d.styles.MessageHeader.pPr.ind = { l: 72, fl: -72 };
      const label = (s) => t(s, { b: true, caps: true });
      d.main.blocks = [
        P(ph('在此键入公司名称'), { style: 'MemoBody', jc: 'right' }),
        P([t('备忘录')], { style: 'MemoTitle' }),
        P([label('收件人：'), D.item('tab'), ...ph('键入姓名')], { style: 'MessageHeader' }),
        P([label('发件人：'), D.item('tab'), ...ph('键入姓名')], { style: 'MessageHeader' }),
        P([label('抄送：'), D.item('tab'), ...ph('键入姓名')], { style: 'MessageHeader' }),
        P([label('日期：'), D.item('tab'), ...field('DATE \\@ "M/d/yyyy"', L.fields ? L.fields.formatDate(new Date(), 'M/d/yyyy') : '')], { style: 'MessageHeader' }),
        P([label('主题：'), D.item('tab'), ...ph('键入主题')], { style: 'MessageHeader', borders: { bottom: { val: 'single', sz: 0.75, color: '000000', space: 8 } }, sp: { a: 18 } }),
        P([t('如何使用此备忘录模板', { b: true })], { style: 'MemoBody' }),
        P([t('备忘录要开门见山，先说需要对方处理的事。标题保持简短；“格式”工具栏的“样式”框中有“标题 1–3”可用来组织结构。')], { style: 'MemoBody' }),
        P([t('装饰形状就是普通的图形对象：单击后按 Delete 可删除，也可拖动到其他位置。')], { style: 'MemoBody' }),
      ];
    },
    fax(d) {
      style(d, 'FaxTitle', '传真标题', { pPr: { jc: 'center', sp: { b: 24, a: 24 } }, rPr: { font: 'Garamond', sz: 36, caps: true, spacing: 6 } });
      style(d, 'FaxLabel', '传真标签', { rPr: { font: 'Garamond', sz: 10, caps: true, b: true } });
      d.main.blocks = [P(ph('公司名称'), { jc: 'center' }, { font: 'Garamond', sz: 14, smallCaps: true }), P([t('传真发送单')], { style: 'FaxTitle' })];
      const tb = D.simpleTable(4, 4, 432, { tblPr: { style: 'TableNormal' } });
      const cells = [['收件人：', '单击此处并键入姓名', '发件人：', '单击此处并键入姓名'], ['传真：', '单击此处并键入传真号码', '页数：', '单击此处并键入页数'], ['电话：', '单击此处并键入电话号码', '日期：', null], ['主题：', '单击此处并键入传真主题', '抄送：', '单击此处并键入姓名']];
      cells.forEach((r, ri) => r.forEach((v, ci) => { const p = tb.rows[ri].cells[ci].blocks[0]; if (ci % 2 === 0) { p.runs = [t(v)]; p.pPr.style = 'FaxLabel'; } else p.runs = v ? ph(v, { font: 'Garamond' }) : field('DATE \\@ "MMMM d, yyyy"', today(), { font: 'Garamond' }); }));
      tb.grid = [60, 156, 60, 156];
      L.tables && L.tables.fixWidths(tb);
      d.main.blocks.push(tb, P([t('☐ 紧急   ☐ 请审阅   ☐ 请提意见   ☐ 请答复   ☐ 请回收', { font: 'Garamond', sz: 10 })], { sp: { b: 18, a: 18 }, borders: { top: { val: 'single', sz: 0.75, color: '000000', space: 6 }, bottom: { val: 'single', sz: 0.75, color: '000000', space: 6 } } }),
        P([t('备注：', { b: true, caps: true, font: 'Garamond' }), ...ph('键入备注', { font: 'Garamond' })], { sp: { b: 12 } }));
    },
    resume(d) {
      d.sect.ml = d.sect.mr = 90;
      style(d, 'Name', '姓名', { pPr: { sp: { a: 6 }, borders: { bottom: { val: 'single', sz: 1.5, color: '000000', space: 6 } } }, rPr: { font: 'Arial', sz: 22, b: true, spacing: -1 } });
      style(d, 'Address1', '地址 1', { pPr: { jc: 'right' }, rPr: { font: 'Arial', sz: 8 } });
      style(d, 'SectionTitle', '节标题', { pPr: { sp: { b: 18, a: 6 }, keepNext: true }, rPr: { font: 'Arial', sz: 10, b: true, caps: true, spacing: 1 } });
      style(d, 'Objective', '求职意向', { pPr: { sp: { a: 12 } }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'CompanyNameRes', '公司名称（简历）', { pPr: { tabs: [{ pos: 160, al: 'left' }, { pos: 432, al: 'right' }], sp: { b: 6 } }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'JobTitle', '职务', { pPr: { sp: { a: 4 } }, rPr: { font: 'Arial Black', sz: 9 } });
      style(d, 'Achievement', '工作成果', { pPr: { ind: { l: 18, fl: -18 }, sp: { a: 3 } }, rPr: { font: 'Arial', sz: 10 } });
      const absId = D.addNum(d, D.makeBulletAbs('•', 'Symbol', ''));
      const ach = (s) => P([t(s)], { style: 'Achievement', num: { id: absId, lvl: 0 } });
      d.main.blocks = [
        P(ph('街道地址 • 城市、省/市、邮政编码 • 电话 • 电子邮件'), { style: 'Address1' }),
        P(ph('您的姓名'), { style: 'Name' }),
        P([t('求职意向')], { style: 'SectionTitle' }), P(ph('在此键入求职意向'), { style: 'Objective' }),
        P([t('工作经历')], { style: 'SectionTitle' }),
        P([...ph('开始 – 结束'), D.item('tab'), ...ph('公司名称', { b: true }), D.item('tab'), ...ph('城市、省/市')], { style: 'CompanyNameRes' }),
        P(ph('职务'), { style: 'JobTitle' }),
        ach('描述您取得的一项成果，尽量带上数字。'), ach('写出您负责的工作，以及因您而带来的变化。'), ach('每项成果一行，每行一句话。'),
        P([t('教育背景')], { style: 'SectionTitle' }),
        P([...ph('Start – End'), D.item('tab'), ...ph('学校名称', { b: true }), D.item('tab'), ...ph('城市、省/市')], { style: 'CompanyNameRes' }),
        P(ph('学位、专业和荣誉'), { style: 'Achievement' }),
        P([t('兴趣爱好')], { style: 'SectionTitle' }), P(ph('有助于申请的社团、志愿服务和兴趣爱好'), { style: 'Objective' }),
        P([t('提示')], { style: 'SectionTitle' }), P([t('单击任意方括号占位符即可直接键入替换。“样式”框中的“节标题”、“职务”和“工作成果”样式可让每条内容对齐。')], { style: 'Objective' }),
      ];
    },
    report(d) {
      d.styles.Title.rPr = { font: 'Arial Black', sz: 32, b: false, kern: 14 };
      d.styles.Title.pPr = { jc: 'left', sp: { b: 180, a: 12 }, outline: 0, borders: { bottom: { val: 'single', sz: 3, color: '000000', space: 6 } } };
      d.styles.Subtitle.rPr = { font: 'Arial', sz: 14, i: false, color: '595959' };
      d.styles.Subtitle.pPr = { jc: 'left', sp: { a: 120 } };
      d.styles.Heading1.rPr = { font: 'Arial Black', sz: 16, b: false };
      d.styles.Heading1.pPr = Object.assign({}, d.styles.Heading1.pPr, { pageBreakBefore: false, sp: { b: 24, a: 6 } });
      d.styles.Heading2.rPr = { font: 'Arial', sz: 13, b: true };
      d.styles.Normal.rPr = { font: 'Garamond', sz: 11 };
      d.styles.Normal.pPr = { sp: { a: 8, line: 1.15, rule: 'auto' } };
      d.main.blocks = [
        P(ph('Company Name'), { jc: 'left' }, { font: 'Arial', sz: 10, caps: true, b: true }),
        P(ph('在此键入报告标题'), { style: 'Title' }),
        P([t('专业报告')], { style: 'Subtitle' }),
        P([...field('DATE \\@ "MMMM d, yyyy"', today())], {}, {}),
        P([t('目录')], { style: 'TOCHeading', pageBreakBefore: true }),
        P([t('在下面的目录中按 F9 可更新。', { i: true, color: '7F7F7F' })]),
        P([D.item('fb', { fid: 'toc1', instr: ' TOC \\o "1-3" \\h \\z \\u ' }), D.item('fs', { fid: 'toc1' }), t('右键单击并选择“更新域”以生成目录。'), D.item('fe', { fid: 'toc1' })]),
        P([t('引言')], { style: 'Heading1', pageBreakBefore: true }),
        P([t('请将此文字替换为您自己的内容。用“标题 1”和“标题 2”设置的标题会进入目录，更新目录时会重新生成。')]),
        P([t('使用本报告')], { style: 'Heading2' }),
        P([t('用“文件 ▸ 页面设置”更改页边距和纸张大小。用“视图 ▸ 页眉和页脚”添加页眉或页脚；页码来自“插入 ▸ 页码”。')]),
        P([t('调研结果')], { style: 'Heading1' }),
        P([t('在此描述您的调研结果。用“插入 ▸ 图片 ▸ 图表”插入图表，用“插入 ▸ 引用 ▸ 题注”添加题注。')]),
        P([t('结论')], { style: 'Heading1' }),
        P([t('总结您的结论和建议。')]),
      ];
      d.hf.rf = { kind: 'ftr', id: 'rf', blocks: [D.para([D.item('tab'), D.item('tab'), D.text('第 '), ...field('PAGE', '1'), D.text(' 页')], { style: 'Footer' })] };
      d.sect.refs = { hdr: {}, ftr: { default: 'rf' } };
      d.styles.TOCHeading = d.styles.TOCHeading || D.builtinStyles().TOCHeading;
    },
    agenda(d) {
      style(d, 'AgendaTitle', '议程标题', { pPr: { jc: 'center', sp: { a: 6 } }, rPr: { font: 'Tahoma', sz: 26, b: true, color: '1F3864' } });
      style(d, 'AgendaItem', '议程项目', { pPr: { tabs: [{ pos: 360, al: 'left' }], sp: { a: 6 } }, rPr: { font: 'Tahoma', sz: 10 } });
      d.main.blocks = [P(ph('会议标题'), { style: 'AgendaTitle' }), P([...field('DATE \\@ "dddd, MMMM d, yyyy"', L.fields ? L.fields.formatDate(new Date(), 'dddd, MMMM d, yyyy') : ''), t('  •  '), ...ph('时间'), t('  •  '), ...ph('地点')], { jc: 'center', sp: { a: 18 } }, { font: 'Tahoma', sz: 10 })];
      const tb = D.simpleTable(6, 3, 432);
      const rows = [['时间', '议题', '负责人'], ['上午 9:00', '欢迎与介绍', '主持人'], ['上午 9:15', '回顾上次会议', '记录人'], ['上午 9:30', '项目进展', '团队负责人'], ['上午 10:15', '自由讨论', '全体'], ['上午 10:45', '行动事项与后续步骤', '主持人']];
      rows.forEach((r, ri) => r.forEach((v, ci) => { const p = tb.rows[ri].cells[ci].blocks[0]; p.runs = [t(v, ri === 0 ? { b: true, color: 'FFFFFF' } : {})]; p.pPr.style = 'AgendaItem'; if (ri === 0) tb.rows[ri].cells[ci].tcPr.shd = { val: 'clear', fill: '1F3864' }; }));
      tb.rows[0].trPr.header = true;
      tb.grid = [72, 252, 108];
      L.tables && L.tables.fixWidths(tb);
      d.main.blocks.push(tb, P([t('其他说明：', { b: true })], { style: 'AgendaItem', sp: { b: 18 } }), P(ph('键入需要的备注或准备事项'), { style: 'AgendaItem' }));
    },
    minutes(d) {
      style(d, 'MinutesTitle', '会议记录标题', { pPr: { sp: { a: 2 } }, rPr: { font: 'Tahoma', sz: 24, b: true, color: '1F3864' } });
      style(d, 'MinutesHeading', '会议记录小标题', { pPr: { sp: { b: 16, a: 6 }, keepNext: true, borders: { bottom: { val: 'single', sz: 0.75, color: '8EAADB', space: 2 } } }, rPr: { font: 'Tahoma', sz: 11, b: true, color: '1F3864' } });
      style(d, 'MinutesText', '会议记录正文', { pPr: { sp: { a: 4 }, tabs: [{ pos: 90, al: 'left' }] }, rPr: { font: 'Tahoma', sz: 10 } });
      const absId = D.addNum(d, D.makeBulletAbs('•', 'Symbol', ''));
      const head = (s) => P([t(s)], { style: 'MinutesHeading' });
      const line = (label, runs) => P([t(label, { b: true }), D.item('tab'), ...runs], { style: 'MinutesText' });
      const topic = (n) => P([t(n + '.  ', { b: true }), ...ph('议题', { b: true })], { style: 'MinutesText', sp: { b: 8, a: 2 } });
      const point = (s) => P(ph(s), { style: 'MinutesText', num: { id: absId, lvl: 0 } });
      d.main.blocks = [
        P(ph('会议名称'), { style: 'MinutesTitle' }),
        P([t('会议记录', { color: '595959' })], { style: 'MinutesText', sp: { a: 12 } }),
        line('日期：', dateField('dddd, MMMM d, yyyy')), line('时间：', ph('开始 – 结束')), line('地点：', ph('会议室或视频链接')),
        line('主持人：', ph('姓名')), line('记录人：', ph('姓名')),
        head('出席人员'), P(ph('所有出席人员的姓名'), { style: 'MinutesText' }), line('请假：', ph('姓名')),
        head('讨论'),
        topic(1), point('讨论了什么'), point('做出了什么决定'),
        topic(2), point('讨论了什么'), point('做出了什么决定'),
        head('行动事项'),
        table([['行动', '负责人', '截止日期'], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', '']], [276, 96, 72], '1F3864', 'MinutesText'),
        head('下次会议'), line('日期：', ph('日期和时间')), line('地点：', ph('会议室或视频链接')),
      ];
    },
    todo(d) {
      style(d, 'ListTitle', '清单标题', { pPr: { sp: { a: 0 } }, rPr: { font: 'Trebuchet MS', sz: 28, b: true, color: '2A64C8' } });
      style(d, 'ListHeading', '清单小标题', { pPr: { sp: { b: 18, a: 6 }, keepNext: true, borders: { bottom: { val: 'single', sz: 1.5, color: '2A64C8', space: 2 } } }, rPr: { font: 'Trebuchet MS', sz: 13, b: true, color: '2A64C8' } });
      style(d, 'Checklist', '清单项', { pPr: { ind: { l: 24, fl: -24 }, tabs: [{ pos: 24, al: 'left' }], sp: { a: 7 } }, rPr: { font: 'Trebuchet MS', sz: 11 } });
      const box = (label) => P([t('☐', { font: 'Segoe UI Symbol', sz: 13 }), D.item('tab'), ...ph(label)], { style: 'Checklist' });
      const head = (s) => P([t(s)], { style: 'ListHeading' });
      d.main.blocks = [
        P([t('待办事项清单')], { style: 'ListTitle' }),
        P(dateField('dddd, MMMM d, yyyy', { font: 'Trebuchet MS', sz: 11, color: '595959' }), { sp: { a: 6 } }),
        head('今天'), box('最重要的任务'), box('任务'), box('任务'), box('任务'),
        head('本周'), box('任务'), box('任务'), box('任务'),
        head('电话与杂事'), box('电话或杂事'), box('电话或杂事'),
        head('备注'), P(ph('需要记住的事', { font: 'Trebuchet MS', sz: 11 }), { sp: { a: 6 } }),
        P([t('提示：选中 ☐ 并键入 ☒（“插入 ▸ 符号”）即可打勾，或用“格式 ▸ 字体”加删除线。', { i: true, color: '7F7F7F', sz: 9, font: 'Trebuchet MS' })], { sp: { b: 18 } }),
      ];
    },
    flyer(d) {
      d.sect.mt = d.sect.mb = 54;
      style(d, 'FlyerKicker', '传单导语', { pPr: { jc: 'center', sp: { b: 36, a: 6 } }, rPr: { font: 'Trebuchet MS', sz: 16, caps: true, spacing: 3, color: 'E0682D', b: true } });
      style(d, 'FlyerTitle', '传单标题', { pPr: { jc: 'center', sp: { a: 12 } }, rPr: { font: 'Arial Black', sz: 48, color: '1F3864', spacing: -1 } });
      style(d, 'FlyerTagline', '传单标语', { pPr: { jc: 'center', sp: { a: 30 } }, rPr: { font: 'Georgia', sz: 18, i: true, color: '595959' } });
      style(d, 'FlyerWhen', '传单时间', { pPr: { jc: 'center', sp: { b: 6, a: 6 } }, rPr: { font: 'Trebuchet MS', sz: 22, b: true, color: '1F3864' } });
      style(d, 'FlyerText', '传单正文', { pPr: { jc: 'center', sp: { a: 10, line: 1.2, rule: 'auto' }, ind: { l: 36, r: 36 } }, rPr: { font: 'Georgia', sz: 13 } });
      const rule = { val: 'single', sz: 3, color: 'E0682D', space: 12 };
      d.main.blocks = [
        P(ph('诚邀您参加'), { style: 'FlyerKicker' }),
        P(ph('活动名称'), { style: 'FlyerTitle' }),
        P(ph('一句让人想来参加的话'), { style: 'FlyerTagline' }),
        P(ph('6 月 6 日 星期六'), { style: 'FlyerWhen', borders: { top: rule } }),
        P(ph('14:00 – 18:00'), { style: 'FlyerWhen' }),
        P(ph('地点名称、街道地址、城市', { sz: 16, b: false }), { style: 'FlyerWhen', borders: { bottom: rule }, sp: { a: 30 } }),
        P(ph('介绍活动：会发生什么、面向谁、需要带什么。两三句话即可。'), { style: 'FlyerText' }),
        P(ph('美食 • 音乐 • 儿童游戏', { b: true, color: 'E0682D' }), { style: 'FlyerText', sp: { a: 36 } }),
        P([t('请回复：', { b: true }), ...ph('姓名、电话或电子邮件，以及回复日期')], { style: 'FlyerText' }),
        P(ph('免费入场 · 欢迎所有人', { sz: 11, color: '7F7F7F' }), { style: 'FlyerText' }),
      ];
    },
    recipe(d) {
      style(d, 'RecipeTitle', '食谱标题', { pPr: { sp: { a: 4 } }, rPr: { font: 'Georgia', sz: 28, b: true, color: '7A3B10' } });
      style(d, 'RecipeHeading', '食谱小标题', { pPr: { sp: { b: 16, a: 6 }, keepNext: true }, rPr: { font: 'Georgia', sz: 14, b: true, color: '7A3B10' } });
      style(d, 'RecipeText', '食谱正文', { pPr: { sp: { a: 4, line: 1.15, rule: 'auto' } }, rPr: { font: 'Georgia', sz: 11 } });
      const bullets = D.addNum(d, D.makeBulletAbs('•', 'Symbol', ''));
      const steps = D.addNum(d, D.makeNumberAbs('decimal', '%1.'));
      const ing = (s) => P(ph(s), { style: 'RecipeText', num: { id: bullets, lvl: 0 } });
      const step = (s) => P(ph(s), { style: 'RecipeText', num: { id: steps, lvl: 0 }, sp: { a: 6 } });
      const head = (s) => P([t(s)], { style: 'RecipeHeading' });
      d.main.blocks = [
        P(ph('菜名'), { style: 'RecipeTitle' }),
        P(ph('一句话介绍这道菜：来历、什么时候做', { i: true, color: '595959' }), { style: 'RecipeText', sp: { a: 10 } }),
        table([['份数', '准备时间', '烹饪时间', '烤箱'], ['4', '15 分钟', '30 分钟', '180 °C / 350 °F']], [108, 108, 108, 108], '7A3B10', 'RecipeText'),
        head('材料'), ing('用量与食材'), ing('用量与食材'), ing('用量与食材'), ing('用量与食材'), ing('用量与食材'),
        head('做法'), step('第一步'), step('下一步'), step('下一步'), step('上桌'),
        head('备注'), P(ph('变化做法、搭配什么、能保存多久'), { style: 'RecipeText' }),
      ];
    },
  };
  T.build = function (id) {
    const d = D.newDoc();
    const prev = D.doc;
    D.doc = d;
    try { if (B[id]) B[id](d); } finally { D.doc = prev; }
    d.props.title = '';
    d.props.template = (LIST.find((x) => x.id === id) || {}).name || 'Normal';
    return d;
  };
  T.create = function (id) {
    const d = T.build(id);
    L.app.newDocument({ build: (nd) => Object.assign(nd, d) });
  };
  /* a click on a MACROBUTTON placeholder (or the caret moved into it) selects the whole field, so typing replaces it;
     after the mouse is let go: selecting it while the button is down makes the drag that follows move it */
  const selectPlaceholder = () => {
    const E = L.ed;
    if (!E.sel || !E.collapsed() || !L.fields || T.busy) return;
    const f = L.fields.at(E.sel.f);
    if (!f || !/^\s*MACROBUTTON\b/i.test(f.it.instr)) return;
    const g = L.O.findField(f.story, f.fid);
    if (!g.begin || !g.end) return;
    const c = E.sel.f;
    if ((c.p === g.begin.p && c.o <= g.begin.o) || (c.p === g.end.p && c.o >= g.end.o)) return;
    T.busy = true;
    E.setSel(D.pos(g.begin.p, g.begin.o), D.pos(g.end.p, g.end.o), { noScroll: true });
    T.busy = false;
  };
  L.bus.on('sel', () => { if (!L.ed.pressing()) selectPlaceholder(); });
  L.bus.on('sel-done', selectPlaceholder);
})();
