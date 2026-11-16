// v13：按用户 m00964 + m00985 重排
//   1) 旁站过程记录模块：模块名单独占一行；每一条记录各占一个整行（通栏），互不接排
//   2) 「是否」类三项（人员到位 / 机械压力表 / 原材料检验）台账无数据 → 整条删除
//   3) 旁站内容三台阶各占一行（左侧三列纵向合并）
// 约束：单元格内容必须单段，禁止 '\n'；TableRowHeightRule 直接传数值 1。

const oldTable = doc.getTables()[0];
if (oldTable) oldTable.deleteTable();
const leftovers = doc.getParagraphs();
for (let i = leftovers.length - 1; i >= 1; i -= 1) {
  leftovers[i].remove();
}

// ---------- 标题 + 表外信息行 ----------
const p0 = doc.getParagraphs()[0];
p0.setText('旁站记录表（隧道开挖）');
p0.setStyle({ horizontalAlign: api.Enum.HorizontalAlign.CENTER, textStyle: { bl: 1, fs: 15 } });
doc.appendParagraph('工程项目名称：新建深汕铁路工程　　施工合同段：SSSG-8 标　　编号：');

// ---------- 旁站过程记录 8 条（逐条独立成行）----------
const recs = [
  '1、工法：三台阶法；围岩等级：设计 上台阶Ⅲ级、中台阶左侧Ⅳ级、中台阶右侧Ⅳ级，实际 上台阶Ⅲ级、中台阶左侧Ⅳ级、中台阶右侧Ⅳ级；循环进尺：上台阶 3.6m、中台阶左侧 3.0m、中台阶右侧 3.0m。',
  '2、开挖时间：2026-11-05 03:20～18:20。',
  '3、格栅／型钢安装：2026-11-05 21:30～2026-11-06 00:20。',
  '4、钢筋网片安装：2026-11-05 21:30～2026-11-06 00:20。',
  '5、锁脚锚管安装：2026-11-06 00:20～02:20。',
  '6、注浆：2026-11-06 02:20～03:30；浆液材料及配合比 水泥∶水＝1∶1；锁脚注浆量 169.9L、中空锚杆注浆量 149.6L（上台阶）。',
  '7、报检：2026-11-06 03:30～03:50；监理报验 黄钰，施工报验 王梓鸣。',
  '8、喷射混凝土：2026-11-06 04:00～07:28；强度等级 C30；设计方量 上台阶 16.34m³、中台阶左侧 3.33m³、中台阶右侧 3.33m³；实际方量 29.0m³、8.0m³、8.0m³。',
];

// ---------- 行装配 ----------
const E = '';
const rows = [];
// r0 日期行
rows.push(['日　期', '2026年11月5日', '天　气', '/', '环境温度', '/']);
// r1..r3 旁站项目名称 / 旁站内容（三台阶各一行）
rows.push(['旁站项目名称', '淘金山隧道', '', '旁站内容', '上台阶：107028.2～107031.8', '']);
rows.push([E, E, E, E, '中台阶左侧：107016.6～107019.6', '']);
rows.push([E, E, E, E, '中台阶右侧：107019.6～107022.6', '']);
// r4 起止时间
rows.push(['旁站开始时间', '2026-11-05 03:20', '', '旁站结束时间', '2026-11-05 18:20', '']);
// r5 模块名单独一行
rows.push(['旁站过程记录：', E, E, E, E, E]);
// r6..r13 每条记录一行
for (const t of recs) rows.push([t, E, E, E, E, E]);
// 其他 / 发现
rows.push(['其他需要说明的事情：/', E, E, E, E, E]);
rows.push(['发现问题及处理意见：/', E, E, E, E, E]);
// 间隔空行
rows.push([E, E, E, E, E, E]);
// 签字三行
rows.push(['承包单位：上海工程局 SSSG-8 标项目经理部　　项目监理机构：深汕铁路 SSJL-8 标监理项目部', E, E, E, E, E]);
rows.push(['质 检 员：＿＿＿＿＿＿＿＿　　旁站监理人员：＿＿＿＿＿＿＿＿', E, E, E, E, E]);
rows.push(['日　　期：＿＿＿＿＿＿＿＿　　日　　期：＿＿＿＿＿＿＿＿', E, E, E, E, E]);

const NW = rows.length;
const colWidths = [95, 128, 82, 101, 97, 50];

const table = doc.insertTableFromData(rows, {
  width: 553,
  columnWidths: colWidths,
  headerRowCount: 0,
});
if (!table) throw new Error('table insert failed');

for (let c = 0; c < colWidths.length; c += 1) table.setColumnWidth(c, colWidths[c]);

// ---------- 合并 ----------
const mergeLog = [];
// 旁站内容三行：左侧 c0/c1/c3 纵向跨 3 行，右侧 c4 横向跨 c4-c5（每行各一处）
for (const c of [0, 3]) {
  const cell = table.getCell(1, c);
  mergeLog.push(['v r1c' + c, cell ? cell.mergeTo(3, 1) : 'no-cell']);
}
// 淘金山隧道：横向并 c1-c2、纵向跨 3 行
const nameCell = table.getCell(1, 1);
mergeLog.push(['v3x2 r1c1', nameCell ? nameCell.mergeTo(3, 2) : 'no-cell']);
for (let r = 1; r <= 3; r += 1) {
  const cell = table.getCell(r, 4);
  mergeLog.push(['h r' + r + 'c4', cell ? cell.mergeTo(1, 2) : 'no-cell']);
}
// 起止时间行两处横向合并
for (const c of [1, 4]) {
  const cell = table.getCell(4, c);
  mergeLog.push(['h r4c' + c, cell ? cell.mergeTo(1, 2) : 'no-cell']);
}
// 其余整行通栏
for (let r = 5; r < NW; r += 1) {
  const cell = table.getCell(r, 0);
  mergeLog.push(['h r' + r, cell ? cell.mergeTo(1, 6) : 'no-cell']);
}

// ---------- 合并后重写文本 ----------
table.setCellText(1, 0, '旁站项目名称');
table.setCellText(1, 1, '淘金山隧道');
table.setCellText(1, 3, '旁站内容');
table.setCellText(1, 4, '上台阶：107028.2～107031.8');
table.setCellText(2, 4, '中台阶左侧：107016.6～107019.6');
table.setCellText(3, 4, '中台阶右侧：107019.6～107022.6');
table.setCellText(4, 0, '旁站开始时间');
table.setCellText(4, 1, '2026-11-05 03:20');
table.setCellText(4, 3, '旁站结束时间');
table.setCellText(4, 4, '2026-11-05 18:20');
table.setCellText(5, 0, '旁站过程记录：');
for (let i = 0; i < recs.length; i += 1) table.setCellText(6 + i, 0, recs[i]);
table.setCellText(14, 0, '其他需要说明的事情：/');
table.setCellText(15, 0, '发现问题及处理意见：/');
table.setCellText(16, 0, '');
table.setCellText(17, 0, '承包单位：上海工程局 SSSG-8 标项目经理部　　项目监理机构：深汕铁路 SSJL-8 标监理项目部');
table.setCellText(18, 0, '质 检 员：＿＿＿＿＿＿＿＿　　旁站监理人员：＿＿＿＿＿＿＿＿');
table.setCellText(19, 0, '日　　期：＿＿＿＿＿＿＿＿　　日　　期：＿＿＿＿＿＿＿＿');

// ---------- 边框 / 行高 / 内边距 ----------
table.setTableBorder({ preset: api.Enum.DocsTableBorderPreset.All, color: '#000000', width: 1 });

const heights = [34, 26, 26, 26, 30, 24, 20, 20, 20, 20, 20, 20, 20, 20, 30, 30, 6, 24, 26, 26];
for (let r = 0; r < heights.length; r += 1) table.setRowHeight(r, heights[r], 1);
table.setDefaultCellMargin({ top: 2, bottom: 2, left: 2, right: 2 });

// ---------- 字号 ----------
const paras = doc.getParagraphs();
for (let i = 1; i < paras.length; i += 1) paras[i].setStyle({ textStyle: { fs: 9 } });

return {
  paragraphs: paras.length,
  rowCount: table.getRowCount(),
  colCount: table.getColumnCount(),
  recCount: recs.length,
  mergeFail: mergeLog.filter((m) => m[1] !== true),
  cell5: table.getCellText(5, 0),
  cell6: table.getCellText(6, 0).slice(0, 20),
  cell14: table.getCellText(14, 0),
};
