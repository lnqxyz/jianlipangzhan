#!/usr/bin/env python3
"""Univer 导出 docx 后的收尾处理（只改 word/document.xml，其余条目原样回写）。

处理五件事：

1. 页面改 A4 并统一页边距，使标题「居中」相对表格也居中。
   Univer 导出固定 pgSz 14400x16840（10in 宽）、页边距 750，而表格宽约 8300，
   标题按整页宽度居中会明显偏右。改成 A4 11906x16838 + 上下 1440 / 左右 1800，
   可用文本宽 = 11906 - 1800*2 = 8306。
2. 表格宽度与 tblGrid 等比缩放，撑满可用文本宽。
3. 按行规格重建单元格（gridSpan / vMerge）。
   Univer 的 mergeTo / mergeCells 只在逻辑层合并：导出时被吞掉的格子会以空
   <w:tc> 残留、跨列表格的 gridSpan 也不对，Word 里表现为多出来的小方块和
   贯穿竖线。这里按 SPECS 给出的每行 (span, vMerge) 规格重建，文本按原顺序
   回填到非 continue 的格子。
4. 连续多行并成一个格（「旁站过程记录」区域）：Univer 的表格单元格只能渲染
   第一段，所以源文档里每条记录各占一行；这里把这些行合并成「一行一格」，
   每行原文成为一个独立段落，格内逐行往下排，不再有行间横线。
5. 行高按参考图比例重设；段落行距写死单倍、段前段后 0，避免渲染器默认行距
   把整表顶到第二页。

用法：
    python3 postprocess.py <in.docx> <out.docx>
退出码：0 = 自检通过；1 = 自检不通过；2 = 用法/文件错误
"""
import re
import sys
import zipfile

PAGE_W = 11906
PAGE_H = 16838
MAR_TB = 1440
MAR_LR = 1800
CONTENT_W = PAGE_W - MAR_LR * 2  # 8306

SPACING = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
RUN_PR = '<w:rPr><w:sz w:val="18"/></w:rPr>'

TBL_RE = re.compile(r'<w:tbl>.*?</w:tbl>', re.S)
ROW_RE = re.compile(r'<w:tr[ >].*?</w:tr>', re.S)
TC_RE = re.compile(r'<w:tc>.*?</w:tc>', re.S)
GRIDCOL_RE = re.compile(r'<w:gridCol w:w="(\d+)"\s*/>')
PARA_RE = re.compile(r'<w:p>.*?</w:p>|<w:p/>', re.S)

# 每行的单元格规格：[(gridSpan, vMerge)]，vMerge 取 '' / 'restart' / 'continue'
R = 'restart'
C = 'continue'
N = ''

# 输出行规格：cells = 该行的格子；src = 吃掉源文档的几行；
# merge_paras = True 时，把吃掉的各行文本合成「一格多段」。
ROW_SPECS = [
    # 行高(twip) cells                                   src  merge_paras
    dict(h=460,  cells=[(1, N)] * 6,                      src=1),   # 0 日期/天气/环境温度
    dict(h=640,  cells=[(1, R), (2, R), (1, R), (2, R)],  src=1,   # 1 旁站项目名称/隧道名/旁站内容/整段内容（三行合并成一格）
         red_cells=(3,)),
    dict(h=640,  cells=[(1, C), (2, C), (1, C), (2, C)],  src=1),   # 2 纵向续格
    dict(h=640,  cells=[(1, C), (2, C), (1, C), (2, C)],  src=1),   # 3 纵向续格
    dict(h=480,  cells=[(1, N), (2, N), (1, N), (2, N)],  src=1),   # 4 旁站开始/结束时间
    dict(h=3400, cells=[(6, N)],                          src=9,    # 5 旁站过程记录：+ 记录 1-8（同一格内逐段；是否三条由 CHECKS 注入）
         merge_paras=True),
    dict(h=900,  cells=[(6, N)],                          src=1),   # 6 其他需要说明的事情
    dict(h=900,  cells=[(6, N)],                          src=1),   # 7 发现问题及处理意见
    dict(h=90,   cells=[(6, N)],                          src=1),   # 8 间隔空行
    dict(h=660,  cells=[(6, N)],                          src=1),   # 9 承包单位/项目监理机构
    dict(h=700,  cells=[(6, N)],                          src=1),   # 10 质检员/旁站监理人员
    dict(h=700,  cells=[(6, N)],                          src=1),   # 11 日期/日期
]
SRC_ROWS = sum(s['src'] for s in ROW_SPECS)

# 「旁站过程记录」格末尾的固定三条检查项（题面文字来自需求文档 2.1，一个字不动）。
# 勾选按台账依据判定：只有 CHECK_EVIDENCE 里列到的表/项才勾「是」并标红，
# 其余一律「是□　否□」全黑、留手填（0.2.9 起）。
CHECKS = [
    '施工单位质检员、技术员、试验员等有关人员是否到位。',
    '施工机械、压力表状态是否良好。',
    '原材料是否经过检验。',
]

# 依据表：表号 → 可勾「是」的检查项序号（1 人员 / 2 机械压力表 / 3 原材料）。
# 出处：注浆台账 A5「注浆设备」+ Q6「注浆压力(MPa)」→ 表 2/4/5 的机械；
# 配料单开盘单 B4 搅拌机编号 / L6 搅拌时间、开盘信息汇总 K/AC–AL 报告号
# → 表 8 的机械与原材料。其余 8 张表三项在台账里都没有列，一律不勾。
CHECK_EVIDENCE = {
    2: {2},
    4: {2},
    5: {2},
    8: {2, 3},
}

REC_NO_RE = re.compile(r'^\s*(\d+)、')


def next_check_no(texts):
    """检查项编号紧接已有过程记录，避免 1、2 之后跳 9、10、11。"""
    n = 0
    for t in texts:
        m = REC_NO_RE.match(t)
        if m:
            n = max(n, int(m.group(1)))
    return n + 1


def checks_from(start, table_no=0):
    """三条检查项：有台账依据的画「⟦是☑⟧」（红），没有的画「是□」全黑留手填。"""
    has = CHECK_EVIDENCE.get(table_no, set())
    out = []
    for i, t in enumerate(CHECKS, 1):
        marks = ('\u3000\u3000⟦是☑⟧\u3000否□' if i in has
                 else '\u3000\u3000是□\u3000否□')
        out.append(f'{start + i - 1}、{t}{marks}')
    return out


def span_of(tc: str) -> int:
    m = re.search(r'<w:gridSpan w:val="(\d+)"', tc)
    return int(m.group(1)) if m else 1


def text_of(tc: str) -> str:
    return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tc))


def escape(text: str) -> str:
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


RED_RUN_PR = '<w:rPr><w:color w:val="FF0000"/><w:sz w:val="18"/></w:rPr>'
MARK_RE = re.compile(r'\u27e6([^\u27e7]*)\u27e7')


def make_para(text: str, red_all: bool = False) -> str:
    """生成一段：单倍行距、9pt。

    文本中用 ⟦…⟧ 包住的片段以红色 run 输出（台账来源数据标红），
    其余片段保持黑色。输出中不再保留记号本身。

    red_all=True 时整段合成一个红 run（⟦⟧ 记号失效，直接去掉）：
    用于「旁站内容」「旁站过程记录」两格；但「/」仍留黑。
    """
    body = ''
    if red_all and text and text.strip() != '/':
        plain = MARK_RE.sub(lambda m: m.group(1), text)
        body = (f'<w:r>{RED_RUN_PR}<w:t xml:space="preserve">'
                f'{escape(plain)}</w:t></w:r>')
        return f'<w:p><w:pPr>{SPACING}{RUN_PR}</w:pPr>{body}</w:p>'
    if text:
        pos = 0
        for m in MARK_RE.finditer(text):
            if m.start() > pos:
                body += (f'<w:r>{RUN_PR}<w:t xml:space="preserve">'
                         f'{escape(text[pos:m.start()])}</w:t></w:r>')
            body += (f'<w:r>{RED_RUN_PR}<w:t xml:space="preserve">'
                     f'{escape(m.group(1))}</w:t></w:r>')
            pos = m.end()
        if pos < len(text):
            body += (f'<w:r>{RUN_PR}<w:t xml:space="preserve">'
                     f'{escape(text[pos:])}</w:t></w:r>')
    return f'<w:p><w:pPr>{SPACING}{RUN_PR}</w:pPr>{body}</w:p>'


def rewrite_tcpr(tc: str, span: int, vmerge: str) -> str:
    """保留边框与边距，只重写 gridSpan / vMerge（顺序符合 OOXML 规定）。"""
    m = re.search(r'<w:tcPr>(.*?)</w:tcPr>', tc, re.S)
    if not m:
        return tc
    inner = re.sub(r'<w:gridSpan w:val="\d+"/>', '', m.group(1))
    inner = re.sub(r'<w:vMerge(?: w:val="\w+")?/>', '', inner)
    prefix = ''
    if span > 1:
        prefix += f'<w:gridSpan w:val="{span}"/>'
    if vmerge == R:
        prefix += '<w:vMerge w:val="restart"/>'
    elif vmerge == C:
        prefix += '<w:vMerge/>'
    return tc[:m.start()] + f'<w:tcPr>{prefix}{inner}</w:tcPr>' + tc[m.end():]


def build_cell(tpl: str, span: int, vmerge: str, text: str,
               red_all: bool = False) -> str:
    """以 tpl 为外形模板，生成一格；text 写入其第一段。"""
    cell = rewrite_tcpr(tpl, span, vmerge)
    return re.sub(PARA_RE, lambda _m: make_para(text, red_all), cell, count=1)


def build_multi_para_cell(tpl: str, span: int, texts, red_range=None) -> str:
    """一格多段：把 tpl 的所有段落替换成 texts 依次生成的段落。

    red_range=(start, end)：该区间内的段落整段标红（过程记录条目），
    区间外保持「⟦⟧ 才红」或全黑——固定格头留黑，检查项的「是☑」靠 ⟦⟧ 标红。
    """
    cell = rewrite_tcpr(tpl, span, N)
    paras = PARA_RE.findall(cell)

    def _mk(i, t):
        red = red_range is not None and red_range[0] <= i < red_range[1]
        return make_para(t, red)

    body = ''.join(_mk(i, t) for i, t in enumerate(texts)) or make_para('')
    if not paras:
        return cell.replace('</w:tcPr>', '</w:tcPr>' + body, 1)
    first = cell.index(paras[0])
    last = first + sum(len(p) for p in paras)
    return cell[:first] + body + cell[last:]


def set_row_height(row: str, height: int) -> str:
    """把 <w:trPr> 里的 trHeight 改成指定 twip，atLeast。"""
    new = f'<w:trHeight w:val="{height}" w:hRule="atLeast"/>'
    if re.search(r'<w:trHeight[^/]*/>', row):
        return re.sub(r'<w:trHeight[^/]*/>', new, row, count=1)
    if '<w:trPr>' in row:
        return row.replace('<w:trPr>', '<w:trPr>' + new, 1)
    return row.replace('<w:tr>', '<w:tr><w:trPr>' + new + '</w:trPr>', 1)


def rebuild_table(table: str, stats, table_no: int = 0) -> str:
    grid = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', table, re.S)
    if not grid:
        return table
    ncols = len(GRIDCOL_RE.findall(grid.group(1)))

    matches = list(ROW_RE.finditer(table))
    rows = [m.group(0) for m in matches]

    if len(rows) == len(ROW_SPECS):
        stats['already'] += 1
        return table
    if len(rows) != SRC_ROWS:
        print(f'row count {len(rows)} != expected {SRC_ROWS}', file=sys.stderr)
        stats['bad'] += 1
        return table

    out, pos, idx = [], 0, 0
    for spec in ROW_SPECS:
        consumed = rows[idx:idx + spec['src']]
        spec_cells = spec['cells']
        if sum(s for s, _v in spec_cells) != ncols:
            stats['bad'] += 1
            idx += spec['src']
            continue
        texts = []
        for r in consumed:
            for c in TC_RE.findall(r):
                t = text_of(c)
                if t.strip():
                    texts.append(t)
        tpl = TC_RE.findall(consumed[0])[0]
        head = consumed[0][:consumed[0].index('<w:tc>')]
        last = consumed[-1]
        tail = last[last.rindex('</w:tc>') + len('</w:tc>'):]
        h = dyn_h5() if (spec is ROW_SPECS[5]) else spec['h']
        head = set_row_height(head, h)

        built = []
        if spec.get('merge_paras'):
            built.append(build_multi_para_cell(
                tpl, spec_cells[0][0],
                texts + checks_from(next_check_no(texts), table_no),
                red_range=(1, len(texts))))  # 0 是「旁站过程记录：」固定格头，留黑
        else:
            red_cells = spec.get('red_cells', ())
            for ci, (span, vm) in enumerate(spec_cells):
                text = '' if vm == C else (texts.pop(0) if texts else '')
                built.append(build_cell(tpl, span, vm, text,
                                        red_all=(ci in red_cells)))

        out.append(table[pos:matches[idx].start()])
        out.append(head + ''.join(built) + tail)
        pos = matches[idx + len(consumed) - 1].end()
        idx += len(consumed)
    out.append(table[pos:])
    table = ''.join(out)

    # tblGrid / tblW 等比缩放到 CONTENT_W
    m = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', table, re.S)
    widths = [int(x) for x in GRIDCOL_RE.findall(m.group(1))]
    old = sum(widths)
    if old:
        scaled = [round(w * CONTENT_W / old) for w in widths]
        scaled[scaled.index(max(scaled))] += CONTENT_W - sum(scaled)
        new_grid = '<w:tblGrid>' + ''.join(
            f'<w:gridCol w:w="{w}"/>' for w in scaled) + '</w:tblGrid>'
        table = table[:m.start()] + new_grid + table[m.end():]
        table = re.sub(r'<w:tblW w:w="\d+" w:type="dxa"/>',
                       f'<w:tblW w:w="{CONTENT_W}" w:type="dxa"/>', table, count=1)
    return table


# 签字区块边框矩阵（与参考图一致）：row6（其他）与 row7（发现问题框）底线保留；
# row8 间隔行、row9 承包单位、row10 质检员上下横线全部去掉；row11 日期行只留表底。
RECORDS_N = 12  # r5 实际条目数（--records 覆盖），动态行高用


def dyn_h5() -> int:
    return max(1200, min(3600, 360 * RECORDS_N + 400))


SIGN_BORDERS = {
    8: ('nil', 'nil'),
    9: ('nil', 'nil'),
    10: ('nil', 'nil'),
    11: ('nil', 'single'),
}


def strip_sign_borders(table: str) -> str:
    rows = list(ROW_RE.finditer(table))
    for idx, (top_v, bot_v) in SIGN_BORDERS.items():
        if idx >= len(rows):
            continue
        row = rows[idx].group(0)
        new_row = row
        for edge, val in (('top', top_v), ('bottom', bot_v)):
            pat = re.compile(r'<w:%s w:val="\w+"[^/]*/>' % edge)
            if val == 'nil':
                new_row = pat.sub(f'<w:{edge} w:val="nil"/>', new_row)
        table = table.replace(row, new_row, 1)
    return table


def set_spacing(xml: str) -> str:
    def repl(m):
        attrs = m.group(1) or ''
        ppr = m.group(2) or ''
        inner = ppr[7:-8] if ppr else ''
        inner = re.sub(r'<w:spacing[^>]*/>', '', inner)
        return f'<w:p{attrs}><w:pPr>{SPACING}{inner}</w:pPr>'

    return re.sub(r'<w:p( [^>]*)?>(<w:pPr>.*?</w:pPr>)?', repl, xml, flags=re.S)


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description='20 源行 docx → 12 行成品（A4/合并/标红/检查项）')
    ap.add_argument('src', help='输入 docx（20 源行）')
    ap.add_argument('dst', help='输出 docx')
    ap.add_argument('--records', type=int, default=0,
                    help='旁站过程记录实际条目数（含模块名行，0=按样板 12 条算）')
    ap.add_argument('--table-no', type=int, default=0,
                    help='表序号 1–12；决定三条检查项里哪几项有台账依据可勾「是」（0=一律不勾）')
    a = ap.parse_args()
    if a.records < 0:
        print('--records 不能为负', file=sys.stderr)
        return 2
    if a.table_no and not 1 <= a.table_no <= 12:
        print('--table-no 应在 1–12', file=sys.stderr)
        return 2
    src, dst = a.src, a.dst
    global RECORDS_N
    RECORDS_N = a.records or 12
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}

    xml = blobs['word/document.xml'].decode('utf-8')

    sect = ('<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425" w:num="1"/>'
            '</w:sectPr>') % (PAGE_W, PAGE_H, MAR_TB, MAR_LR, MAR_TB, MAR_LR)
    if not re.search(r'<w:sectPr.*?</w:sectPr>', xml, re.S):
        print('no sectPr found', file=sys.stderr)
        return 1
    xml = re.sub(r'<w:sectPr.*?</w:sectPr>', sect, xml, count=1, flags=re.S)

    stats = {'bad': 0, 'already': 0}
    tables = TBL_RE.findall(xml)
    if len(tables) != 1:
        print(f'expected 1 table, got {len(tables)}', file=sys.stderr)
        return 1
    xml = TBL_RE.sub(lambda m: rebuild_table(m.group(0), stats, a.table_no), xml, count=1)
    xml = strip_sign_borders(xml)
    xml = set_spacing(xml)

    # 自检：每行 gridSpan 之和 == 网格列数；tblW == 网格总宽；页面为 A4；
    # 纵向合并每个 continue 都必须能追到上方仍在延续的 restart 锚点
    ok = stats['bad'] == 0
    for t in TBL_RE.findall(xml):
        grid = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', t, re.S)
        ncols = len(GRIDCOL_RE.findall(grid.group(1)))
        active = set()  # 当前仍在延续的锚点列
        for r in ROW_RE.findall(t):
            cells = TC_RE.findall(r)
            if sum(span_of(c) for c in cells) != ncols:
                ok = False
            next_active = set()
            col = 0
            for c in cells:
                span = span_of(c)
                if 'vMerge w:val="restart"' in c:
                    next_active.add(col)
                elif 'vMerge' in c:
                    if col not in active:
                        ok = False  # continue 没有对应锚点
                    next_active.add(col)
                col += span
            active = next_active
        tw = re.search(r'<w:tblW w:w="(\d+)" w:type="dxa"/>', t)
        gw = sum(int(x) for x in GRIDCOL_RE.findall(grid.group(1)))
        if not tw or int(tw.group(1)) != gw:
            ok = False
    if not re.search(r'<w:pgSz w:w="%d"' % PAGE_W, xml):
        ok = False

    # 自检检查项：有依据的项必须是红 run 的「是☑」，其余必须是「是□」且不带红记号。
    # 注意 make_para 输出时已去掉 ⟦⟧ 记号，只能按「是☑ + RED_RUN_PR」判红。
    n_has = len(CHECK_EVIDENCE.get(a.table_no, set()))
    n_checked = xml.count('是☑')
    n_red = xml.count(RED_RUN_PR + '<w:t xml:space="preserve">是☑</w:t>')
    n_open = xml.count('是□')
    if n_checked != n_has or n_red != n_has or n_open != len(CHECKS) - n_has:
        ok = False
    if any(t not in xml for t in CHECKS):
        ok = False

    blobs['word/document.xml'] = xml.encode('utf-8')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, blobs[n])

    print(f'rebuilt_rows={len(ROW_SPECS)} bad_rows={stats["bad"]} '
          f'already={stats["already"]} page=A4 content_width={CONTENT_W} '
          f'checks={n_checked}/{len(CHECKS)} selfcheck={"PASS" if ok else "FAIL"}')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
