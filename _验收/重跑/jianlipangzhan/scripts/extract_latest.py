#!/usr/bin/env python3
"""从监理/注浆/配料单台账里找「有工序时间的最新日期」的数据行，组装 12 张旁站记录的内容。

支持一次传多份台账，按文件名关键词各自分派后合并成 1..12 号表：
  监理台账        → 1 隧道开挖 / 7 喷射混凝土 / 9 防排水安装 / 10 超前水平钻孔及加深炮孔
                    / 11 二衬预埋槽道 / 12 综合接地 / 3 超前小导管注浆 / 6 二衬带模注浆
  旁站注浆记录表  → 2 砂浆锚杆注浆 / 4 中空锚杆注浆 / 5 锁脚锚管注浆
  配料单及拌合物  → 8 混凝土浇筑（浇筑起止时间从监理台账对应浇筑时间列补）

取数口径（关键）：
  1. 每张表各自找「本模块有工序起止时间的最新一天」，不是全台账统一取一天。
  2. 旁站项目名称取该模块所在 sheet 的置顶标题（r0）。
  3. 旁站的开始/结束时间一律取该工序自己的操作起止时间；取不到就说明这一行没有旁站时间，
     换更早的日期，直到取到为止。
  4. 台账是唯一数据源，不补全、不推断；取不到的字段填 "/"。
  5. 台账原文一律包 ⟦…⟧ 记号，后处理转红色。

输出 JSON（stdout）：
{"ledger": "...", "type": "...", "ledger_date": "...", "tables": [{"no":1,...}, ...],
 "missing": [...], "warnings": [...]}

退出码：0=有数据；1=台账能读但所有目标表都无数据；2=用法/文件错误；3=认不出台账类型。
"""
import argparse
import datetime
import json
import re
import sys

xlrd = None  # 惰性导入：读 .xls 时在 open_book 里补上

TAIZHANG_KEYS = ('监理台账', '施工台账')
ZHUPU_KEYS = ('旁站注浆记录表', '注浆记录表台账', '旁站灌注记录表')
PEILIAO_KEYS = ('配料单及拌合物', '配料单及拌和物', '配料单台账', '配料单')

# 12 类表名（需求文档表 0 原文）
TITLES = {
    1: '隧道开挖', 2: '砂浆锚杆注浆', 3: '超前小导管注浆', 4: '中空锚杆注浆',
    5: '锁脚锚管注浆', 6: '二衬带模注浆', 7: '喷射混凝土', 8: '混凝土浇筑',
    9: '防排水安装', 10: '超前水平钻孔、加深炮孔施作', 11: '二衬预埋槽道', 12: '综合接地',
}


# ------------------------------------------------------------ 基础工具 --

def detect_type(name: str) -> str:
    for k in PEILIAO_KEYS:
        if k in name:
            return '配料单及拌合物'
    for k in ZHUPU_KEYS:
        if k in name:
            return '旁站注浆记录表'
    for k in TAIZHANG_KEYS:
        if k in name:
            return '监理台账'
    return ''


def open_book(path: str):
    global xlrd
    if path.lower().endswith('.xls'):
        import xlrd as _xlrd
        xlrd = _xlrd
        return xlrd.open_workbook(path), 'xlrd'
    import openpyxl
    return openpyxl.load_workbook(path, read_only=True, data_only=True), 'openpyxl'


def iter_rows(book, mode, sheet, max_cells=40):
    if mode == 'xlrd':
        for r in range(sheet.nrows):
            vals = []
            for c in range(min(sheet.ncols, max_cells)):
                v = sheet.cell(r, c)
                if v.ctype == 3:
                    try:
                        vals.append(xlrd.xldate_as_datetime(v.value, book.datemode))
                        continue
                    except Exception:
                        pass
                vals.append(v.value if v.ctype != 0 else '')
            yield r, vals
    else:
        for r, row in enumerate(sheet.iter_rows(max_col=max_cells, values_only=True)):
            yield r, list(row)


def sheet_rows(book, mode, name, max_cells=40):
    sheet = book.sheet_by_name(name) if mode == 'xlrd' else book[name]
    return list(iter_rows(book, mode, sheet, max_cells=max_cells))


def sheet_names(book, mode):
    return book.sheet_names() if mode == 'xlrd' else book.sheetnames


def sheet_title(rows):
    """取置顶框标题：r0 第一个非空单元格。"""
    for _r, vals in rows[:3]:
        for v in vals:
            s = _clean_txt(v, 80)
            if s:
                return s
    return ''


def parse_date(v):
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    m = re.match(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', str(v).strip())
    if m:
        try:
            return datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def fmt_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    return _clean_txt(v, 30) or '/'


def fmt_dt(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y-%m-%d %H:%M') if (v.hour or v.minute) else v.strftime('%Y-%m-%d')
    return _clean_txt(v, 30) or '/'


def _as_time(v):
    if isinstance(v, datetime.time):
        return v
    if isinstance(v, datetime.datetime):
        return v.time()
    m = re.match(r'(\d{1,2})[:：](\d{2})', str(v).strip())
    if m:
        return datetime.time(int(m.group(1)), int(m.group(2)))
    return None


FULLW = str.maketrans({'：': ':', '～': '~', '－': '-', '—': '-', '\n': ' ', '\u3000': ' '})
DATE_PAT = re.compile(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})')
TIME_PAT = re.compile(r'(\d{1,2}):(\d{2})')


def _norm(v):
    return str(v).translate(FULLW).strip() if v not in ('', None) else ''


def parse_span(v, base=None):
    """把工序时间串解析成 {start, end, d0, d1, t0, t1}。

    兼容：'2026.7.24-10:45至2026.7.24-19:56'、'2026.7.12 8:30-2026.7.14 7:40'、
    '2026.7.14 8:20-8:25'、'4:00~4:50'（无日期，用 base 补）。
    """
    s = _norm(v)
    if not s:
        return None
    dates = []
    for m in DATE_PAT.finditer(s):
        try:
            dates.append(datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass
    times = []
    for m in TIME_PAT.finditer(s):
        try:
            times.append(datetime.time(int(m.group(1)), int(m.group(2))))
        except ValueError:
            pass
    if not times:
        return None
    if not dates:
        if base is None:
            return None
        dates = [base]
    d0 = dates[0]
    d1 = dates[-1] if len(dates) > 1 else d0
    t0 = times[0]
    t1 = times[-1] if len(times) > 1 else None
    start = datetime.datetime.combine(d0.date(), t0)
    end = None
    if t1 is not None:
        end = datetime.datetime.combine(d1.date(), t1)
        if end < start:
            end = datetime.datetime.combine(d0.date() + datetime.timedelta(days=1), t1)
    return {'start': start, 'end': end, 'd0': d0, 'd1': d1, 't0': t0, 't1': t1}


def mk(v):
    return f'⟦{v}⟧'


def dk(m):
    """米数值 → DK 里程。106913.8 → DK106+913.8"""
    try:
        m = float(m)
        km = int(m) // 1000
        rem = m - km * 1000
        return f'DK{km}+{rem:g}'
    except Exception:
        return _clean_txt(m, 20)


def num_s(v):
    """数值显示：47.0 → 47；0.36 保留。"""
    if isinstance(v, float) and v == int(v) and abs(v) < 1e9:
        return str(int(v))
    if isinstance(v, (int, float)):
        return f'{v:g}'
    return _clean_txt(v, 20)


def _clean_txt(v, limit=60):
    s = str(v).translate(FULLW).strip() if v not in ('', None) else ''
    s = re.sub(r'\s+', ' ', s)
    return s[:limit]


def _g(vals, i):
    return vals[i] if i < len(vals) and vals[i] not in ('', None) else ''


def _num(vals, i):
    v = _g(vals, i)
    return v if isinstance(v, (int, float)) or re.match(r'^-?\d+(\.\d+)?$', str(v).strip()) else ''


def _mile(vals):
    """从一行的前几列里抠出两个里程数，返回 'DK…～DK…' 或 ''。"""
    nums = [float(v) for v in vals[:5]
            if isinstance(v, (int, float)) and 100000 < float(v) < 120000]
    if len(nums) >= 2:
        return f'{dk(min(nums[:2]))}～{dk(max(nums[:2]))}'
    if len(nums) == 1:
        return dk(nums[0])
    return ''


def tf(dt, base_date=None):
    """旁站时间显示：与记录日期同天写 HH:MM，跨天写 MM-DD HH:MM。"""
    if not isinstance(dt, datetime.datetime):
        return '/'
    if base_date and dt.date() == base_date.date():
        return mk(dt.strftime('%H:%M'))
    return mk(dt.strftime('%m-%d %H:%M'))


def make_table(no, project, date, content, start, end, records, note='', source=None):
    recs = [f'{i}、{r}' for i, r in enumerate([x for x in records if x], 1)]
    fields = {'date': fmt_date(date) if date else '/',
              'weather': '/', 'temperature': '/',
              'project': project or '/',
              'content': content or '/',
              'start': start or '/', 'end': end or '/'}
    tb = dict(no=no, title=TITLES[no], fields=fields, records=recs or ['1、/'])
    if note:
        tb['missing_reason'] = note
    if source:
        # 溯源：这张表的每个值取自哪个台账、哪个 sheet、Excel 第几行、哪几列
        tb['source'] = source
    return tb


def src(sheet, row0, cells='', row_end0=None):
    """构造一处溯源记录。sheet=工作表名；row0/row_end0=iter_rows 给的 0 基行号，
    写进 JSON 时转成 Excel 行号；给 row_end0 表示这条内容取自一段连续行。"""
    if sheet is None or row0 is None:
        return None
    try:
        row = int(row0) + 1          # 0 基 → Excel 1 基
    except (TypeError, ValueError):
        return None
    out = {'sheet': sheet, 'row': row, 'cells': cells}
    if row_end0 is not None:
        try:
            out['row_end'] = int(row_end0) + 1
        except (TypeError, ValueError):
            pass
    return out


# ---------------------------------------------------------------- 监理台账 --

STEP_SHEETS = ('上台阶开挖', '中台阶左侧', '中台阶右侧', '下台阶左侧', '下台阶右侧')
POUR_SHEETS = ('暗洞二衬', '仰拱', '仰拱填充', '通信洞室')


def _step_date(vals):
    ds = [vals[i] for i in range(9, min(12, len(vals)))
          if isinstance(vals[i], datetime.datetime)]
    return max(ds) if ds else None


def _span_row(rows, col, base_idx=None):
    """从一行里按列号取工序时间。返回 (span, base_date)。"""
    return None


def _collect_pours(books):
    """从监理台账扫出所有「混凝土浇筑时间」→ [{date,start,end,label,grade,design,actual,sheet}]。"""
    out = []
    if '监理台账' not in books:
        return out
    book, mode, _p = books['监理台账']
    for name in sheet_names(book, mode):
        if not any(k in name for k in POUR_SHEETS):
            continue
        rows = sheet_rows(book, mode, name, max_cells=32)
        hdr = None          # (行号, 浇筑时间列, 里程列起点)
        section = ''
        for r, vals in rows:
            texts = [_norm(v) for v in vals]
            # 段标题（通信洞室：初支/边墙及底板/拱墙二衬）
            joined = ' '.join(t for t in texts if t)
            if joined and not any(t == '序号' for t in texts) and len([t for t in texts if t]) <= 2:
                if any(k in joined for k in ('初支', '边墙', '拱墙', '仰拱', '底板')):
                    section = joined
            hit = next((i for i, t in enumerate(texts) if '混凝土浇筑时间' in t), None)
            if hit is not None:
                hdr = hit
                continue
            if hdr is None:
                continue
            sp = parse_span(_g(vals, hdr), base=None)
            if not sp:
                continue
            grade = ''
            design = actual = ''
            for i, t in enumerate(texts):
                if re.fullmatch(r'C\d+', t) and not grade:
                    grade = t
            # 方量：浇筑时间列之后两个数字
            tail = [_num(vals, hdr + k) for k in (1, 2)]
            nums = [v for v in tail if v != '']
            if len(nums) >= 2:
                design, actual = nums[0], nums[1]
            out.append({'date': sp['start'], 'start': sp['start'], 'end': sp['end'],
                        'label': (section + ' ' if section else '') + _mile(vals),
                        'mile': _mile(vals), 'grade': grade,
                        'design': design, 'actual': actual, 'sheet': name, 'r': r,
                        'col': hdr})
    return out


def scan_jianli(book, mode):
    """监理台账：12 类里属于它的 6 张（1/7/9/10/11/12）+ 3 前置小导管 + 6 带模注浆。"""
    names = sheet_names(book, mode)
    tables = {}
    warnings = []

    # ---- 逐 sheet 读一遍，留下原始行
    data = {}
    for name in names:
        data[name] = sheet_rows(book, mode, name, max_cells=40)

    def title_of(name):
        return sheet_title(data.get(name, []))

    def proj_of(name):
        t = title_of(name) or '淘金山隧道进口'
        m = re.match(r'(.+?)进口', t)
        return (m.group(1) if m else t) or '淘金山隧道'

    # ============ 1 隧道开挖 ============
    step_rows = []      # (date, sheet, r, vals)
    for name in names:
        if not ('台阶开挖' in name or any(name.startswith(s) for s in STEP_SHEETS)):
            continue
        if '开挖' not in name:
            continue
        for r, vals in data[name]:
            d = _step_date(vals)
            if d is None:
                continue
            sp = parse_span(_g(vals, 12), base=d)
            if not sp:
                continue     # 没有开挖时间＝这一行不能当旁站依据，跳过
            step_rows.append((sp['start'].date(), name, r, vals, sp))
    if step_rows:
        latest = max(x[0] for x in step_rows)
        hits = sorted([x for x in step_rows if x[0] == latest], key=lambda x: (x[2], x[1]))
        content, recs = [], []
        s0 = min(x[4]['start'] for x in hits)
        e0 = max((x[4]['end'] or x[4]['start']) for x in hits)
        for d, name, r, vals, sp in hits:
            label = name.replace('开挖', '')
            seg = _mile(vals)
            if seg:
                content.append(f'{mk(label)}：{mk(seg)}')
            parts = []
            gong = _clean_txt(_g(vals, 1), 20)
            rock = _clean_txt(_g(vals, 8) or _g(vals, 7), 10)
            reach = _num(vals, 4)
            if gong:
                parts.append(f'工法 {mk(gong)}')
            if rock:
                parts.append(f'围岩 {mk(rock)}')
            if reach != '':
                parts.append(f'循环进尺 {mk(num_s(reach))} m')
            total = _num(vals, 5)
            if total != '':
                parts.append(f'累计进尺 {mk(num_s(total))} m')
            parts.append(f'开挖时间 {mk(_norm(_g(vals, 12)))}')
            for nm, c in (('格栅/型钢安装', 13), ('钢筋网片', 14), ('锁脚锚杆安装', 15),
                          ('注浆', 16), ('报检', 23), ('喷锚', 25)):
                txt = _norm(_g(vals, c))
                if txt and txt != '/':
                    parts.append(f'{nm} {mk(txt)}')
            for nm, c in (('锁脚钢管', 29), ('中空锚杆', 30), ('砂浆锚杆', 31)):
                txt = _clean_txt(_g(vals, c), 10)
                if txt and txt != '/':
                    parts.append(f'{nm} {mk(txt)} 根')
            qty = _clean_txt(_g(vals, 32), 10)
            if qty and qty != '/':
                segq = _clean_txt(_g(vals, 33), 20)
                parts.append(f'监控量测点 {mk(qty)} 个' + (f'（{mk(segq)}）' if segq and segq != '/' else ''))
            bm = _clean_txt(_g(vals, 36), 30)
            if bm and bm != '/':
                parts.append(f'备注 {mk(bm)}')
            sup = _clean_txt(_g(vals, 24), 12)
            if sup:
                parts.append(f'监理报验 {mk(sup)}')
            recs.append(f'{mk(label)}（{mk(seg) if seg else "/"}）：' + '；'.join(parts) + '。')
        tables[1] = make_table(1, proj_of(hits[0][1]), s0,
                               '；'.join(content), tf(s0, s0), tf(e0, s0), recs,
                               source=[src(f'监理台账/{name}', r,
                                           'c3/c4 起止里程、c5 循环进尺、c6 累计进尺、c8/c9 围岩、'
                                           'c13 开挖时间、c14~c17 格栅/网片/锁脚/注浆时间、'
                                           'c24 报检时间、c25 监理报验、c26 喷锚时间、c27 砼强度')
                                       for d, name, r, vals, sp in hits])
    else:
        tables[1] = make_table(1, '', None, '', '', '', [], '台账无带开挖时间的台阶记录')

    # ============ 7 喷射混凝土 ============
    spray_rows = []
    for name in names:
        rows = data[name]
        if any(name.startswith(s) or '台阶开挖' in name for s in STEP_SHEETS) and '开挖' in name:
            c_list, c_grade, c_des, c_act, c_sup, c_date = 25, 26, 27, 28, 24, 9
        elif '仰拱初支' in name:
            c_list, c_grade, c_des, c_act, c_sup, c_date = 13, 14, 15, 16, 12, 9
        elif '通信洞室' in name:
            continue    # 通信洞室初支的喷锚在分段表里，另处理
        else:
            continue
        for r, vals in rows:
            sp = parse_span(_g(vals, c_list), base=_step_date(vals) or _span_date_fallback(vals))
            if not sp:
                continue
            spray_rows.append((sp['start'].date(), name, r, vals, sp))
    # 通信洞室 初支 段（r2 表头，c13 喷锚时间）
    if '通信洞室' in data:
        hdr = None
        for r, vals in data['通信洞室']:
            texts = [_norm(v) for v in vals]
            if any('喷锚时间' in t for t in texts):
                hdr = next(i for i, t in enumerate(texts) if '喷锚时间' in t)
                continue
            if hdr is None:
                continue
            sp = parse_span(_g(vals, hdr))
            if sp:
                spray_rows.append((sp['start'].date(), '通信洞室初支', r, vals, sp))
    if spray_rows:
        latest = max(x[0] for x in spray_rows)
        hits = sorted([x for x in spray_rows if x[0] == latest], key=lambda x: (x[2], x[1]))
        content, recs = [], []
        s0 = min(x[4]['start'] for x in hits)
        e0 = max((x[4]['end'] or x[4]['start']) for x in hits)
        for d, name, r, vals, sp in hits:
            label = name.replace('开挖', '') if '开挖' in name else name
            if name == '通信洞室初支':
                grade = _clean_txt(_g(vals, 14), 12)
                des, act = _num(vals, 15), _num(vals, 16)
                seg = _mile(vals)
            elif '仰拱初支' in name:
                grade = _clean_txt(_g(vals, 14), 12)
                des, act = _num(vals, 15), _num(vals, 16)
                seg = _mile(vals)
            else:
                grade = _clean_txt(_g(vals, 26), 12)
                des, act = _num(vals, 27), _num(vals, 28)
                seg = _mile(vals)
            content.append(f'{label} {mk(seg) if seg else "/"}（{mk(grade) if grade else "/"}）')
            parts = []
            if seg:
                parts.append(f'部位 {mk(seg)}')
            if grade:
                parts.append(f'强度 {mk(grade)}')
            if des != '' and act != '':
                parts.append(f'设计方量 {mk(des)} m³、实际 {mk(act)} m³')
            parts.append(f'喷锚时间 {mk(_norm(_g(vals, 13 if name == "通信洞室初支" else (13 if "仰拱初支" in name else 25))))}')
            recs.append(f'{label}：' + '；'.join(parts) + '。')
        tables[7] = make_table(7, proj_of(hits[0][1]) if hits[0][1] in data else '', s0,
                               '；'.join(content), tf(s0, s0), tf(e0, s0), recs,
                               source=[src(f'监理台账/{n}', r,
                                           ('c14 喷锚时间、c15 砼强度、c16 砼设计方量、c17 砼实际方量'
                                            if n == '仰拱初支' else
                                            '第 3 行「初支」段表头：c14 喷锚时间、c15 砼强度、'
                                            'c16 设计方量、c17 实际方量'
                                            if n.startswith('通信洞室') else
                                            'c26 喷锚时间、c27 砼强度、c28 砼设计方量、c29 砼实际方量'))
                                       for d, n, r, vals, sp in hits])
    else:
        tables[7] = make_table(7, '', None, '', '', '', [], '台账无带喷锚时间的记录')

    # ============ 9 防排水安装 ============
    wp_rows = []
    for name in ('暗洞二衬', '仰拱'):
        if name not in data:
            continue
        cols = {'暗洞二衬': (12, 9), '仰拱': (9, 10)}[name]     # (土工布/防水板, 盲管/防排水板)
        for r, vals in data[name]:
            if r < 3:
                continue
            sp1 = parse_span(_g(vals, cols[0]))
            sp2 = parse_span(_g(vals, cols[1]))
            if not (sp1 or sp2):
                continue
            # 主工序＝土工布/防水板铺设（一次连续作业），盲管安装是前一道工序；
            # 旁站起止只取主工序自己的时间，避免把 7-12 的盲管和 7-20 的土工布拉成一次旁站。
            main = sp1 or sp2
            wp_rows.append({'date': main['start'].date(), 'start': main['start'],
                            'end': main['end'] or main['start'],
                            'sheet': name, 'r': r, 'vals': vals,
                            'cloth': _norm(_g(vals, cols[0])), 'pipe': _norm(_g(vals, cols[1])),
                            'check': _norm(_g(vals, 13 if name == '暗洞二衬' else 11)),
                            'sup': _clean_txt(_g(vals, 14 if name == '暗洞二衬' else 12), 12)})
    if wp_rows:
        latest = max(x['date'] for x in wp_rows)
        hits = [x for x in wp_rows if x['date'] == latest]
        content, recs = [], []
        s0 = min(x['start'] for x in hits)
        e0 = max(x['end'] for x in hits)
        for x in hits:
            seg = _mile(x['vals'])
            content.append(f'{x["sheet"]} {mk(seg)}' if seg else x['sheet'])
            parts = []
            if x['cloth']:
                parts.append(f'土工布、防水板铺设 {mk(x["cloth"])}')
            if x['pipe']:
                parts.append(f'盲管安装/防排水板 {mk(x["pipe"])}')
            if x['check']:
                parts.append(f'报检 {mk(x["check"])}')
            if x['sup']:
                parts.append(f'监理报验 {mk(x["sup"])}')
            recs.append(f'{x["sheet"]}（{mk(seg) if seg else "/"}）：' + '；'.join(parts) + '。')
        sh = hits[0]['sheet']
        tables[9] = make_table(9, proj_of(sh) if sh in data else '', s0,
                               '；'.join(content), tf(s0, s0), tf(e0, s0), recs,
                               source=[src(f'监理台账/{x["sheet"]}', x['r'],
                                           'c10 盲管安装/防排水板、c13 土工布/防水板铺设、'
                                           'c14 报检时间、c15 监理报验'
                                           if x['sheet'] == '暗洞二衬' else
                                           'c10 土工布/防水板铺设、c11 盲管安装/防排水板、'
                                           'c12 报检时间、c13 监理报验')
                                       for x in hits])
    else:
        tables[9] = make_table(9, '', None, '', '', '', [], '台账无防排水工序时间记录')

    # ============ 10 超前水平钻 + 加深炮孔 ============
    recs10, c10, d10s, e10s, d10 = [], [], [], [], None
    drill_best = None
    if '超前水平钻' in data:
        for r, vals in data['超前水平钻']:
            if r < 2:
                continue
            sp = parse_span(_g(vals, 5), base=parse_date(_g(vals, 4)))
            if not sp:
                continue
            if drill_best is None or sp['start'] >= drill_best['start']:
                drill_best = {'start': sp['start'], 'end': sp['end'], 'vals': vals, 'r': r}
    blast_best = None
    if '加深炮孔' in data:
        groups = {}
        cur = None      # 当前组：组首行给出日期/时间/位置，后续孔号行只有孔号+孔深
        for r, vals in data['加深炮孔']:
            if r < 2:
                continue
            d = parse_date(_g(vals, 2))
            txt = _norm(_g(vals, 3))
            if d and txt:
                sp = parse_span(txt, base=d)
                if sp:
                    cur = {'start': sp['start'], 'end': sp['end'], 'head': vals,
                           'r': r, 'r_last': r, 'holes': [(vals, _num(vals, 5))]}
                    groups.setdefault(sp['start'].date(), []).append(cur)
                continue
            if cur is not None and _num(vals, 4) != '':
                cur['holes'].append((vals, _num(vals, 5)))
                cur['r_last'] = r
        if groups:
            k = max(groups)
            g = groups[k]
            blast_best = {'start': min(x['start'] for x in g),
                          'end': max((x['end'] or x['start']) for x in g),
                          'groups': g}
    cands = []
    if drill_best:
        cands.append(('drill', drill_best['start'], drill_best))
    if blast_best:
        cands.append(('blast', blast_best['start'], blast_best))
    if cands:
        d10 = max(x[1] for x in cands)
        picks = [x for x in cands if x[1].date() == d10.date()]
        for kind, _dt, obj in picks:
            if kind == 'drill':
                v = obj['vals']
                recs10.append('超前水平钻：里程 {}；检测长度 {} 米；施做时间 {}；'
                              '下次施做里程 {}；搭接 {} 米。'.format(
                                  mk(f'{dk(_g(v, 1))}～{dk(_g(v, 2))}'),
                                  mk(_g(v, 3)),
                                  mk(_norm(_g(v, 5))),
                                  mk(dk(_g(v, 6))) if _g(v, 6) not in ('', None) else '/',
                                  mk(_g(v, 7)) if _g(v, 7) not in ('', None) else '/'))
                c10.append(f'超前水平钻 {mk(f"{dk(_g(v, 1))}～{dk(_g(v, 2))}")}')
                d10s.append(obj['start'])
                e10s.append(obj['end'] or obj['start'])
            else:
                gs = obj['groups']
                v = gs[0]['head']
                holes = [h for x in gs for h in x['holes']]
                depths = [h[1] for h in holes if h[1] != '']
                recs10.append('加深炮孔：施做 {}；掌子面里程 {}；当日 {} 孔；'
                              '孔深 {} 米；终孔位置 {}；外插角 {}°。'.format(
                                  mk(obj['start'].strftime('%Y-%m-%d')),
                                  mk(dk(_g(v, 1))) if _g(v, 1) not in ('', None) else '/',
                                  mk(len(holes)),
                                  mk('、'.join(str(d) for d in depths[:3])) + (f' 等{len(depths)}孔' if len(depths) > 3 else '') if depths else '/',
                                  mk(dk(_g(v, 6))) if _g(v, 6) not in ('', None) else '/',
                                  mk(num_s(_g(v, 7))) if _g(v, 7) not in ('', None) else '/'))
                _seg10 = f'{dk(_g(v, 1))}～{dk(_g(v, 6))}' if _g(v, 1) not in ('', None) and _g(v, 6) not in ('', None) else ''
                c10.append(f'加深炮孔 {mk(_seg10)}（{mk(len(holes))} 孔）' if _seg10 else f'加深炮孔 {mk(len(holes))} 孔')
                d10s.append(obj['start'])
                e10s.append(obj['end'] or obj['start'])
        _src10 = []
        for kind, _dt, obj in picks:
            if kind == 'drill':
                _src10.append(src('监理台账/超前水平钻', obj['r'],
                                  'c2/c3 起止里程、c4 检测长度、c5 施做时间、c6 时间、'
                                  'c7 下次施做里程、c8 搭接'))
            else:
                _g0 = obj['groups'][0]
                _src10.append(src('监理台账/加深炮孔', _g0.get('r'),
                                  '组首行 c2 钻孔位置、c3 日期、c4 时间、c5 孔号、c6 钻孔长度、'
                                  'c7 终孔位置、c8 外插角（表头未标该列）；同组后续行只有孔号与孔深',
                                  _g0.get('r_last')))
        tables[10] = make_table(10, proj_of('超前水平钻' if '超前水平钻' in data else '加深炮孔'),
                                d10, '；'.join(c10),
                                tf(min(d10s), d10), tf(max(e10s), d10), recs10,
                                source=[x for x in _src10 if x])
    else:
        tables[10] = make_table(10, '', None, '', '', '', [], '台账无超前水平钻/加深炮孔时间记录')

    # ============ 11 二衬预埋槽道 ============
    if '接触网槽道' in data:
        best = None
        for r, vals in data['接触网槽道']:
            if r < 2:
                continue
            sp = parse_span(_g(vals, 10), base=parse_date(_g(vals, 9)))
            if not sp:
                continue
            if best is None or sp['start'] >= best['start']:
                best = {'start': sp['start'], 'end': sp['end'], 'vals': vals, 'r': r}
        if best:
            v = best['vals']
            cnt = _g(v, 5)
            cnt_s = f'{int(cnt)}' if isinstance(cnt, float) and cnt == int(cnt) else _clean_txt(cnt, 6)
            recs11 = ['槽道安装：二衬里程 {}；槽道里程 {}；型号 {}；长度 {} mm；根数 {}；'
                      '安装位置 {}；槽道间距 {} mm；中心距施工缝 {}。'.format(
                          mk(_clean_txt(_g(v, 1), 20)), mk(dk(_g(v, 2))), mk(_clean_txt(_g(v, 3), 8)),
                          mk(num_s(_g(v, 4))), mk(cnt_s or '/'),
                          mk(_clean_txt(_g(v, 6), 12)), mk(num_s(_g(v, 7))),
                          mk(_clean_txt(_g(v, 8), 8))),
                      '安装时间：{}。'.format(mk(_norm(_g(v, 10))))]
            tables[11] = make_table(11, proj_of('接触网槽道'),
                                    best['start'], f'二衬 {mk(_clean_txt(_g(v, 1), 20))}　槽道里程 {mk(dk(_g(v, 2)))}',
                                    tf(best['start'], best['start']),
                                    tf(best['end'] or best['start'], best['start']), recs11,
                                    source=[src('监理台账/接触网槽道', best['r'],
                                                'c2 二衬里程、c3 槽道里程、c4 型号、c5 长度、c6 根数、'
                                                'c7 安装位置、c8 槽道间距、c9 中心距施工缝、'
                                                'c10 施工日期、c11 时间')])
        else:
            tables[11] = make_table(11, '', None, '', '', '', [], '台账无槽道安装时间记录')
    else:
        tables[11] = make_table(11, '', None, '', '', '', [], '台账无槽道 sheet')

    # ============ 12 综合接地 ============
    gr_rows = []
    for name in ('暗洞二衬', '仰拱'):
        if name not in data:
            continue
        gi, gr = (16, 18) if name == '暗洞二衬' else (14, 16)
        for r, vals in data[name]:
            if r < 3:
                continue
            sp = parse_span(_g(vals, gi))
            if not sp:
                continue
            gr_rows.append({'r': r, 'week': sp, 'date': sp['start'].date(), 'start': sp['start'],
                            'end': sp['end'] or sp['start'], 'sheet': name, 'vals': vals,
                            'g': _norm(_g(vals, gi)), 'res': _clean_txt(_g(vals, gr), 12),
                            'sup': _clean_txt(_g(vals, gr + 1), 12)})
    if gr_rows:
        latest = max(x['date'] for x in gr_rows)
        hits = [x for x in gr_rows if x['date'] == latest]
        content, recs = [], []
        for x in hits:
            seg = _mile(x['vals'])
            content.append(f'{x["sheet"]} {mk(seg)}' if seg else x['sheet'])
            parts = [f'接地钢筋安装 {mk(x["g"])}']
            if x['res']:
                parts.append(f'接地电阻 {mk(x["res"])}')
            if x['sup']:
                parts.append(f'监理报验 {mk(x["sup"])}')
            recs.append(f'{x["sheet"]}（{mk(seg) if seg else "/"}）：' + '；'.join(parts) + '。')
        _s12 = min(x['start'] for x in hits)
        tables[12] = make_table(12, proj_of(hits[0]['sheet']), _s12, '；'.join(content),
                                tf(_s12, _s12), tf(max(x['end'] for x in hits), _s12), recs,
                                source=[src(f'监理台账/{x["sheet"]}', x['r'],
                                            'c17 接地安装、c19 接地电阻、c20 监理报验'
                                            if x['sheet'] == '暗洞二衬' else
                                            'c15 接地安装、c17 接地电阻、c18 监理报验')
                                        for x in hits])
    else:
        tables[12] = make_table(12, '', None, '', '', '', [], '台账无接地安装时间记录')

    # ============ 3 超前小导管注浆 ============
    adv = None
    for name in names:
        if '超前' not in name:
            continue
        rows = data[name]
        head = ' '.join(_norm(v) for _r, vals in rows[:2] for v in vals if _norm(v))
        if '超前支护' not in head and '小导管' not in head:
            continue
        for r, vals in rows:
            if r < 2:
                continue
            base = parse_date(_g(vals, 5))
            if base is None and _norm(_g(vals, 5)):
                m = DATE_PAT.search(_norm(_g(vals, 5)))
                if m:
                    base = parse_date(m.group(0))
            sp = parse_span(_g(vals, 6), base=base)
            if not sp and base:
                sp = {'start': base, 'end': None}
            if not sp:
                continue
            if adv is None or sp['start'] >= adv['start']:
                adv = {'start': sp['start'], 'end': sp['end'], 'base': base, 'vals': vals,
                       'sheet': name, 'r': r}
    if adv:
        v = adv['vals']
        recs3 = ['超前小导管注浆：里程 {}；超前类型 {}；超前长度 {} m；根数 {}；'
                 '注浆时间 {}；注浆量 {} m³；搭接 {} m。'.format(
                     mk(f'{dk(_g(v, 1))}～{dk(_g(v, 2))}'), mk(_clean_txt(_g(v, 3), 12)),
                     mk(_clean_txt(_g(v, 4), 8)), mk(num_s(_g(v, 9))) if _g(v, 9) not in ('', None) else '/',
                     mk(_norm(_g(v, 6))), mk(num_s(_g(v, 7))) if _g(v, 7) not in ('', None) else '/',
                     mk(_clean_txt(_g(v, 8), 8)))]
        seg3 = f'{dk(_g(v, 1))}～{dk(_g(v, 2))}'
        c3_txt = f'{_clean_txt(_g(v, 3), 12) or "超前小导管"} {mk(seg3)}'
        if _g(v, 9) not in ('', None):
            c3_txt += f'（{mk(num_s(_g(v, 9)))} 根）'
        tables[3] = make_table(3, proj_of(adv['sheet']), adv['start'], c3_txt,
                               tf(adv['start'], adv['start']),
                               tf(adv['end'] or adv['start'], adv['start']), recs3,
                               source=[src(f'监理台账/{adv["sheet"]}', adv['r'],
                                           'c2/c3 起止里程、c4 超前类型、c5 超前长度、c6 施做日期、'
                                           'c7 注浆时间、c8 注浆量、c9 搭接、c10 根数')])
    else:
        tables[3] = make_table(3, '', None, '', '', '', [], '台账无超前小导管注浆时间记录')

    # ============ 6 二衬带模注浆 ============
    tm = None
    if '暗洞二衬' in data:
        for r, vals in data['暗洞二衬']:
            if r < 3:
                continue
            sp = parse_span(_g(vals, 26))
            if not sp:
                continue
            if tm is None or sp['start'] >= tm['start']:
                tm = {'start': sp['start'], 'end': sp['end'], 'vals': vals, 'r': r}
    if tm:
        v = tm['vals']
        seg6 = _mile(v)
        recs6 = ['二衬带模注浆：衬砌段 {}；砼强度 {}；注浆时间 {}；带模注浆水泥浆用量 {} m³。'.format(
            mk(seg6) if seg6 else '/', mk(_clean_txt(_g(v, 7), 8)),
            mk(_norm(_g(v, 26))), mk(_g(v, 27)) if _g(v, 27) not in ('', None) else '/')]
        tables[6] = make_table(6, proj_of('暗洞二衬'), tm['start'],
                               f'二衬 {mk(seg6)}' if seg6 else '/',
                               tf(tm['start'], tm['start']),
                               tf(tm['end'] or tm['start'], tm['start']), recs6,
                               source=[src('监理台账/暗洞二衬', tm['r'],
                                           'c8 砼设计标号、c27 注浆时间、c28 带模注浆水泥浆用量')])
    else:
        tables[6] = make_table(6, '', None, '', '', '', [], '台账无二衬带模注浆时间记录')

    return tables, warnings


def _span_date_fallback(vals):
    """给没有日期列的 sheet 兜底：从整行文本里抠第一个日期。"""
    for v in vals:
        if isinstance(v, datetime.datetime):
            return v
        m = DATE_PAT.search(_norm(v))
        if m:
            return parse_date(m.group(0))
    return None


# ---------------------------------------------------------------- 注浆台账 --

def scan_zhupu(book, mode):
    """注浆台账：每 sheet 一张独立记录表，标题分桶 → 2 砂浆锚杆 / 4 中空锚杆 / 5 锁脚锚管。"""
    tables = {}
    warnings = []
    names = sheet_names(book, mode)
    buckets = {2: [], 4: [], 5: []}

    for name in names:
        rows = sheet_rows(book, mode, name, max_cells=40)
        title = ''
        meta = {'date': None, 'no': '', 'part': '', 'mileage': '', 'equip': '', 'proj': '',
                'rows': 0, 'mix': '', 'rock': '', 'press': [], 'vol': [],
                't0': None, 't1': None, 'site': ''}
        for r, vals in rows[:6]:
            joined = ''.join(_norm(v) for v in vals if _norm(v))
            squeezed = ''.join(joined.split())
            if '记录表' in squeezed and not title:
                title = squeezed
            if '注浆日期' in joined:
                for v in vals:
                    d = parse_date(v)
                    if d is not None:
                        meta['date'] = d
            if '编' in joined and '号' in joined:
                m2 = re.search(r'(TJSSDJK-[A-Z]+-\d{8}-\d+)', joined)
                if m2:
                    meta['no'] = m2.group(1)
            if '工程部位' in joined:
                for i, v in enumerate(vals):
                    if _norm(v) == '工程部位':
                        nxt = next((x for x in vals[i + 1:] if _norm(x)), None)
                        if nxt is not None:
                            meta['part'] = _norm(nxt)
                    elif _norm(v) == '工程里程':
                        after = [x for x in vals[i + 1:] if _norm(x)]
                        nums = [dk(x) for x in after[:3] if isinstance(x, (int, float))]
                        if len(nums) >= 2:
                            meta['mileage'] = f'{nums[0]}～{nums[1]}'
                        elif nums:
                            meta['mileage'] = nums[0]
            if '工程名称' in joined:
                for i, v in enumerate(vals):
                    if _norm(v) == '工程名称':
                        nxt = next((x for x in vals[i + 1:] if _norm(x)), None)
                        if nxt is not None:
                            meta['proj'] = _norm(nxt)
            if '注浆设备' in joined:
                for i, v in enumerate(vals):
                    if _norm(v) == '注浆设备':
                        nxt = next((x for x in vals[i + 1:] if _norm(x)), None)
                        if nxt is not None:
                            meta['equip'] = _norm(nxt)
        if not title:
            continue
        col, data_start = {}, None
        for r, vals in rows:
            texts = [_norm(v) for v in vals]
            if any('锚杆编号' in t for t in texts):
                for i, t in enumerate(texts):
                    if t == '锚杆编号':
                        col['no'] = i
                    elif '配合比' in t:
                        col['mix'] = i
                    elif '压力' in t:
                        col['press'] = i
                    elif '注浆量' in t:
                        col['vol'] = i
                    elif '开始时间' in t:
                        col['t0'] = i
                    elif '终止时间' in t:
                        col['t1'] = i
                    elif t == '围岩等级':
                        col['rock'] = i
                    elif t == '注浆部位':
                        col['site'] = i
                data_start = r + 1
                break
        if data_start is None:
            continue
        tmin = tmax = None
        for r, vals in rows[data_start:]:
            no_v = _g(vals, col.get('no', 0)) if col.get('no', 0) >= 0 else ''
            if str(no_v).strip() not in ('', 'None'):
                if meta.get('row0') is None:
                    meta['row0'] = r        # 该 sheet 第一根锚杆所在行（0 基）
                meta['row1'] = r            # 最后一根锚杆所在行，用于标出行区间
            if no_v in ('', None) or not re.match(r'^\d+(\.0)?$', str(no_v).strip()):
                continue    # 只有「锚杆编号」列是数字的行才算一根锚杆
            meta['rows'] += 1
            if col.get('mix') is not None and '水泥' in _norm(_g(vals, col['mix'])):
                meta['mix'] = _norm(_g(vals, col['mix']))
            if col.get('rock') is not None and not meta['rock'] and _norm(_g(vals, col['rock'])):
                meta['rock'] = _norm(_g(vals, col['rock']))
            if col.get('site') is not None and not meta['site'] and _norm(_g(vals, col['site'])):
                meta['site'] = _norm(_g(vals, col['site']))
            if col.get('press') is not None:
                try:
                    meta['press'].append(float(_norm(_g(vals, col['press']))))
                except (ValueError, TypeError):
                    pass
            if col.get('vol') is not None:
                try:
                    meta['vol'].append(float(_norm(_g(vals, col['vol']))))
                except (ValueError, TypeError):
                    pass
            if col.get('t0') is not None:
                tv = _as_time(_g(vals, col['t0']))
                if tv:
                    tmin = tv if tmin is None else min(tmin, tv)
            if col.get('t1') is not None:
                tv = _as_time(_g(vals, col['t1']))
                if tv:
                    tmax = tv if tmax is None else max(tmax, tv)
        meta['t0'], meta['t1'] = tmin, tmax
        no = None
        if '锁脚' in title:
            no = 5
        elif '中空' in title:
            no = 4
        elif '锚杆注浆' in title:
            no = 2 if '砂' in (meta['mix'] or '') else 4
        if no:
            buckets[no].append((meta['date'], name, title, meta))

    for no in (2, 4, 5):
        items = [x for x in buckets[no] if x[0] is not None and x[3]['rows'] > 0]
        if not items:
            tables[no] = make_table(no, '', None, '', '', '', [], '注浆台账无该类表页数据')
            continue
        latest = max(x[0] for x in items)
        same = sorted([x for x in items if x[0] == latest], key=lambda x: x[1])
        recs, content, starts, ends = [], [], [], []
        for d, sname, t, meta in same:
            parts = [f'锚杆 {mk(meta["rows"])} 根']
            if meta['no']:
                parts.append(f'记录表编号 {mk(meta["no"])}')
            if meta['part']:
                parts.append(f'部位 {mk(meta["part"])}')
            if meta['mileage']:
                parts.append(f'里程 {mk(meta["mileage"])}')
            if meta['rock']:
                parts.append(f'围岩 {mk(meta["rock"])}')
            if meta['site']:
                parts.append(f'注浆部位 {mk(meta["site"])}')
            if meta['mix']:
                parts.append(f'配合比 {mk(meta["mix"])}')
            if meta['t0'] and meta['t1']:
                parts.append(f'注浆时间 {mk(meta["t0"].strftime("%H:%M"))}～{mk(meta["t1"].strftime("%H:%M"))}')
                starts.append(meta['t0'])
                ends.append(meta['t1'])
            if meta['vol']:
                vols = [f'{v:g}' for v in meta['vol']]
                shown = '、'.join(vols[:3]) + (f' 等{len(vols)}根' if len(vols) > 3 else '')
                total_s = '%.2f' % sum(meta['vol'])
                total_s = total_s.rstrip('0').rstrip('.')
                # 合计是本脚本加出来的数，台账里没有这一格 —— 不标红，避免把推算值说成台账值
                parts.append(f'注浆量 {mk(shown)} L（合计 {total_s} L，推算）')
            if meta['press']:
                pm, px = min(meta['press']), max(meta['press'])
                parts.append(f'注浆压力 {mk(f"{pm:g}～{px:g}")} MPa')
            recs.append(f'{meta["part"] or sname}：' + '；'.join(parts) + '。')
            content.append(f'{meta["part"] or sname}' + (f'（{mk(meta["mileage"])}）' if meta['mileage'] else ''))
        d0 = latest
        s_dt = datetime.datetime.combine(d0.date(), min(starts)) if starts else None
        e_dt = datetime.datetime.combine(d0.date(), max(ends)) if ends else None
        _pj = same[0][3].get('proj') or '淘金山隧道进口'
        _pj = re.sub(r'(进口|出口|斜井)$', '', _pj) or '淘金山隧道'
        _src245 = []
        for d, sname, t, meta in same:
            _src245.append(src(f'注浆台账/{sname}', 1,
                               '第 2 行记录表标题、第 3 行记录表编号、第 4 行工程部位与工程里程、'
                               '第 5 行注浆日期、第 6 行列头', 5))
            _src245.append(src(f'注浆台账/{sname}', meta.get('row0'),
                               '逐根数据：编号、围岩等级、注浆部位、配合比、注浆起止时间、'
                               '注浆压力、注浆量', meta.get('row1')))
        tables[no] = make_table(no, _pj, d0, '；'.join(content),
                                tf(s_dt, d0), tf(e_dt, d0), recs,
                                source=[x for x in _src245 if x])
    return tables, warnings


# ---------------------------------------------------------------- 配料单 ----

def scan_peiliao(book, mode, pours=None):
    """配料单台账 → 表 8 混凝土浇筑。只取「结构混凝土」（排除喷射），
    浇筑起止时间从监理台账对应浇筑时间补。"""
    tables = {}
    warnings = []
    names = sheet_names(book, mode)
    main_name = next((n for n in names if '开盘信息汇总' in n), names[0])
    rows = sheet_rows(book, mode, main_name, max_cells=120)
    head_idx = None
    for r, vals in rows:
        if vals and _norm(vals[0]) == '序号':
            head_idx = r
            break
    if head_idx is None:
        tables[8] = make_table(8, '', None, '', '', '', [], '开盘信息汇总未找到表头行')
        return tables, ['未找到开盘信息汇总表头行']

    cand = []
    for r, vals in rows[head_idx + 1:]:
        d = vals[2] if len(vals) > 2 else None
        if not isinstance(d, datetime.datetime):
            continue
        unit = str(vals[3]) if len(vals) > 3 and vals[3] else ''
        if '淘金山' not in unit:
            continue
        grade = _clean_txt(_g(vals, 5), 20)
        part = _clean_txt(_g(vals, 4), 60)
        cand.append({'d': d, 'vals': vals, 'grade': grade, 'part': part, 'r': r})
    if not cand:
        tables[8] = make_table(8, '', None, '', '', '', [], '配料单无淘金山记录')
        return tables, warnings

    # 结构混凝土（非喷射）优先；没有就退回全部
    struct = [x for x in cand if '喷射' not in x['grade'] and '喷' not in x['part']]
    pool = struct or cand
    latest = max(x['d'] for x in pool)
    hits = [x for x in pool if x['d'] == latest]
    row = hits[0]
    v = row['vals']
    g = lambda i: _clean_txt(_g(v, i), 60) or '/'

    # 从监理台账补浇筑起止
    s_dt = e_dt = None
    pour_hit = None
    for p in (pours or []):
        if p['date'].date() != latest.date():
            continue
        kw = ''
        for k in ('二衬', '仰拱', '填充', '通信洞室', '挡墙'):
            if k in row['part']:
                kw = k
                break
        if kw and kw not in (p.get('label') or '') and kw not in (p.get('sheet') or ''):
            continue
        s_dt, e_dt = p['start'], p['end'] or p['start']
        pour_hit = p
        break

    recs = ['浇筑部位 {}；强度等级 {}；计划方量 {} m³、实际方量 {} m³。'.format(
                mk(_clean_txt(_g(v, 4), 60)), mk(row['grade'] or '/'),
                mk(_clean_txt(_g(v, 6), 12)), mk(_clean_txt(_g(v, 7), 12))),
            '通知单号 {}；配合比报告编号 {}；砂率 {}；外加剂掺量 {}；搅拌机组 {}。'.format(
                mk(_clean_txt(_g(v, 1), 40)), mk(_clean_txt(_g(v, 10), 40)),
                mk(_clean_txt(_g(v, 12), 10)), mk(_clean_txt(_g(v, 16), 10)),
                mk(_clean_txt(_g(v, 19), 12))),
            ]
    # 原材料检验报告：只写台账里有报告号的材料，空/「/」的不写
    _reps = [('水泥', 28), ('粉煤灰', 29), ('细骨料', 31), ('粗骨料', 32),
             ('外加剂', 35), ('速凝剂', 36), ('水', 37)]
    _rep_txt = [f'{n} {mk(t)}' for n, i in _reps
                for t in [_clean_txt(_g(v, i), 40)] if t and t != '/']
    if _rep_txt:
        recs.append('原材料检验报告：' + '；'.join(_rep_txt) + '。')
    slump = _clean_txt(_g(v, 39), 12)
    air = _clean_txt(_g(v, 41), 12)
    t_out = _clean_txt(_g(v, 43), 12)
    if slump != '/':
        recs.append(f'首盘坍落度 {mk(slump)} mm' + (f'；含气量 {mk(air)}' if air != '/' else '')
                    + (f'；出机温度 {mk(t_out)} ℃' if t_out != '/' else '') + '。')
    if s_dt:
        recs.insert(0, f'浇筑起止时间（监理台账）：{mk(s_dt.strftime("%Y-%m-%d %H:%M"))} ～ '
                       f'{mk((e_dt or s_dt).strftime("%Y-%m-%d %H:%M"))}。')

    # 项目名（隧道名）取「单位工程」列 c4（实测值为「淘金山隧道」），
    # 不取 c10「工程名称」——那一列全项目都写「隧道工程」，取不到隧道名。
    _pj8 = _clean_txt(_g(v, 3), 40)
    _pj8 = re.sub(r'(进口|出口|斜井)$', '', _pj8)
    if not _pj8 or '隧道' not in _pj8:
        _pj8 = _clean_txt(_g(v, 9), 40)
        _pj8 = re.sub(r'(进口|出口|斜井)$', '', _pj8)
    if not _pj8 or '隧道' not in _pj8 or _pj8 == '隧道工程':
        _pj8 = ''
    _src8 = [src('配料单/开盘信息汇总', row['r'],
                 'c2 通知单号、c3 浇筑日期、c4 单位工程、c5 浇筑部位、c6 砼强度等级、'
                 'c7 计划方量、c8 实际方量、c10 工程名称、c11 配合比报告编号、c13 砂率、'
                 'c17 外加剂掺量、c20 搅拌机组、c29/c30/c32/c33/c36/c37/c38 原材料报告、'
                 'c40 首盘坍落度、c42 含气量、c44 出机温度')]
    if s_dt and pour_hit:
        _src8.append(src(f'监理台账/{pour_hit["sheet"]}', pour_hit.get('r'),
                         f'c{pour_hit.get("col", 0) + 1} 混凝土浇筑时间（旁站起止时间来源）'))
    tables[8] = make_table(8, _pj8, latest,
                           mk(_clean_txt(_g(v, 4), 60)), tf(s_dt, latest) if s_dt else '/',
                           tf(e_dt, latest) if e_dt else '/', recs,
                           '' if s_dt else '配料单无浇筑起止时间，且监理台账未匹配到同日浇筑记录',
                           source=[x for x in _src8 if x])
    return tables, warnings


# ------------------------------------------------------------------ main --

def main() -> int:
    ap = argparse.ArgumentParser(
        description='台账 → 12 张旁站记录的内容（JSON）。可一次传多份台账，自动合并。')
    ap.add_argument('ledgers', nargs='+',
                    help='台账文件（.xls/.xlsx），可多份：监理台账 / 旁站注浆记录表 / 配料单及拌合物')
    ap.add_argument('--type', default='',
                    help='仅当文件名认不出时手动指定（监理台账/旁站注浆记录表/配料单及拌合物）')
    ap.add_argument('--pretty', action='store_true', help='JSON 缩进输出')
    a = ap.parse_args()

    books = {}
    for p in a.ledgers:
        t = a.type or detect_type(p)
        if not t:
            print(f'认不出台账类型：{p}（文件名需含 监理台账/旁站注浆记录表/配料单及拌合物）',
                  file=sys.stderr)
            return 3
        try:
            book, mode = open_book(p)
        except Exception as e:
            print(f'读台账失败：{e}', file=sys.stderr)
            return 2
        books[t] = (book, mode, p)

    tables = {}
    warnings = []
    pours = []
    if '监理台账' in books:
        book, mode, _p = books['监理台账']
        pours = _collect_pours(books)
        t, w = scan_jianli(book, mode)
        tables.update(t)
        warnings += w
    if '旁站注浆记录表' in books:
        book, mode, _p = books['旁站注浆记录表']
        t, w = scan_zhupu(book, mode)
        tables.update(t)
        warnings += w
    if '配料单及拌合物' in books:
        book, mode, _p = books['配料单及拌合物']
        t, w = scan_peiliao(book, mode, pours)
        tables.update(t)
        warnings += w

    # 补齐缺的表（没给对应台账时）
    for no in range(1, 13):
        if no not in tables:
            tables[no] = make_table(no, '', None, '', '', '', [], '未提供该类台账')

    ordered = [tables[i] for i in range(1, 13)]
    missing = [f'{tb["no"]} {tb["title"]}：{tb.pop("missing_reason")}'
               for tb in ordered if 'missing_reason' in tb]
    dates = [parse_date(tb['fields']['date']) for tb in ordered if tb['fields']['date'] != '/']
    dates = [d for d in dates if d]
    out = {'ledger': '、'.join(a.ledgers),
           'type': '、'.join(books.keys()),
           'ledger_date': fmt_date(max(dates)) if dates else None,
           'tables': ordered, 'missing': missing, 'warnings': warnings}
    print(json.dumps(out, ensure_ascii=False, indent=2 if a.pretty else None))

    has_data = any('⟦' in r and '/' not in r.replace('⟦', '').replace('⟧', '')
                   for tb in ordered for r in tb['records'])
    return 0 if (dates and has_data) else 1


if __name__ == '__main__':
    sys.exit(main())
