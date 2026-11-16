#!/usr/bin/env python3
"""把 extract_latest.py 的 JSON 填进 20 源行 docx 模板（每个类型一份），供 postprocess.py 后处理。

用法：python3 fill_docs.py <extract.json> <模板.docx> <输出目录> [--no 1,7]
- 模板 = assets/build_form.js 生成的 20 源行 docx（Univer 导出，未后处理）
- 输出 = <输出目录>/raw-<序号>.docx（20 源行、内容已替换、含 ⟦⟧ 记号）
退出码：0=全部生成；1=行数不符；2=用法错误。
"""
import json
import re
import sys
import zipfile
from pathlib import Path

ROW_RE = re.compile(r'<w:tr\b.*?</w:tr>', re.S)
TC_RE = re.compile(r'<w:tc>.*?</w:tc>', re.S)
TBL_RE = re.compile(r'<w:tbl>.*?</w:tbl>', re.S)


def texts_of(tc):
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', tc))


def set_row_texts(row, texts):
    """把一行的各格按顺序改写成 texts（超出的格清空）；按位置重建，不用 replace 防撞重复。"""
    out, pos = [], 0
    idx = 0
    for m in TC_RE.finditer(row):
        out.append(row[pos:m.start()])
        c = m.group(0)
        t = texts[idx] if idx < len(texts) else ''
        idx += 1
        esc = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        # 清掉所有成对 <w:t>…</w:t> 的内容
        new_c = re.sub(r'(<w:t(?: [^>]*)?>).*?(</w:t>)', r'\1\2', c, flags=re.S)
        if esc:
            r_m = re.search(r'<w:r>(<w:rPr>.*?</w:rPr>)?', new_c, re.S)
            insert = f'<w:t xml:space="preserve">{esc}</w:t>'
            if r_m:
                new_c = new_c[:r_m.end()] + insert + new_c[r_m.end():]
            else:
                p_m = re.search(r'(<w:pPr>.*?</w:pPr>)', new_c, re.S)
                rpr = ''
                if p_m:
                    rm = re.search(r'<w:rPr>.*?</w:rPr>', p_m.group(1), re.S)
                    rpr = rm.group(0) if rm else ''
                run = f'<w:r>{rpr}<w:t xml:space="preserve">{esc}</w:t></w:r>'
                if p_m:
                    new_c = new_c.replace(p_m.group(1), p_m.group(1) + run, 1)
                else:
                    new_c = new_c.replace('</w:p>', run + '</w:p>', 1)
        out.append(new_c)
        pos = m.end()
    out.append(row[pos:])
    return ''.join(out)


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description='extract JSON → 20 源行 docx（每类型一份）')
    ap.add_argument('extract', help='extract_latest.py 的 JSON 输出')
    ap.add_argument('template', help='20 源行 docx 模板（build_form.js 产物）')
    ap.add_argument('outdir', help='输出目录')
    ap.add_argument('--no', default='', help='只生成这些序号（逗号分隔），默认 tables 里的全部')
    ap.add_argument('--title-date', default='', help='覆盖首行日期文本（默认用 fields.date）')
    a = ap.parse_args()
    if len(sys.argv) < 4:
        return 2

    data = json.load(open(a.extract))
    outdir = Path(a.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    only = {int(x) for x in a.no.split(',') if x.strip()} if a.no else None

    with zipfile.ZipFile(a.template) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    xml = blobs['word/document.xml'].decode('utf-8')
    tbl = TBL_RE.search(xml)
    if not tbl:
        print('模板里没有表格', file=sys.stderr)
        return 1
    rows = ROW_RE.findall(tbl.group(0))
    if len(rows) != 20:
        print(f'模板行数 {len(rows)} != 20', file=sys.stderr)
        return 1

    made = []
    for t in data['tables']:
        no = t['no']
        if only and no not in only:
            continue
        f = t['fields']
        # —— 记录条目整理：JSON 里「工程部位：…」已折进记录开头（同一条），
        # 无需再折；剥旧编号后统一重编；空记录留白（不填「/」）
        recs = []
        for r in t['records']:
            r2 = re.sub(r'^\d+、', '', r) if re.match(r'^\d+、', r) else r
            if r2.strip() and r2.strip() != '/':
                recs.append(r2)
        recs = recs or []
        recs = [f'{i}、{r}' for i, r in enumerate(recs, 1)]
        new_rows = list(rows)
        date_txt = a.title_date or (f.get('date') or '')
        if date_txt not in ('/', ''):
            date_txt = f'\u27e6{date_txt}\u27e7'
        new_rows[0] = set_row_texts(rows[0], ['日\u3000期', date_txt, '天\u3000气',
                                              f.get('weather', ''), '环境温度', f.get('temperature', '')])
        # 项目名称（隧道名）来自台账标题/工程名称 → 标红
        project = (f.get('project') or '').strip()
        proj_txt = f'\u27e6{project}\u27e7' if project and project != '/' else ''
        # 旁站内容整段放进同一个格（r1 的第 4 格），r2/r3 留空；
        # postprocess 会把这三行的第 4 格纵向合并成一格
        content = (f.get('content') or '').strip()
        # 旁站内容整段都来自台账，整段标红（不再逐段包记号）
        if content != '/' and '\u27e6' not in content:
            content = f'\u27e6{content}\u27e7'
        new_rows[1] = set_row_texts(rows[1], ['旁站项目名称', proj_txt, '', '旁站内容', content])
        new_rows[2] = set_row_texts(rows[2], ['', '', '', '', ''])
        new_rows[3] = set_row_texts(rows[3], ['', '', '', '', ''])
        new_rows[4] = set_row_texts(rows[4], ['旁站开始时间', f.get('start', ''), '',
                                              '旁站结束时间', f.get('end', '')])
        new_rows[5] = set_row_texts(rows[5], ['旁站过程记录：'])
        # 模板 r6-r13 共 8 格：超过 8 条的记录合并进第 8 格（用空行分隔），不丢内容
        for i in range(8):
            if i < 7:
                txt = recs[i] if i < len(recs) else ''
            else:
                tail = recs[7:] if len(recs) > 8 else (recs[7] if len(recs) == 8 else '')
                txt = '\u3000'.join(tail) if isinstance(tail, list) else tail
            new_rows[6 + i] = set_row_texts(rows[6 + i], [txt])
        # 其他 / 发现 / 间隔 / 签字：模板原样保留
        new_tbl = tbl.group(0)
        for old, new in zip(rows, new_rows):
            if old != new:
                new_tbl = new_tbl.replace(old, new, 1)
        new_xml = xml[:tbl.start()] + new_tbl + xml[tbl.end():]
        # 标题：把模板标题里的「（隧道开挖）」换成本类型名
        new_xml = new_xml.replace('旁站记录表（隧道开挖）',
                                  f'旁站记录表（{t["title"]}）', 1)
        out = outdir / f'raw-{no}.docx'
        blobs['word/document.xml'] = new_xml.encode('utf-8')
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in names:
                z.writestr(n, blobs[n])
        made.append(str(out))
    print('made:', ' '.join(made))
    counts = {str(t['no']): len(t['records']) for t in data['tables']
              if not only or t['no'] in only}
    print('record_counts:', json.dumps(counts, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
