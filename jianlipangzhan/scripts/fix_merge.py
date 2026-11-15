#!/usr/bin/env python3
"""修复 Univer 导出 docx 时的合并单元格残留空 <w:tc>。

背景：Univer 的 mergeCells 只在逻辑层合并；导出 docx 时会把被吞掉的那一格
重新吐成一个空的 <w:tc>，导致该行 gridSpan 之和 > tblGrid 列数，Word 里
表现为多出一条竖线。

本脚本只做一件事：逐行统计 gridSpan 之和，若大于表格网格列数，则删除该行
里内容为空的 <w:tc>，直到之和等于网格列数。其余 XML 一字不改。

用法：
    python3 fix_merge.py <in.docx> <out.docx>
退出码：0 = 无需修复或已修复并自检通过；1 = 修复后仍不自洽；2 = 用法/文件错误
"""
import re
import sys
import zipfile

TBL_RE = re.compile(r'<w:tbl>.*?</w:tbl>', re.S)
ROW_RE = re.compile(r'<w:tr[ >].*?</w:tr>', re.S)
TC_RE = re.compile(r'<w:tc>.*?</w:tc>', re.S)


def span_of(tc: str) -> int:
    m = re.search(r'<w:gridSpan w:val="(\d+)"', tc)
    return int(m.group(1)) if m else 1


def text_of(tc: str) -> str:
    return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tc))


def fix_row(row: str, ncols: int):
    cells = TC_RE.findall(row)
    total = sum(span_of(c) for c in cells)
    if total <= ncols:
        return row, 0
    out = list(cells)
    removed = 0
    for i in range(len(out) - 1, -1, -1):
        if total <= ncols:
            break
        if text_of(out[i]).strip() == '' and span_of(out[i]) == 1:
            total -= 1
            del out[i]
            removed += 1
    if total != ncols:
        return row, -1
    new_row = row[:row.index('<w:tc>')] + ''.join(out) + row[row.rindex('</w:tc>') + len('</w:tc>'):]
    return new_row, removed


def main() -> int:
    if len(sys.argv) != 3:
        print('usage: fix_merge.py <in.docx> <out.docx>', file=sys.stderr)
        return 2
    src, dst = sys.argv[1], sys.argv[2]
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}

    xml = blobs['word/document.xml'].decode('utf-8')
    total_removed = 0
    bad = 0

    def fix_table(m):
        nonlocal total_removed, bad
        table = m.group(0)
        grid = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', table, re.S)
        ncols = len(re.findall(r'<w:gridCol', grid.group(1))) if grid else 0
        if not ncols:
            return table

        def fix_tr(rm):
            nonlocal total_removed, bad
            new, removed = fix_row(rm.group(0), ncols)
            if removed < 0:
                bad += 1
            total_removed += max(removed, 0)
            return new

        return ROW_RE.sub(fix_tr, table)

    xml = TBL_RE.sub(fix_table, xml)

    # 自检：修复后每行 gridSpan 之和必须 == 网格列数
    ok = True
    for t in TBL_RE.findall(xml):
        grid = re.search(r'<w:tblGrid>(.*?)</w:tblGrid>', t, re.S)
        ncols = len(re.findall(r'<w:gridCol', grid.group(1))) if grid else 0
        for r in ROW_RE.findall(t):
            if sum(span_of(c) for c in TC_RE.findall(r)) != ncols:
                ok = False

    blobs['word/document.xml'] = xml.encode('utf-8')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, blobs[n])

    print(f'removed_empty_tc={total_removed} unrepairable_rows={bad} selfcheck={"PASS" if ok else "FAIL"}')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
