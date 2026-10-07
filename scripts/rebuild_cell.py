# -*- coding: utf-8 -*-
"""重建 docx 中某个表格单元格的段落序列（格式完全沿用原单元格模板）。

典型场景：**开题报告 / 任务书**这类"整份正文装在一张大表格里"的文档 ——
`document.xml` 的 body 只有几个段落 + 1 个表格，每个小节是表格的一行。
要改写某小节（例如"研究的主要内容和思路"改成章节大纲框架），
按段落序号定位会完全错位，必须按**行**定位。

本脚本做三件事，且只改目标单元格：

1. 保留该单元格的前 `--keep-first N` 个段落（通常是小节标签）；
2. 用原单元格里的段落作模板，按 spec 文件重建其余段落（`deepcopy` 模板 → 只替换文字），
   因此新段落的 `pPr` / run 的 `rPr`（字体、字号、行距、缩进）与原文**逐项一致**；
3. 可选地做一批"精确文本替换"（用于顺手修掉残缺碎片），并断言每处**恰好命中 1 次**
   —— 命中 0 次说明定位串写错了，命中 >1 次说明会误伤，都必须让脚本失败而不是静默无改动。

spec 文件格式（UTF-8，一行一个段落）：

```
第1章 绪论                 ← 顶格段落（按 --flush-template 的格式）
|1.1 研究背景与意义        ← 行首 `|` = 缩进段落（按 --indent-template 的格式，`|` 本身不写入）
|1.2 国内外研究现状
# 以 # 开头或以空白行 = 注释/空行，跳过
```

用法：
    python rebuild_cell.py <src.docx> <out.docx> --spec spec.txt \\
        [--table 0] [--find "研究的主要内容和思路"] [--row 6] [--cell 0] \\
        [--keep-first 1] [--indent-template 1] [--flush-template 0] \\
        [--replace "有待拓展国家哲学社。" --replace-with "有待拓展。"]

`--find` 与 `--row` 二选一：`--find` 在表格各行的**首个单元格文本**里做子串匹配，
更稳（行序变了也不怕）。默认 `--table 0` `--cell 0`（正文首表、单元格 0）。
"""
import argparse
import copy
import os
import sys
import zipfile

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_NS = '{http://www.w3.org/XML/1998/namespace}'


def qn(tag):
    return W + tag


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(qn('t')))


def iter_body_tables(root):
    body = root.find(qn('body'))
    for el in body:
        if el.tag == qn('tbl'):
            yield el


def load_spec(path):
    out = []
    with open(path, 'r', encoding='utf-8') as f:
        for raw in f:
            line = raw.rstrip('\r\n')
            if not line.strip() or line.lstrip().startswith('#'):
                continue
            if line.startswith('|'):
                out.append((True, line[1:].strip()))
            else:
                out.append((False, line.strip()))
    return out


def build_para(indent_tmpl, flush_tmpl, text, indent):
    tmpl = indent_tmpl if indent else flush_tmpl
    p = copy.deepcopy(tmpl)
    for ch in list(p):
        if ch.tag != qn('pPr'):
            p.remove(ch)
    r_tmpl = tmpl.find(qn('r'))
    r = etree.SubElement(p, qn('r'))
    if r_tmpl is not None:
        rPr = r_tmpl.find(qn('rPr'))
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
    t = etree.SubElement(r, qn('t'))
    t.text = text
    t.set(XML_NS + 'space', 'preserve')
    return p


def find_row(table, find, row_idx):
    rows = table.findall(qn('tr'))
    if find:
        hits = [i for i, tr in enumerate(rows)
                if find in ''.join(text_of(tc) for tc in tr.findall(qn('tc')))]
        if len(hits) != 1:
            raise SystemExit('--find %r 命中 %d 行（必须恰好 1 行）' % (find, len(hits)))
        return hits[0], rows
    if row_idx is None or not (0 <= row_idx < len(rows)):
        raise SystemExit('--row 越界：%r（共 %d 行）' % (row_idx, len(rows)))
    return row_idx, rows


def save_package(src, dst, root):
    xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = dst + '.tmp'
    with zipfile.ZipFile(src) as z:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = xml
                zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                zi.compress_type = item.compress_type
                zi.external_attr = item.external_attr
                zo.writestr(zi, data)
    if os.path.exists(dst):
        os.remove(dst)
    os.replace(tmp, dst)


def main(argv=None):
    ap = argparse.ArgumentParser(description='重建表格单元格的段落序列（格式沿用原模板）')
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--spec', required=True)
    ap.add_argument('--table', type=int, default=0)
    ap.add_argument('--row', type=int, default=None)
    ap.add_argument('--find', default=None)
    ap.add_argument('--cell', type=int, default=0)
    ap.add_argument('--keep-first', type=int, default=1)
    ap.add_argument('--indent-template', type=int, default=None)
    ap.add_argument('--flush-template', type=int, default=None)
    ap.add_argument('--replace', action='append', default=[])
    ap.add_argument('--replace-with', action='append', default=[])
    args = ap.parse_args(argv)

    if len(args.replace) != len(args.replace_with):
        raise SystemExit('--replace 与 --replace-with 必须成对出现')

    spec = load_spec(args.spec)
    if not spec:
        raise SystemExit('spec 为空：%s' % args.spec)

    with zipfile.ZipFile(args.src) as z:
        root = etree.fromstring(z.read('word/document.xml'))

    tables = list(iter_body_tables(root))
    if not (0 <= args.table < len(tables)):
        raise SystemExit('--table 越界（body 内共 %d 张表）' % len(tables))
    table = tables[args.table]
    row_idx, rows = find_row(table, args.find, args.row)
    tcs = rows[row_idx].findall(qn('tc'))
    if not (0 <= args.cell < len(tcs)):
        raise SystemExit('--cell 越界（该行共 %d 个单元格）' % len(tcs))
    cell = tcs[args.cell]
    ps = cell.findall(qn('p'))
    n_old = len(ps)
    if n_old <= args.keep_first:
        raise SystemExit('该单元格只有 %d 段，--keep-first=%d 无法保留' % (n_old, args.keep_first))

    ind_idx = args.keep_first if args.indent_template is None else args.indent_template
    flu_idx = (args.keep_first - 1) if args.flush_template is None else args.flush_template
    for nm, ix in (('--indent-template', ind_idx), ('--flush-template', flu_idx)):
        if not (0 <= ix < n_old):
            raise SystemExit('%s 越界：%d（该单元格共 %d 段）' % (nm, ix, n_old))
    indent_tmpl = copy.deepcopy(ps[ind_idx])
    flush_tmpl = copy.deepcopy(ps[flu_idx])

    anchor = ps[args.keep_first - 1]
    for p in ps[args.keep_first:]:
        cell.remove(p)

    cur = anchor
    n_ind = n_flu = 0
    for indent, text in spec:
        newp = build_para(indent_tmpl, flush_tmpl, text, indent)
        cur.addnext(newp)
        cur = newp
        n_ind += int(indent)
        n_flu += int(not indent)

    # ---- exact text replacements, each must hit exactly once
    n_rep = 0
    for old, new in zip(args.replace, args.replace_with):
        hits = 0
        for t in root.iter(qn('t')):
            if t.text and old in t.text:
                hits += t.text.count(old)
                t.text = t.text.replace(old, new)
        if hits != 1:
            raise SystemExit('替换 %r 命中 %d 次（必须恰好 1 次）' % (old, hits))
        n_rep += 1

    # ---- structural self-check
    nested = [r for r in cell.iter(qn('r')) if r.getparent().tag != qn('p')]
    if nested:
        raise SystemExit('重建后单元格内出现嵌套 w:r（%d 处）' % len(nested))

    save_package(args.src, args.dst, root)
    print('target: table%d row%d cell%d  paragraphs %d -> %d  (顶格 %d / 缩进 %d, 替换 %d 处)'
          % (args.table, row_idx, args.cell, n_old, len(cell.findall(qn('p'))), n_flu, n_ind, n_rep))
    print('wrote %s (%d bytes)' % (args.dst, os.path.getsize(args.dst)))


if __name__ == '__main__':
    main()
