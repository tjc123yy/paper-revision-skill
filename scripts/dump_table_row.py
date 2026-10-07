# -*- coding: utf-8 -*-
"""打印 docx 正文首表的"行 → 单元格 → 段落 → run"结构，用于定位要改的小节。

开题报告 / 任务书这类文档的全部内容都装在一张大表格里，
按段落序号定位会错位，必须按行定位。本脚本给出每行的行号、
每个单元格的段落文本，以及每段的 `pPr` 与每个 run 的 `rFonts / sz / b / u`，
可以直接看出：

* 目标小节在第几行、单元格内第几段（配合 `rebuild_cell.py --find/--keep-first`）；
* 该单元格里"缩进段"与"顶格段"的模板分别是第几段（`<w:ind firstLineChars>` 有无）；
* 段落里有没有 `w:br` / `w:noBreakHyphen` / `w:softHyphen` 这类**不含文字**的元素
  （按 w:t 拼文本时它们会"消失"，容易误判成缺分隔符）。

用法：
    python dump_table_row.py <docx> [--table 0] [--rows 5,6] [--find 研究的主要内容]
                             [--max-text 60] [--runs]
"""
import argparse
import zipfile

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def qn(tag):
    return W + tag


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(qn('t')))


def ppr_summary(p):
    pPr = p.find(qn('pPr'))
    if pPr is None:
        return '-'
    bits = []
    ps = pPr.find(qn('pStyle'))
    if ps is not None:
        bits.append('style=%s' % (ps.get(qn('val')) or ''))
    sp = pPr.find(qn('spacing'))
    if sp is not None:
        bits.append('spacing(line=%s,rule=%s,before=%s,after=%s)' % (
            sp.get(qn('line')), sp.get(qn('lineRule')),
            sp.get(qn('before')), sp.get(qn('after'))))
    ind = pPr.find(qn('ind'))
    if ind is not None:
        bits.append('ind(firstLineChars=%s,firstLine=%s)' % (
            ind.get(qn('firstLineChars')), ind.get(qn('firstLine'))))
    return ' '.join(bits) if bits else '-'


def run_summary(r):
    rPr = r.find(qn('rPr'))
    d = {}
    if rPr is not None:
        rf = rPr.find(qn('rFonts'))
        if rf is not None:
            d['ascii'] = rf.get(qn('ascii'))
            d['ea'] = rf.get(qn('eastAsia'))
        sz = rPr.find(qn('sz'))
        if sz is not None:
            d['sz'] = sz.get(qn('val'))
        for k in ('b', 'u', 'i'):
            if rPr.find(qn(k)) is not None:
                d[k] = '1'
    extras = []
    for tag in ('br', 'noBreakHyphen', 'softHyphen', 'drawing', 'pict', 'object',
                'bookmarkStart', 'bookmarkEnd', 'fldChar', 'instrText', 'tab'):
        n = len(list(r.iter(qn(tag))))
        if n:
            extras.append('%s=%d' % (tag, n))
    return '%s%s %r' % (d, (' ' + ' '.join(extras)) if extras else '',
                        ''.join(t.text or '' for t in r.iter(qn('t'))))


def main():
    ap = argparse.ArgumentParser(description='打印正文首表的行/单元格/段落结构')
    ap.add_argument('docx')
    ap.add_argument('--table', type=int, default=0)
    ap.add_argument('--rows', default=None, help='只看这些行，如 5,6')
    ap.add_argument('--find', default=None, help='只打印包含该文字的行')
    ap.add_argument('--max-text', type=int, default=60)
    ap.add_argument('--runs', action='store_true', help='打印每个 run 的字体/字号')
    args = ap.parse_args()

    with zipfile.ZipFile(args.docx) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    body = root.find(qn('body'))
    tables = [el for el in body if el.tag == qn('tbl')]
    if not (0 <= args.table < len(tables)):
        raise SystemExit('--table 越界（body 内共 %d 张表）' % len(tables))
    tbl = tables[args.table]
    want = set(int(x) for x in args.rows.split(',')) if args.rows else None
    print('body 内表格 %d 张；本表 %d 行' % (len(tables), len(tbl.findall(qn('tr')))))
    for ri, tr in enumerate(tbl.findall(qn('tr'))):
        tcs = tr.findall(qn('tc'))
        rowtext = ''.join(text_of(tc) for tc in tcs)
        if want is not None and ri not in want:
            continue
        if args.find and args.find not in rowtext:
            continue
        print('=== row %d  (%d cells)  %s' % (ri, len(tcs), rowtext[:args.max_text]))
        for ci, tc in enumerate(tcs):
            for pi, p in enumerate(tc.findall(qn('p'))):
                txt = text_of(p)
                print('   c%d p%-2d %-58s | %s'
                      % (ci, pi, ppr_summary(p), txt[:args.max_text]))
                if args.runs:
                    for r in p.findall(qn('r')):
                        print('        run %s' % run_summary(r))


if __name__ == '__main__':
    main()
