# -*- coding: utf-8 -*-
"""按最终 PDF 的真实分页，回写目录（TOC 域）里的**缓存页码**。

问题（见 SKILL.md 陷阱十四 / 四十二）：目录每个条目的页码是 `PAGEREF` 域的缓存结果，
就是一个普通 `<w:t>`。正文增删后页码会整体偏移，而
`<w:updateFields w:val="true"/>` 只让 Word **打开时**重算；
`libreoffice-kit convert` 出的 PDF 里仍然是旧页码 —— 于是"渲染看着正常"会骗过常规验收。

本脚本：转 PDF → 量真实分页 → 回写缓存页码；配合 `--check` 做**不动点校验**。

用法::

    python fix_toc_pages.py 论文.docx out.pdf --body-first-page 5 [--front-first-page 4]
    python fix_toc_pages.py 论文.docx out.pdf --body-first-page 5 --check   # 只校验不写

参数含义（都必须按渲染出的**页脚**确认，不要凭封面/摘要页数推算）：

* ``--body-first-page``：PDF 里"正文第 1 页"是第几页（本次为 5）。
* ``--front-first-page``：PDF 里罗马数字页码" I "是第几页；不给则视为没有罗马数字前缀页。

退出码：正常回写后若有条目未定位到 → 1；``--check`` 时 0 表示全部一致。
"""
import argparse
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docxlib as D                      # noqa: E402
from lxml import etree                   # noqa: E402

# 目录段落的样式号（toc 1 / toc 2 / toc 3）。不同模板可能不同，必要时用 --toc-styles 覆盖。
DEFAULT_TOC_STYLES = {'20', '22', '15'}

# 正文标题的"一、""（一）""1．"等编号由 numPr 自动生成，**不落在 XML 文字里**（陷阱十四），
# 而 PDF 里渲染出来了。所以比对前要把行首编号剥掉再做**等值**比较；
# 不能用"包含即可"——那会让"图3.17 蓝牙模块实物图"命中标题"蓝牙模块"。
NUM_PREFIX = [
    r'^[一二三四五六七八九十]+、',
    r'^（[一二三四五六七八九十]+）',
    r'^\d+[．.、]',
    r'^（\d+）',
    r'^[\(\（]\d+[\)\）]',
]

ROMAN = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
         'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX']


def norm(s):
    return re.sub(r'\s+', '', s or '')


def strip_num(s):
    prev = None
    while prev != s:
        prev = s
        for pat in NUM_PREFIX:
            s = re.sub(pat, '', s)
    return s


def pdf_lines(path):
    import pypdf
    r = pypdf.PdfReader(path)
    out = []
    for i, p in enumerate(r.pages):
        for line in (p.extract_text() or '').splitlines():
            out.append((i + 1, norm(line)))
    return out, len(r.pages)


def build_entries(root, toc_styles):
    """返回 [(标题文字, 当前缓存页码, 书签名)]，按文档顺序。"""
    body = root.find(D.qn('body'))
    paras = [c for c in body if c.tag == D.qn('p')]
    bm2text = {}
    for p in paras:
        t = D.get_text(p)
        for bs in p.iter(D.qn('bookmarkStart')):
            nm = bs.get(D.qn('name'))
            if nm:
                bm2text[nm] = t

    entries = []
    for p in paras:
        if D.para_style(p) not in toc_styles:
            continue
        name = None
        st = None
        for r in p.iter(D.qn('r')):
            ins = r.find(D.qn('instrText'))
            fld = r.find(D.qn('fldChar'))
            if ins is not None and 'PAGEREF' in (ins.text or ''):
                m = re.search(r'PAGEREF\s+(\S+)', ins.text)
                name = m.group(1) if m else None
                st = 'ref'
                continue
            if fld is not None:
                ft = fld.get(D.qn('fldCharType'))
                if ft == 'separate' and st == 'ref':
                    st = 'num'
                elif ft == 'end' and st == 'num':
                    st = None
                continue
            if st == 'num':
                ts = r.findall(D.qn('t'))
                if ts:
                    entries.append([bm2text.get(name, ''), ts[0], name])
                    st = None
    return entries


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('pdf')
    ap.add_argument('--body-first-page', type=int, required=True,
                    help='PDF 中"正文第 1 页"的页序号（看页脚确认）')
    ap.add_argument('--front-first-page', type=int, default=None,
                    help='PDF 中罗马数字"I"的页序号；省略表示无前缀页')
    ap.add_argument('--toc-styles', default=','.join(sorted(DEFAULT_TOC_STYLES)))
    ap.add_argument('--check', action='store_true', help='只校验不写回')
    args = ap.parse_args()

    toc_styles = set(x.strip() for x in args.toc_styles.split(',') if x.strip())
    lines, npages = pdf_lines(args.pdf)
    print('PDF 页数: %d' % npages)

    with zipfile.ZipFile(args.docx) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    entries = build_entries(root, toc_styles)
    print('目录条目: %d' % len(entries))

    start = (args.front_first_page or args.body_first_page)

    def find_page(title):
        key = norm(title)
        if not key:
            return None
        for pno, l in lines:                                  # 完全相等
            if pno >= start and l == key:
                return pno
        for pno, l in lines:                                  # 剥掉行首自动编号后相等
            if pno >= start and strip_num(l) == key:
                return pno
        return None

    def printed(pno):
        if args.front_first_page and pno < args.body_first_page:
            idx = pno - args.front_first_page + 1
            return ROMAN[idx] if 0 < idx < len(ROMAN) else str(idx)
        return str(pno - args.body_first_page + 1)

    updated, skipped = 0, []
    for title, tnode, _bm in entries:
        pno = find_page(title)
        if pno is None:
            skipped.append(norm(title)[:24])
            continue
        want = printed(pno)
        if (tnode.text or '') != want:
            if args.check:
                print('  不一致: %-28s 缓存=%s 实测=%s' % (norm(title)[:26], tnode.text, want))
            else:
                tnode.text = want
            updated += 1

    if skipped:
        print('未定位到页码（保持原值）: %s' % skipped)
    print('待更新/不一致条目: %d' % updated)

    if args.check:
        if updated == 0 and not skipped:
            print('[不动点校验] 全部条目页码已与实测分页一致 —— PASS')
            return 0
        print('[不动点校验] 仍有 %d 条不一致 —— FAIL' % (updated + len(skipped)))
        return 1

    D.save_root(args.docx, args.docx, root)
    print('已写回:', args.docx)
    return 0 if not skipped else 1


if __name__ == '__main__':
    sys.exit(main())
