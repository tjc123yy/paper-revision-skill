# -*- coding: utf-8 -*-
"""脚注落点与编号核对（配置驱动，无硬编码）。

检查四件事（对应 SKILL.md 陷阱四十五/四十六）：

1. **批注锚点 vs 脚注落点**（给 `--comments` 指向改前带批注的 docx 时）：
   把每条批注 `commentRangeStart..commentRangeEnd` 之间的原文打印出来，
   并列出成稿里每个脚注落在哪个段落、段内哪些字符之间 —— 一眼看出"插错段落"。
2. **脚注 id 连续性**：`w:footnoteReference` 的 id 必须恰好是 `0..N-1`（不连续即隐患）。
3. **定义顺序 == 文档引用顺序**：`footnotes.xml` 里内容脚注的出现顺序必须与首次引用顺序一致；
   否则渲染器（LibreOffice/WPS）会把脚注整片挪到文末。
4. **separator 脚注存在**：`w:type="separator"` 与 `w:type="continuationSeparator"` 必须还在。

用法：
    python check_footnotes.py 论文_已修改.docx --comments 论文_备份.docx
    python check_footnotes.py 论文_已修改.docx --json

退出码：0 = 无问题；1 = 有问题。
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import docxlib as D  # noqa: E402
from lxml import etree  # noqa: E402


def para_list(root):
    body = root.find(D.qn('body'))
    return [p for p in body if p.tag == D.qn('p')] if body is not None else []


def flat_with_fn(p):
    """按 run 顺序把段落展平成 (类型, 值) 序列，脚注标记为 ('FN', id)。"""
    seq = []
    for r in p.findall(D.qn('r')):
        fr = r.find(D.qn('footnoteReference'))
        if fr is not None:
            seq.append(('FN', fr.get(D.qn('id'))))
        else:
            for t in r.findall(D.qn('t')):
                for ch in (t.text or ''):
                    seq.append(('t', ch))
    return seq


def comment_anchors(z):
    """返回 [(批注id, 锚点原文)]。"""
    root = etree.fromstring(z.read('word/document.xml'))
    out = []
    for cs in root.iter(D.qn('commentRangeStart')):
        cid = cs.get(D.qn('id'))
        buf, node = [], cs
        while True:
            node = node.getnext()
            if node is None:
                break
            if etree.QName(node).localname == 'commentRangeEnd' and \
                    node.get(D.qn('id')) == cid:
                break
            buf.append(''.join(t.text or '' for t in node.iter(D.qn('t'))))
        out.append((cid, ''.join(buf)))
    return out


def main():
    ap = argparse.ArgumentParser(description='脚注落点与编号核对')
    ap.add_argument('docx')
    ap.add_argument('--comments', default=None,
                    help='改前带批注的 docx：打印每条批注锚点，便于逐条核对')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()

    z, root, _ = D.open_docx(args.docx)
    problems = []

    # ---- 1. 引用顺序与落点 ----
    paras = para_list(root)
    ref_order, placement = [], []
    for i, p in enumerate(paras):
        seq = flat_with_fn(p)
        fns = [(j, v) for j, (k, v) in enumerate(seq) if k == 'FN']
        for pos, fid in fns:
            if fid not in ref_order:
                ref_order.append(fid)
            prev_ch = seq[pos - 1][1] if pos > 0 else ''
            next_ch = seq[pos + 1][1] if pos + 1 < len(seq) else ''
            ctx = ''.join(v for k, v in seq[max(0, pos - 16):pos + 3])
            placement.append({'footnoteId': fid, 'paraIndex': i, 'charPos': pos,
                              'prevChar': prev_ch, 'nextChar': next_ch,
                              'context': ctx,
                              'paragraphHead': D.get_text(p)[:24]})

    # ---- 2. id 连续性 ----
    want = [str(i) for i in range(len(ref_order))]
    if sorted(ref_order, key=lambda x: int(x)) != want:
        problems.append('脚注 id 不连续：引用顺序 %s（应为 0..%d）'
                        % (ref_order, len(ref_order) - 1))

    # ---- 3. 定义顺序 == 引用顺序 ----
    fnroot = etree.fromstring(z.read('word/footnotes.xml'))
    defs = [f.get(D.qn('id')) for f in fnroot.iter(D.qn('footnote'))
            if f.get(D.qn('type')) is None]
    if defs != ref_order:
        problems.append('footnotes.xml 定义顺序 != 文档引用顺序：定义 %s / 引用 %s'
                        % (defs, ref_order))

    # ---- 4. separator ----
    types = [f.get(D.qn('type')) for f in fnroot.iter(D.qn('footnote'))
             if f.get(D.qn('type'))]
    for need in ('separator', 'continuationSeparator'):
        if need not in types:
            problems.append('缺少特殊脚注 w:type="%s"（Word 会判文件需修复）' % need)

    # ---- 批注锚点对照 ----
    anchors = []
    if args.comments:
        zc = __import__('zipfile').ZipFile(args.comments)
        anchors = comment_anchors(zc)

    result = {
        'docx': args.docx,
        'footnoteCount': len(ref_order),
        'referenceOrder': ref_order,
        'definitionOrder': defs,
        'placement': placement,
        'commentAnchors': [{'id': c, 'text': t} for c, t in anchors],
        'problems': problems,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print('== 脚注核对:', os.path.basename(args.docx), '==')
        print('脚注数 %d | 引用顺序 %s' % (len(ref_order), ref_order))
        print('定义顺序 %s' % defs)
        print('特殊脚注 %s' % types)
        if anchors:
            print('')
            print('--- 批注锚点（逐条对着看脚注是否落在同一句）---')
            for cid, t in anchors:
                print('  [批注 %s] %s' % (cid, t[:70]))
        print('')
        print('--- 脚注落点 ---')
        for pl in placement:
            print('  id%-3s 段%-4d 位置%-4d 前=%r 后=%r | …%s'
                  % (pl['footnoteId'], pl['paraIndex'], pl['charPos'],
                     pl['prevChar'], pl['nextChar'], pl['context']))
        print('')
        if problems:
            for p in problems:
                print('[FAIL] ' + p)
        else:
            print('[PASS] id 连续、定义顺序与引用顺序一致、特殊脚注齐全')
    sys.exit(1 if problems else 0)


if __name__ == '__main__':
    main()
