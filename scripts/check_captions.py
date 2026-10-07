# -*- coding: utf-8 -*-
"""图表编号与引用双向核对（参数化，无硬编码样式号/章节号）

检测四类问题：
    1. 题注编号重复（如正文里出现两个"图5-10"）
    2. 有题注但正文从未引用
    3. 正文引用了但找不到对应题注（"如表6-5所示"却没有"表6-5 ..."题注段）
    4. 图片数量与图题注数量不匹配（有图无题注 / 题注多于图）

用法：
    python check_captions.py <docx> [--json] [--allow-uncaptioned N]

--allow-uncaptioned：允许 N 张"本来就不需要题注"的图（封面校徽、页眉装饰、
    签名占位图等）。这类图不在正文里，格式规范也不要求题注，属于合理豁免；
    默认 0，即要求每一张 w:drawing 都有图题注。

退出码：0 = 无问题；1 = 存在问题。
"""
import argparse
import collections
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import docxlib as D  # noqa: E402

# 题注：以 图/表 + 编号 开头（编号形如 5-10 / 5.10 / 5_10）
CAP_RE = re.compile(r'^([图表])\s*([0-9]+)\s*[-.－_]\s*([0-9]+)')
# 正文引用：图/表 + 编号，允许"如图5-10所示""表 6-5"
REF_RE = re.compile(r'([图表])\s*([0-9]+)\s*[-.－_]\s*([0-9]+)')


def normalize_prefix(text, kind):
    """'（一）' / '一、' / '1．' 等前缀不参与题注判定，去掉后再看是否以 图N-M 开头。"""
    return text.strip()


def body_children(root):
    body = root.find(D.qn('body'))
    return list(body) if body is not None else []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--allow-uncaptioned', type=int, default=0,
                    help='允许多少张无需题注的图（封面/页眉/签名占位图）')
    args = ap.parse_args()

    _, root, _ = D.open_docx(args.docx)

    fig_caps, tbl_caps = [], []
    fig_refs, tbl_refs = collections.Counter(), collections.Counter()
    drawings = 0
    for el in body_children(root):
        if el.tag == D.qn('tbl'):
            drawings += len(el.findall('.//' + D.qn('drawing')))
            continue
        if el.tag != D.qn('p'):
            continue
        drawings += len(el.findall('.//' + D.qn('drawing')))
        txt = D.get_text(el).strip()
        if not txt:
            continue
        m = CAP_RE.match(txt)
        if m:
            key = '%s-%s' % (m.group(2), m.group(3))
            (fig_caps if m.group(1) == '图' else tbl_caps).append((key, txt))
            continue
        for r in REF_RE.finditer(txt):
            key = '%s-%s' % (r.group(2), r.group(3))
            (fig_refs if r.group(1) == '图' else tbl_refs)[key] += 1

    fig_ids = [k for k, _ in fig_caps]
    tbl_ids = [k for k, _ in tbl_caps]

    problems = []
    dup_fig = [k for k, v in collections.Counter(fig_ids).items() if v > 1]
    dup_tbl = [k for k, v in collections.Counter(tbl_ids).items() if v > 1]
    if dup_fig:
        problems.append('图题注编号重复: ' + ','.join(sorted(dup_fig)))
    if dup_tbl:
        problems.append('表题注编号重复: ' + ','.join(sorted(dup_tbl)))

    # 图号与表号是两套独立编号，必须分开比对（否则"图6-1"会把"表6-1"的缺失掩盖掉）
    dangling_fig = sorted(k for k in fig_refs if k not in set(fig_ids))
    dangling_tbl = sorted(k for k in tbl_refs if k not in set(tbl_ids))
    if dangling_fig:
        problems.append('正文引用了但无图题注: ' + ','.join(dangling_fig))
    if dangling_tbl:
        problems.append('正文引用了但无表题注: ' + ','.join(dangling_tbl))

    unused_fig = [k for k in fig_ids if fig_refs.get(k, 0) == 0]
    unused_tbl = [k for k in tbl_ids if tbl_refs.get(k, 0) == 0]
    if unused_fig:
        problems.append('图有题注但正文未引用: ' + ','.join(unused_fig))
    if unused_tbl:
        problems.append('表有题注但正文未引用: ' + ','.join(unused_tbl))

    if drawings - args.allow_uncaptioned != len(fig_caps):
        problems.append('图片数与图题注数不等: 图片 %d（其中 %d 张豁免）/ 图题注 %d'
                        % (drawings, args.allow_uncaptioned, len(fig_caps)))

    out = {
        'docx': args.docx,
        'drawings': drawings,
        'allowedUncaptioned': args.allow_uncaptioned,
        'figureCaptions': len(fig_caps),
        'tableCaptions': len(tbl_caps),
        'problems': problems,
    }
    if args.json:
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print('== 图表编号核对:', os.path.basename(args.docx), '==')
        print('图片(w:drawing) %d（豁免 %d）| 图题注 %d | 表题注 %d'
              % (drawings, args.allow_uncaptioned, len(fig_caps), len(tbl_caps)))
        if problems:
            for p in problems:
                print('[FAIL] ' + p)
        else:
            print('[PASS] 编号唯一、题注与引用双向齐全、图片数与图题注数一致')
    sys.exit(1 if problems else 0)


if __name__ == '__main__':
    main()
