# -*- coding: utf-8 -*-
"""交付前自查：章节字数均衡 / 中英文标点混用 / 上标引用编号 / 参考文献编号。

这几项都是"批注里不会写、但导师一定会看到"的硬伤：
  * 导师说"注意各章节字数均衡"，靠肉眼估不出来 —— 本脚本按章统计并给最大/最小比；
  * 脚本里手写的中文段落很容易混进 ASCII 直引号 `"`（SKILL 陷阱十九）；
  * 上标编号必须"首次出现顺序 = 1..N"（顺序编码制）且覆盖 1..N；
  * 文后参考文献必须是 `[1]..[N]` 连续。

用法::

    python check_balance.py 论文.docx
    python check_balance.py 论文.docx --chapter-style 2 --tolerance 0.30
    python check_balance.py 论文.docx --start 摘要 --json

参数（缺失即自动推断，不写死任何学生/章节/样式号）：

* ``--chapter-style``：章标题的样式号；省略时按 styles.xml 的样式名自动识别
  （`heading 1` / `标题 1`），再退化为"带 outlineLvl=0 的段落"。
* ``--start``：从文本等于/包含该串的段落开始统计（默认从第一个章标题开始），
  用来跳过封面与目录。
* ``--tolerance``：允许的 (最大章字数 / 最小章字数 - 1)，默认 0.30。

退出码：0 = 无问题；1 = 存在问题（字数失衡、直引号、编号不连续等）。
"""
import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docxlib as D  # noqa: E402

ASCII_QUOTES = ('"', "'")
# 半角标点后紧跟汉字，是"中英文标点混用"的典型形态（英文缩写、数字后接汉字不算）
HALF_PUNCT = re.compile(r'([,;:!?])(?=[\u4e00-\u9fff])')
REF_RE = re.compile(r'^\[(\d+)\]')


def pick_chapter_style(z, root, override):
    if override:
        return override
    names = D.read_styles(z)
    norm = {}
    for sid, nm in names.items():
        norm[sid] = (nm or '').replace(' ', '').lower()
    for sid, nm in norm.items():
        if nm in ('heading1', '标题1'):
            return sid
    # 退化：带 outlineLvl=0 的段落样式
    for p in root.iter(D.qn('p')):
        pPr = p.find(D.qn('pPr'))
        if pPr is None:
            continue
        ol = pPr.find(D.qn('outlineLvl'))
        if ol is not None and ol.get(D.qn('val')) == '0':
            sid = D.para_style(p)
            if sid:
                return sid
    return None


def outline0(root):
    out = []
    for p in root.iter(D.qn('p')):
        pPr = p.find(D.qn('pPr'))
        if pPr is None:
            continue
        ol = pPr.find(D.qn('outlineLvl'))
        if ol is not None and ol.get(D.qn('val')) == '0':
            out.append(p)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--chapter-style', default=None)
    ap.add_argument('--start', default=None, help='从包含该文本的段落之后开始统计')
    ap.add_argument('--tolerance', type=float, default=0.30)
    ap.add_argument('--strict-balance', action='store_true',
                    help='把"章节字数失衡"也当失败项；默认只作提示（不同模板的章标题'
                         '样式常混入声明页/封面等伪标题，硬判会误报，见 SKILL 陷阱十三）')
    ap.add_argument('--exclude', default='摘要,Abstract,目录,参考文献,致谢,结语,附录,声明,授权,承诺,封面',
                    help='不计入"章节字数均衡"比对的章标题关键字（逗号分隔）；'
                         '前置/后置部分与结语本就不参与正文均衡')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()
    excludes = [x.strip() for x in args.exclude.split(',') if x.strip()]

    z, root, _ = D.open_docx(args.docx)
    body = root.find(D.qn('body'))
    paras = [el for el in body if el.tag == D.qn('p')]
    texts = [D.get_text(p) for p in paras]

    sid = pick_chapter_style(z, root, args.chapter_style)
    chapters = [p for p in paras if sid is not None and D.para_style(p) == sid]
    if not chapters:
        chapters = outline0(root)

    problems, notes = [], []

    # --- 1. 章节字数 ---
    idx = [paras.index(p) for p in chapters]
    if args.start:
        hit = [i for i, t in enumerate(texts) if args.start in t]
        if hit:
            idx = [i for i in idx if i >= hit[-1]]
    counts = []
    for k, i in enumerate(idx):
        end = idx[k + 1] if k + 1 < len(idx) else len(paras)
        head = texts[i].strip()
        n = sum(len(texts[j].strip()) for j in range(i + 1, end))
        counts.append((head, n))
    if len(counts) >= 2:
        def _excluded(title):
            # 标题里常有对齐用的空格（"摘   要""致  谢"），比对前先去掉空白
            t = re.sub(r'\s+', '', title)
            return any(x in t for x in excludes)

        vals = [(t, n) for t, n in counts if n > 0 and not _excluded(t)]
        if vals:
            ratio = max(n for _, n in vals) / float(min(n for _, n in vals)) - 1.0
            if ratio > args.tolerance:
                big = max(vals, key=lambda x: x[1])
                small = min(vals, key=lambda x: x[1])
                msg = ('章节字数失衡: 最大「%s」%d 字 / 最小「%s」%d 字，相差 %.0f%%（上限 %.0f%%）'
                       % (big[0][:16], big[1], small[0][:16], small[1],
                          ratio * 100, args.tolerance * 100))
                (problems if args.strict_balance else notes).append(
                    msg if args.strict_balance else '[提示] ' + msg)
            else:
                notes.append('正文各章字数均衡: 最大/最小相差 %.0f%%（上限 %.0f%%）'
                             % (ratio * 100, args.tolerance * 100))

    # --- 2. ASCII 直引号 / 半角标点 ---
    joined = '\n'.join(texts)
    ascii_q = sum(joined.count(c) for c in ASCII_QUOTES)
    if ascii_q:
        problems.append('正文存在 %d 个 ASCII 直引号（应使用中文弯引号 “ ”）' % ascii_q)
    half = HALF_PUNCT.findall(joined)
    if half:
        problems.append('存在 %d 处半角标点后紧跟汉字（应改为全角 ，；：！？）' % len(half))
    if not ascii_q and not half:
        notes.append('中英文标点规范: 无 ASCII 直引号、无半角标点后接汉字')

    # --- 3. 上标引用编号 ---
    nums, empty = D.superscript_numbers(root)
    order, seen = [], set()
    for n in nums:
        if n not in seen:
            seen.add(n)
            order.append(n)
    if empty:
        problems.append('存在 %d 个空文本上标' % empty)
    if nums:
        if order != list(range(1, len(order) + 1)):
            problems.append('上标编号首次出现顺序不是 1..N 递增: %s' % order)
        if sorted(set(nums)) != list(range(1, len(set(nums)) + 1)):
            problems.append('上标编号不连续: %s' % sorted(set(nums)))
        else:
            notes.append('上标引用 %d 处，编号首次出现顺序 1..%d 且连续'
                         % (len(nums), len(set(nums))))

    # --- 4. 参考文献编号 ---
    refs = [t.strip() for t in texts if REF_RE.match(t.strip())]
    ref_nums = [int(REF_RE.match(t).group(1)) for t in refs]
    if ref_nums:
        if ref_nums != list(range(1, len(ref_nums) + 1)):
            problems.append('参考文献编号不连续: %s' % ref_nums)
        else:
            notes.append('参考文献 %d 条，编号 [1]..[%d] 连续' % (len(ref_nums), len(ref_nums)))
        # 正文引用编号集合 vs 文献表条数
        if nums and sorted(set(nums)) != list(range(1, len(ref_nums) + 1)):
            problems.append('正文引用编号 %s 与文献表条数 %d 不匹配'
                            % (sorted(set(nums)), len(ref_nums)))

    out = {
        'docx': os.path.basename(args.docx),
        'chapterStyle': sid,
        'chapters': [{'title': t, 'chars': n} for t, n in counts],
        'superscripts': nums,
        'references': len(ref_nums),
        'problems': problems,
        'notes': notes,
    }
    if args.json:
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print('== 交付前自查:', out['docx'], '==')
        print('章标题样式号:', sid)
        if counts:
            print('-- 各章字数 --')
            for t, n in counts:
                print('   %-24s %5d 字' % (t[:24], n))
        for n_ in notes:
            print('[OK] ' + n_)
        for p_ in problems:
            print('[FAIL] ' + p_)
        print('结论:', '存在问题 %d 项' % len(problems) if problems else '全部通过')
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main())
