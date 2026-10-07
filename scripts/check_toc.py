# -*- coding: utf-8 -*-
"""目录与正文标题双向核对（参数化，样式号自动识别，无硬编码）

回答"目录里改了、正文没改"这类意见：把目录条目与正文标题各自剥掉
"一、/（一）/1．"等序号前缀后做**双向差集**，两边都为空才算一致。

目录样式与标题样式默认按 styles.xml 里的**样式名**识别
（目录：名字以 toc/目录 开头；标题：名字以 heading/标题 开头），
可用 --toc-styles / --heading-styles 覆盖为样式号列表。

用法：
    python check_toc.py <docx> [--toc-styles 17,18,14] [--heading-styles 3,4,5] [--json]

退出码：0 = 双向一致；1 = 存在只在一边出现的条目。
"""
import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import docxlib as D  # noqa: E402

# 目录条目结尾的页码（阿拉伯数字或罗马数字）
PAGENUM_RE = re.compile(r'[0-9IVXivx]+\s*$')
# 中文序号前缀：一、 / （一） / (1) / 1． / 1. / 1、
ORDINAL_RE = re.compile(
    r'^(?:[一二三四五六七八九十]+\s*[、.．]'
    r'|[(（]\s*[一二三四五六七八九十0-9]+\s*[)）]'
    r'|[0-9]+\s*[、.．])\s*')


def norm(text):
    """归一到"可比较的标题名"。

    两边的形态并不一样，必须都归一：
      * 目录条目里带 Word 生成的自动编号（"1.1 课题研究背景"）与**已缓存**的旧页码；
      * 正文标题的编号由 numPr 自动生成，**不落在 XML 文字里**（只有"课题研究背景"）。
    所以统一去掉所有数字串（编号、页码）与空白，只留标题文字本身。
    """
    t = text.strip()
    t = PAGENUM_RE.sub('', t).strip()
    prev = None
    while prev != t:
        prev = t
        t = ORDINAL_RE.sub('', t).strip()
    t = re.sub(r'\d+', '', t)
    t = re.sub(r'\s+', '', t)
    return t


def collect(root, style_map, toc_styles, heading_styles):
    """分别收集目录条目文本与正文标题文本。

    目录条目里同一个段落可能混着两个域：HYPERLINK 的显示文字（条目名）和
    PAGEREF 的显示文字（**已缓存的旧页码**）。直接 get_text() 会把旧页码粘在
    条目名后面（"1.1 课题研究背景1"），于是目录与正文永远对不上。
    这里用 w:fldChar/w:instrText 的状态机剥掉 PAGEREF 域的显示文字，
    只留下条目名。

    返回 (toc_entries, heading_texts, toc_pagenums)。
    """
    body = root.find(D.qn('body'))
    toc, heads, pagenums = [], [], []
    for p in body.iter(D.qn('p')):
        sid = D.para_style(p)
        if not sid:
            continue
        txt_all = D.get_text(p).strip()
        if not txt_all:
            continue
        if sid in toc_styles:
            name, pages = split_toc_entry(p)
            if name:
                toc.append(name)
            pagenums.append(pages)
        elif sid in heading_styles:
            heads.append(txt_all)
    return toc, heads, pagenums


def split_toc_entry(p):
    """把目录条目段落拆成 (条目名, 页码缓存)。

    目录条目一般长这样：
        fldChar(begin) instrText(HYPERLINK \\l _TocX) fldChar(separate)
        <条目名，含自动编号，如 "1 " + "绪言">
        tab
        fldChar(begin) instrText(PAGEREF _TocX) fldChar(separate) <页码> fldChar(end)
        fldChar(end)

    所以：HYPERLINK 域里 separate 之后的文字才是条目名；PAGEREF 域里的文字是
    **已缓存的旧页码**，必须丢掉，否则"1 绪言"会被读成"1 绪言13"。
    """
    name_parts, page_parts = [], []
    field = None          # None / 'HYPERLINK' / 'PAGEREF' / 'other'
    active = False        # 是否处于某个域的"结果"部分
    # 注意：WPS 生成的目录把条目文字放在 <w:hyperlink> 里，run 不是 <w:p> 的直接子元素，
    # 所以必须用 iter() 递归遍历，否则条目名会被读成空串（目录条目数变 0）。
    for child in p.iter(D.qn('r')):
        tag = child.tag.split('}')[1]
        if tag != 'r':
            continue
        for sub in child:
            st = sub.tag.split('}')[1]
            if st == 'fldChar':
                t = sub.get(D.qn('fldCharType'))
                if t == 'begin':
                    field, active = 'unknown', False
                elif t == 'separate':
                    active = True
                elif t == 'end':
                    field, active = None, False
            elif st == 'instrText':
                code = (sub.text or '').strip()
                if 'PAGEREF' in code:
                    field = 'PAGEREF'
                elif 'HYPERLINK' in code:
                    field = 'HYPERLINK'
            elif st == 't':
                if field == 'PAGEREF' and active:
                    page_parts.append(sub.text or '')
                elif field in (None, 'HYPERLINK', 'unknown'):
                    name_parts.append(sub.text or '')
    return ''.join(name_parts).strip(), ''.join(page_parts).strip()


def resolve_styles(style_map, given, kind):
    """给定样式号就用给定的；否则按样式名前缀自动识别。

    自动识别要排掉两类伪标题：
      * 字符样式（"标题 2 Char"），它们不是段落样式，不该进标题集合；
      * 版式样式名里带"字符"或"不入目录"的。
    """
    if given:
        return set(given)
    out = set()
    for sid, name in style_map.items():
        low = (name or '').strip().lower()
        if ('char' in low or '字符' in low or '不入目录' in low):
            continue
        if kind == 'toc' and (low.startswith('toc') or low.startswith('目录')):
            out.add(sid)
        if kind == 'heading' and (low.startswith('heading') or low.startswith('标题')):
            out.add(sid)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--toc-styles', default=None)
    ap.add_argument('--heading-styles', default=None)
    ap.add_argument('--toc-title', default='目录',
                    help='目录页自己的标题（它当然不在自己的目录里，默认"目录"）')
    ap.add_argument('--body-start', default=None,
                    help='正文起始标题文字；它之前的标题（封面题目等）不计入比对')
    ap.add_argument('--expect-missing', default=None,
                    help='预期"目录里还没有"的标题（逗号分隔）。目录是 Word 域，'
                         '缓存不会自动更新；本次新增的小节必然暂缺，需在打开 Word '
                         '更新域后才会出现。这些条目一旦出现在目录里就会报"多出"，'
                         '从而保证检查仍然有区分度。')
    ap.add_argument('--toc-only-ok', default=None,
                    help='允许"只在目录、正文没有"的条目（逗号分隔），如 '
                         '摘要,Abstract,目录,参考文献,致谢；这些标题段落不带 '
                         'heading 样式，本就不会进正文标题集合。')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()

    z, root, _ = D.open_docx(args.docx)
    style_map = D.read_styles(z)

    toc_styles = resolve_styles(
        style_map, args.toc_styles.split(',') if args.toc_styles else None, 'toc')
    heading_styles = resolve_styles(
        style_map, args.heading_styles.split(',') if args.heading_styles else None, 'heading')

    toc, heads, toc_pages = collect(root, style_map, toc_styles, heading_styles)

    # 正文起始标题之前的标题（封面题目）不参与比对
    if args.body_start:
        marker = norm(args.body_start)
        for i, h in enumerate(heads):
            if norm(h) == marker:
                heads = heads[i:]
                break

    def key(x):
        """比较用的键：去编号、去页码、去所有空白。"""
        return re.sub(r'[\s\u3000]+', '', norm(x))

    head_map = {}
    for h in heads:
        head_map.setdefault(key(h), h)
    toc_map = {}
    for t in toc:
        toc_map.setdefault(key(t), t)

    # 目录条目的名字里含自动编号（"1 绪言"），正文没有；
    # 用"其一包含另一"兜住这种差异，同时保留集合语义。
    head_keys = set(head_map)
    toc_keys = set(toc_map)

    # 目录里天然会有几条"正文没有对应标题"的条目：摘要、Abstract、目录本身、
    # 参考文献、致谢。它们的标题段落不带 heading 样式（或就是目录页自身的标题），
    # 因此不会出现在正文标题集合里。用 --toc-only-ok 显式承认这些条目。
    toc_only_ok = set()
    for part in (args.toc_only_ok or '').split(','):
        if part.strip():
            toc_only_ok.add(re.sub(r'[\s\u3000]+', '', part.strip()))

    def matched(k, pool):
        if k in pool:
            return True
        return any(k in p or p in k for p in pool if p and k)

    only_toc_all = sorted(toc_map[k] for k in toc_keys
                          if not matched(k, head_keys))
    only_toc = [t for t in only_toc_all
                if re.sub(r'[\s\u3000]+', '', norm(t)) not in toc_only_ok]
    only_toc_ok = [t for t in only_toc_all
                   if re.sub(r'[\s\u3000]+', '', norm(t)) in toc_only_ok]
    only_body = sorted(head_map[k] for k in head_keys
                       if not matched(k, toc_keys))

    expected_missing = set()
    if args.expect_missing:
        for part in args.expect_missing.split(','):
            if part.strip():
                expected_missing.add(re.sub(r'[\s\u3000]+', '', part.strip()))
    still_missing = [m for m in only_body if key(m) in expected_missing]
    # 预期缺失的条目若已经出现在目录里，说明目录已刷新，应当把它从白名单里去掉
    refreshed = [m for m in expected_missing if matched(m, toc_keys)]
    only_body = [m for m in only_body if key(m) not in expected_missing]

    # 目录页标题与封面/摘要等前置标题不在目录里属于正常，排除
    toc_title_key = key(args.toc_title)
    only_body = [m for m in only_body if key(m) != toc_title_key]

    problems = []
    if only_toc:
        problems.append('只在目录、正文没有: ' + ' / '.join(only_toc))
    if only_body:
        problems.append('只在正文、目录没有: ' + ' / '.join(only_body))
    stale_note = []
    if still_missing:
        stale_note.append('目录缓存尚未包含（需在 Word 中更新域）: '
                          + ' / '.join(still_missing))
    if refreshed:
        stale_note.append('白名单条目已出现在目录里，请从 --expect-missing 移除: '
                          + ' / '.join(sorted(refreshed)))

    out = {
        'docx': args.docx,
        'tocStyles': sorted(toc_styles),
        'headingStyles': sorted(heading_styles),
        'tocEntries': len(toc),
        'bodyHeadings': len(heads),
        'onlyInToc': only_toc,
        'onlyInTocAllowed': sorted(only_toc_ok),
        'onlyInBody': only_body,
        'expectedMissingFromToc': sorted(still_missing),
        'whitelistNowPresent': sorted(refreshed),
        'problems': problems,
    }
    if args.json:
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print('== 目录 / 正文标题核对:', os.path.basename(args.docx), '==')
        print('目录样式 %s | 标题样式 %s' % (sorted(toc_styles), sorted(heading_styles)))
        print('目录条目 %d | 正文标题 %d' % (len(toc), len(heads)))
        for n in stale_note:
            print('[NOTE] ' + n)
        if problems:
            for p in problems:
                print('[FAIL] ' + p)
        else:
            print('[PASS] 双向一致（目录与正文章节逐项对应）')
    sys.exit(1 if problems else 0)


if __name__ == '__main__':
    main()
