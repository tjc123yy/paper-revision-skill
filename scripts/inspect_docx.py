# -*- coding: utf-8 -*-
"""论文 docx 基线探查 —— 提取批注、样式号、参考文献区间、表格、字体、图片，
并生成可直接编辑的 task.json 模板（供 verify_revision.py 使用）。

不含任何学生/章节/关键字硬编码；所有判定阈值都可用参数覆盖。

用法：
    python inspect_docx.py <docx路径> [-o 输出.json]
                           [--heading-styles 2,3] [--ref-heading 参考文献]
                           [--ref-end 评语,致谢,附录]

输出：JSON 写到 -o（默认 <docx同目录>/<名字>.inspect.json），
      终端只打印计数摘要（不打印列表内容）。
"""
import argparse
import json
import os
import sys

import docxlib as D

DEFAULT_REF_HEADING = '参考文献'
DEFAULT_REF_END = ('评语', '致谢', '附录', '外文资料', '致  谢')
DEFAULT_HEADING_NAME_PREFIX = ('heading', '标题')


def is_heading(style_id, style_map, extra_styles):
    if not style_id:
        return False
    if extra_styles and style_id in extra_styles:
        return True
    name = (style_map.get(style_id) or '').strip().lower()
    return name.startswith(DEFAULT_HEADING_NAME_PREFIX)


def detect_ref_section(paras, style_map, extra_styles, heading_text, end_markers):
    """定位参考文献区间，返回 (headingStyleId, items, startIndex, endIndex)。"""
    start = None
    for i, p in enumerate(paras):
        if D.get_text(p).strip() == heading_text:
            start = i
            break
    if start is None:
        return None, [], None, None

    ref_style = D.para_style(paras[start])
    items = []
    stop = len(paras)
    for j in range(start + 1, len(paras)):
        p = paras[j]
        txt = D.get_text(p).strip()
        style = D.para_style(p)
        if txt and any(txt.startswith(m) for m in end_markers):
            stop = j
            break
        if txt and is_heading(style, style_map, extra_styles) and style != ref_style:
            stop = j
            break
        if txt:
            items.append(txt)
    return ref_style, items, start, stop


def scan_tables(root):
    """返回表格摘要；caption 为表格上方最近的题注段（与 verify 共用定位逻辑）。"""
    out = []
    for ti, (tbl, caption) in enumerate(D.iter_body_tables(root)):
        rows = tbl.findall(D.qn('tr'))
        first_row = ''
        if rows:
            first_row = '|'.join(D.get_text(tc) for tc in rows[0].findall(D.qn('tc')))
        out.append({
            'index': ti,
            'rows': len(rows),
            'cols': len(rows[0].findall(D.qn('tc'))) if rows else 0,
            'firstRow': first_row,
            'caption': caption,
        })
    return out


def font_distribution(root):
    """统计 eastAsia 字体分布，并为每种字体附一个样例段落文本。"""
    dist = {}
    for p in D.iter_paras(root):
        for r in p.findall(D.qn('r')):
            rPr = r.find(D.qn('rPr'))
            if rPr is None:
                continue
            rf = rPr.find(D.qn('rFonts'))
            if rf is None:
                continue
            f = rf.get(D.qn('eastAsia'))
            if not f:
                continue
            entry = dist.setdefault(f, {'runs': 0, 'sampleParagraph': ''})
            entry['runs'] += 1
            if not entry['sampleParagraph']:
                txt = D.get_text(p).strip()
                entry['sampleParagraph'] = txt[:40]
    return [dict(font=k, **v) for k, v in sorted(dist.items(), key=lambda kv: -kv[1]['runs'])]


def build(path, opts):
    z, root, doc_bytes = D.open_docx(path)
    style_map = D.read_styles(z)
    paras = list(D.iter_paras(root))
    extra_styles = set(opts.heading_styles)

    comments = D.list_comments(z)
    markers = D.count_comment_markers(doc_bytes)

    sect_prs = root.findall('.//' + D.qn('sectPr'))
    pgnums = []
    for s in sect_prs:
        pg = s.find(D.qn('pgNumType'))
        if pg is not None:
            pgnums.append({'fmt': pg.get(D.qn('fmt')), 'start': pg.get(D.qn('start'))})

    settings = z.read(D.SETTINGS) if D.SETTINGS in z.namelist() else b''
    sup_nums, sup_empty = D.superscript_numbers(root)

    ref_style, ref_items, ref_start, ref_stop = detect_ref_section(
        paras, style_map, extra_styles, opts.ref_heading, opts.ref_end
    )

    headings = []
    for p in paras:
        sid = D.para_style(p)
        if sid and is_heading(sid, style_map, extra_styles):
            txt = D.get_text(p).strip()
            if txt:
                headings.append({'text': txt, 'styleId': sid,
                                 'styleName': style_map.get(sid)})

    tables = scan_tables(root)
    media = D.list_media(z)
    fonts = font_distribution(root)

    uniq = sorted(set(sup_nums))
    missing = [n for n in range(1, (max(uniq) if uniq else 0) + 1) if n not in uniq]

    suggested = {
        'docx': os.path.basename(path),
        'refCount': len(ref_items),
        'refHeadingText': opts.ref_heading,
        'refHeadingStyle': ref_style,
        'refEndMarkers': list(opts.ref_end),
        'sections': {
            'count': len(sect_prs),
            'expected': pgnums,
        },
        'superscripts': {
            'expectCount': len(sup_nums),
            'expectCoverage': (max(uniq) if uniq else 0),
        },
        'tables': [
            {
                'index': t['index'],
                'label': t['caption'] or ('表格' + str(t['index'])),
                'rows': t['rows'],
            }
            for t in tables
        ],
        'media': {m['name']: m['bytes'] for m in media},
        'fontChecks': [],
        'commentFilesExpectZero': True,
    }

    return {
        'docx': {
            'path': path,
            'sha256': D.sha256_file(path),
            'bytes': os.path.getsize(path),
        },
        'paragraphCount': len(paras),
        'comments': {'count': len(comments), 'items': comments},
        'commentMarkers': markers,
        'styleMap': style_map,
        'headings': headings,
        'refSection': {
            'headingText': opts.ref_heading,
            'styleId': ref_style,
            'startIndex': ref_start,
            'stopIndex': ref_stop,
            'count': len(ref_items),
            'items': ref_items,
        },
        'sections': {'sectPrCount': len(sect_prs), 'pgNumTypes': pgnums},
        'updateFields': settings.count(b'updateFields') > 0,
        'superscripts': {
            'count': len(sup_nums),
            'unique': uniq,
            'missing': missing,
            'emptyText': sup_empty,
        },
        'tables': tables,
        'media': media,
        'fontsEastAsia': fonts,
        'suggestedConfig': suggested,
    }


def main():
    ap = argparse.ArgumentParser(description='论文 docx 基线探查')
    ap.add_argument('docx')
    ap.add_argument('-o', '--out', default=None)
    ap.add_argument('--heading-styles', default='',
                    help='逗号分隔的额外标题样式号（样式名不含 heading/标题 时用）')
    ap.add_argument('--ref-heading', default=DEFAULT_REF_HEADING)
    ap.add_argument('--ref-end', default=','.join(DEFAULT_REF_END))
    opts = ap.parse_args()

    opts.heading_styles = [s for s in opts.heading_styles.split(',') if s]
    opts.ref_end = tuple(s for s in opts.ref_end.split(',') if s)

    if not os.path.isfile(opts.docx):
        print('文件不存在:', opts.docx)
        sys.exit(2)

    result = build(opts.docx, opts)

    out = opts.out
    if out is None:
        base = os.path.splitext(os.path.basename(opts.docx))[0]
        out = os.path.join(os.path.dirname(os.path.abspath(opts.docx)),
                           base + '.inspect.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 摘要只打印计数，不打印可能被过滤的列表内容
    print('== 探查结果 ==')
    print('段落数:', result['paragraphCount'])
    print('批注条数:', result['comments']['count'])
    print('批注标记:', result['commentMarkers'])
    print('样式数:', len(result['styleMap']), ' 标题段数:', len(result['headings']))
    print('参考文献: 标题样式=', result['refSection']['styleId'],
          ' 条数=', result['refSection']['count'])
    print('sectPr 数:', result['sections']['sectPrCount'],
          ' pgNumType 数:', len(result['sections']['pgNumTypes']))
    print('updateFields:', result['updateFields'])
    print('上标数:', result['superscripts']['count'],
          ' 编号种类:', len(result['superscripts']['unique']),
          ' 缺失个数:', len(result['superscripts']['missing']),
          ' 空文本上标:', result['superscripts']['emptyText'])
    print('表格数:', len(result['tables']), ' 图片数:', len(result['media']),
          ' 中文字体种类:', len(result['fontsEastAsia']))
    print('已写出:', out)


if __name__ == '__main__':
    main()
