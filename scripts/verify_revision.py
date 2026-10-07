# -*- coding: utf-8 -*-
"""论文修改验收脚本（配置驱动，字节级）

全部检查项都由 task.json 提供，脚本内不含任何学生名、章节号、样式号或关键字常量。
task.json 可由 inspect_docx.py 自动生成 suggestedConfig 后再按需调整。

用法：
    python verify_revision.py <docx路径> [--config task.json] [命令行覆盖项]

命令行覆盖（优先于 config）：
    --refs N              期望参考文献条数
    --ref-heading TEXT    参考文献标题文本（默认 参考文献）
    --ref-style ID        参考文献标题样式号
    --ref-end TEXT,..     参考文献区间结束标记
    --heading-styles 2,3  标题样式号
    --sect-count N        期望 sectPr 数
    --sup-count N         期望上标数
    --coverage N          期望上标编号覆盖 1..N

退出码：0 = 无 FAIL（允许 SKIP）；1 = 存在 FAIL。

注意：本脚本只打印计数、布尔与逗号连接的编号串，不打印含方括号编号的原始字符串。
"""
import argparse
import json
import os
import sys

import docxlib as D

FAIL = 'FAIL'
PASS = 'PASS'
SKIP = 'SKIP'


class Report:
    def __init__(self):
        self.rows = []

    def add(self, status, name, detail=''):
        self.rows.append((status, name, detail))
        print('[%s] %s%s' % (status, name, (' -- ' + detail) if detail else ''))

    def ok(self, name, detail=''):
        self.add(PASS, name, detail)

    def fail(self, name, detail=''):
        self.add(FAIL, name, detail)

    def skip(self, name, detail=''):
        self.add(SKIP, name, detail)

    def failed(self):
        return [r for r in self.rows if r[0] == FAIL]


def load_config(path):
    if not path:
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def default_ref_items(paras, style_map, heading_extra, heading_text, end_markers):
    """按与 inspect_docx.py 相同的规则切出参考文献条目。"""
    start = None
    for i, p in enumerate(paras):
        if D.get_text(p).strip() == heading_text:
            start = i
            break
    if start is None:
        return None, []
    ref_style = D.para_style(paras[start])
    items = []
    for j in range(start + 1, len(paras)):
        p = paras[j]
        txt = D.get_text(p).strip()
        style = D.para_style(p)
        if txt and any(txt.startswith(m) for m in end_markers):
            break
        if txt and is_heading(style, style_map, heading_extra) and style != ref_style:
            break
        if txt:
            items.append(txt)
    return ref_style, items


def is_heading(style_id, style_map, extra_styles):
    if not style_id:
        return False
    if extra_styles and style_id in extra_styles:
        return True
    name = (style_map.get(style_id) or '').strip().lower()
    return name.startswith(('heading', '标题'))


def active_bold(run):
    """run 是否处于生效的加粗状态。"""
    rPr = run.find(D.qn('rPr'))
    if rPr is None:
        return False
    b = rPr.find(D.qn('b'))
    if b is None:
        return False
    val = b.get(D.qn('val'))
    return val is None or val in ('1', 'true', 'on')


def main():
    ap = argparse.ArgumentParser(description='论文修改验收（配置驱动）')
    ap.add_argument('docx')
    ap.add_argument('--config', default=None)
    ap.add_argument('--refs', type=int, default=None)
    ap.add_argument('--ref-heading', default=None)
    ap.add_argument('--ref-style', default=None)
    ap.add_argument('--ref-end', default=None)
    ap.add_argument('--heading-styles', default=None)
    ap.add_argument('--sect-count', type=int, default=None)
    ap.add_argument('--sup-count', type=int, default=None)
    ap.add_argument('--coverage', type=int, default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print('文件不存在:', args.docx)
        sys.exit(2)

    cfg = load_config(args.config)
    rep = Report()

    ref_heading = args.ref_heading or cfg.get('refHeadingText') or '参考文献'
    ref_style_cfg = args.ref_style or cfg.get('refHeadingStyle')
    ref_end = (args.ref_end.split(',') if args.ref_end
               else cfg.get('refEndMarkers') or ['评语', '致谢', '附录'])
    heading_extra = set(
        (args.heading_styles or ','.join(cfg.get('headingStyles') or [])).split(',')
    ) - {''}

    z, root, doc_bytes = D.open_docx(args.docx)
    names = z.namelist()
    style_map = D.read_styles(z)
    paras = list(D.iter_paras(root))

    print('== 验收:', os.path.basename(args.docx), '==')

    # --- 1. 批注部件 ---
    comment_parts = [n for n in names if 'comment' in n.lower() or n.endswith('people.xml')]
    if comment_parts:
        rep.fail('批注部件为 0', '仍存在: ' + ','.join(comment_parts))
    else:
        rep.ok('批注部件为 0')

    # --- 2. 批注标记 ---
    markers = D.count_comment_markers(doc_bytes)
    bad = {k: v for k, v in markers.items() if v}
    if bad:
        rep.fail('批注标记为 0', json.dumps(bad, ensure_ascii=False))
    else:
        rep.ok('批注标记为 0')

    # --- 3. 分节与页码 ---
    sect_cfg = cfg.get('sections') or {}
    expect_count = args.sect_count if args.sect_count is not None else sect_cfg.get('count')
    sect_prs = root.findall('.//' + D.qn('sectPr'))
    pgnums = []
    for s in sect_prs:
        pg = s.find(D.qn('pgNumType'))
        pgnums.append({'fmt': pg.get(D.qn('fmt')), 'start': pg.get(D.qn('start'))}
                      if pg is not None else None)
    if expect_count is None:
        rep.skip('sectPr 数量', 'config 未指定，实际 %d' % len(sect_prs))
    elif len(sect_prs) == expect_count:
        rep.ok('sectPr 数量 = %d' % expect_count)
    else:
        rep.fail('sectPr 数量', '期望 %d 实际 %d' % (expect_count, len(sect_prs)))

    expect_pg = sect_cfg.get('expected') or []
    if not expect_pg:
        rep.skip('页码格式', 'config 未指定')
    else:
        missing = [e for e in expect_pg if e not in pgnums]
        if missing:
            rep.fail('页码格式', '缺少: ' + json.dumps(missing, ensure_ascii=False))
        else:
            rep.ok('页码格式齐全')

    # --- 4. updateFields ---
    settings = z.read(D.SETTINGS) if D.SETTINGS in names else b''
    if settings.count(b'updateFields') > 0:
        rep.ok('settings.xml updateFields')
    else:
        rep.fail('settings.xml updateFields', '未找到')

    # --- 5. 参考文献 ---
    ref_style, ref_items = default_ref_items(
        paras, style_map, heading_extra, ref_heading, tuple(ref_end)
    )
    if ref_style is None:
        rep.fail('参考文献标题存在', ref_heading)
    else:
        rep.ok('参考文献标题存在', '样式=' + str(ref_style))
        if ref_style_cfg and ref_style != ref_style_cfg:
            rep.fail('参考文献标题样式', '期望 %s 实际 %s' % (ref_style_cfg, ref_style))
        elif ref_style_cfg:
            rep.ok('参考文献标题样式 = ' + str(ref_style_cfg))

    expect_refs = args.refs if args.refs is not None else cfg.get('refCount')
    if expect_refs is None:
        rep.skip('参考文献条数', 'config 未指定，实际 %d' % len(ref_items))
    elif len(ref_items) == expect_refs:
        rep.ok('参考文献条数 = %d' % expect_refs)
    else:
        rep.fail('参考文献条数', '期望 %d 实际 %d' % (expect_refs, len(ref_items)))

    # --- 6. 上标引用 ---
    sup_nums, sup_empty = D.superscript_numbers(root)
    uniq = sorted(set(sup_nums))
    sup_cfg = cfg.get('superscripts') or {}
    expect_sup = args.sup_count if args.sup_count is not None else sup_cfg.get('expectCount')
    expect_cov = args.coverage if args.coverage is not None else sup_cfg.get('expectCoverage')
    max_empty = sup_cfg.get('maxEmpty', 0)

    if expect_sup is None:
        rep.skip('上标引用数', 'config 未指定，实际 %d' % len(sup_nums))
    elif len(sup_nums) == expect_sup:
        rep.ok('上标引用数 = %d' % expect_sup)
    else:
        rep.fail('上标引用数', '期望 %d 实际 %d' % (expect_sup, len(sup_nums)))

    if sup_empty > max_empty:
        rep.fail('无空文本上标', '空文本上标 %d 个' % sup_empty)
    else:
        rep.ok('无空文本上标')

    if expect_cov is None:
        rep.skip('上标编号覆盖', 'config 未指定，实际编号种类 %d' % len(uniq))
    else:
        missing = [n for n in range(1, expect_cov + 1) if n not in uniq]
        if missing:
            rep.fail('上标编号覆盖 1..%d' % expect_cov,
                     '缺: ' + ','.join(str(n) for n in missing))
        else:
            rep.ok('上标编号覆盖 1..%d' % expect_cov)

    # --- 7. 表格 ---
    # 两种写法：
    #   expectTables: [{"captionContains": "表5-1", "rows": 6}]  —— 按题注定位，推荐
    #   tables:       [{"index": 0, "rows": 9}]                 —— 按 body 内序号定位
    # 论文里常夹着文本框表格（无题注、行数随内容变化），按序号定位容易错位，
    # 所以优先用题注定位；带 captionContains 的项若命中多张表会被判为歧义 FAIL。
    expect_tables = cfg.get('expectTables') or []
    if expect_tables:
        body_tables = list(D.iter_body_tables(root))
        for spec in expect_tables:
            needle = spec.get('captionContains') or ''
            idx = spec.get('index')
            rows_want = spec.get('rows')
            label = spec.get('label') or (needle or ('表格%s' % idx))
            hits = []
            if needle:
                for tbl, cap in body_tables:
                    if cap and needle in cap:
                        hits.append((tbl, cap))
            elif idx is not None and 0 <= idx < len(body_tables):
                hits.append(body_tables[idx])
            if not hits:
                rep.fail('表格存在: ' + label,
                         '题注含 %r 的表格未找到' % needle)
                continue
            if len(hits) > 1:
                rep.fail('表格题注唯一: ' + label,
                         '%r 命中 %d 张表' % (needle, len(hits)))
                continue
            trs = hits[0][0].findall(D.qn('tr'))
            if rows_want is not None and len(trs) != rows_want:
                rep.fail('表格行数: ' + label, '期望 %d 实际 %d' % (rows_want, len(trs)))
            else:
                rep.ok('表格: ' + label, '行数 %d' % len(trs))

    tbl_cfg = cfg.get('tables') or []
    if not expect_tables and not tbl_cfg:
        rep.skip('关键表格', 'config 未指定')
    elif tbl_cfg:
        body_tables = list(D.iter_body_tables(root))
        for spec in tbl_cfg:
            label = spec.get('label') or '表格'
            idx = spec.get('index')
            rows_want = spec.get('rows')
            target = None
            if idx is not None:
                if 0 <= idx < len(body_tables):
                    target = body_tables[idx][0]
            else:
                # 退路：按首行关键字模糊匹配（多张同结构表格时会撞车，优先用 index）
                needle = spec.get('firstRowContains') or ''
                for tbl, _ in body_tables:
                    trs = tbl.findall(D.qn('tr'))
                    if not trs:
                        continue
                    first = ''.join(D.get_text(tc) for tc in trs[0].findall(D.qn('tc')))
                    if needle and needle in first:
                        target = tbl
                        break
            if target is None:
                rep.fail('表格存在: ' + label,
                         'index=%s 或首行关键字=%r 未命中' % (idx, spec.get('firstRowContains')))
                continue
            trs = target.findall(D.qn('tr'))
            if rows_want is not None and len(trs) != rows_want:
                rep.fail('表格行数: ' + label,
                         '期望 %d 实际 %d' % (rows_want, len(trs)))
            else:
                rep.ok('表格: ' + label, '行数 %d' % len(trs))

    # --- 7b. 格式验收（段落级 / run 级有效格式） ---
    # 由 config 的 formatChecks 驱动；检查项缺失即 SKIP。
    #   paragraphs[]: {"role": "...", "contains": "...", "expect": {
    #                    "line":400,"lineRule":"exact","before":0,"after":0,
    #                    "firstLine":480,"jc":"both","eastAsia":"宋体",
    #                    "ascii":"Times New Roman","sz":24}}
    #   匹配规则：段落文本包含 contains（可选：styleId 精确匹配，允许多段命中，
    #   要求全部命中段落都符合 expect）。
    fmt_cfg = cfg.get('formatChecks') or {}
    para_specs = fmt_cfg.get('paragraphs') or []

    def para_attrs(p):
        pPr = p.find(D.qn('pPr'))
        out = {}
        if pPr is not None:
            sp = pPr.find(D.qn('spacing'))
            if sp is not None:
                for a in ('before', 'after', 'line', 'lineRule'):
                    v = sp.get(D.qn(a))
                    if v is not None:
                        out[a] = v
            ind = pPr.find(D.qn('ind'))
            if ind is not None:
                for a in ('firstLine', 'firstLineChars'):
                    v = ind.get(D.qn(a))
                    if v is not None:
                        out[a] = v
            jc = pPr.find(D.qn('jc'))
            if jc is not None:
                out['jc'] = jc.get(D.qn('val'))
        # run 级：取段内第一个有文字的 run
        for r in p.findall(D.qn('r')):
            if not ''.join(t.text or '' for t in r.findall(D.qn('t'))).strip():
                continue
            rPr = r.find(D.qn('rPr'))
            if rPr is None:
                break
            rf = rPr.find(D.qn('rFonts'))
            if rf is not None:
                for a in ('eastAsia', 'ascii'):
                    v = rf.get(D.qn(a))
                    if v:
                        out[a] = v
            sz = rPr.find(D.qn('sz'))
            if sz is not None:
                out['sz'] = sz.get(D.qn('val'))
            break
        return out

    if not para_specs:
        if fmt_cfg:
            rep.skip('格式验收(段落级)', 'config 未给出 paragraphs')
    else:
        # 目录条目与封面/声明等前置部分的样式号与正文标题有重叠（封面标题也是
        # 大纲级别 0、目录条目也可能带 styleId=3），这些区域本不该套用正文标题
        # 规范，所以按样式号定位时要先切掉前置区域与目录条目。
        toc_styles = set(fmt_cfg.get('tocStyles') or ['21', '23', '14'])
        body_marker = fmt_cfg.get('bodyStartContains') or ''
        paras_seq = list(paras)
        body_start = 0
        if body_marker:
            for i, p in enumerate(paras_seq):
                if body_marker in D.get_text(p):
                    body_start = i
                    break
        body_paras = paras_seq[body_start:]

        for spec in para_specs:
            role = spec.get('role') or spec.get('contains') or '?'
            needle = spec.get('contains') or ''
            want = spec.get('expect') or {}
            sid_want = spec.get('styleId')
            pool = paras_seq if spec.get('scope') == 'all' else body_paras
            matched = []
            skipped = 0
            for p in pool:
                if sid_want is not None:
                    if D.para_style(p) != sid_want:
                        continue
                    if D.para_style(p) in toc_styles:
                        skipped += 1
                        continue
                if needle and needle not in D.get_text(p):
                    continue
                matched.append(p)
            if not matched:
                rep.fail('格式段存在: ' + role, '未找到包含 %r 的段落' % needle)
                continue
            bad = []
            for p in matched:
                got = para_attrs(p)
                diff = {k: (v, got.get(k)) for k, v in want.items()
                        if str(got.get(k)) != str(v)}
                if diff:
                    bad.append((D.get_text(p).strip()[:14], diff))
            if bad:
                txt, diff = bad[0]
                detail = '; '.join('%s 期望%s 实际%s' % (k, v[0], v[1])
                                   for k, v in sorted(diff.items()))
                rep.fail('格式: ' + role,
                         '%d/%d 段不符，例「%s」: %s'
                         % (len(bad), len(matched), txt, detail))
            else:
                rep.ok('格式: ' + role, '%d 段全部符合' % len(matched))

    # --- 8. 图片字节 ---
    media_cfg = {k: v for k, v in (cfg.get('media') or {}).items() if not k.endswith('/')}
    if not media_cfg:
        rep.skip('图片字节', 'config 未指定')
    else:
        missing = [n for n in media_cfg if n not in names]
        mismatched = []
        for name, want in media_cfg.items():
            if name in names and want is not None and z.getinfo(name).file_size != want:
                mismatched.append('%s(期望 %d 实际 %d)'
                                  % (name, want, z.getinfo(name).file_size))
        if missing:
            rep.fail('图片部件存在', '缺失 %d 个: %s' % (len(missing), ','.join(sorted(missing)[:5])))
        elif mismatched:
            rep.fail('图片字节一致', '不一致 %d 个: %s' % (len(mismatched), '; '.join(mismatched[:5])))
        else:
            rep.ok('图片字节一致', '共 %d 张' % len(media_cfg))

    # --- 9. 字体抽查 ---
    font_cfg = cfg.get('fontChecks') or []
    if not font_cfg:
        rep.skip('字体抽查', 'config 未指定')
    else:
        for spec in font_cfg:
            needle = spec.get('paragraphContains') or ''
            want = spec.get('expectEastAsia')
            found = set()
            target = None
            for p in paras:
                if needle and needle in D.get_text(p):
                    target = p
                    break
            if target is None:
                rep.fail('字体段存在: ' + needle, '未找到该段落')
                continue
            for r in target.findall(D.qn('r')):
                rPr = r.find(D.qn('rPr'))
                if rPr is None:
                    continue
                rf = rPr.find(D.qn('rFonts'))
                if rf is not None:
                    f = rf.get(D.qn('eastAsia'))
                    if f:
                        found.add(f)
            if want is None:
                rep.skip('字体: ' + needle, 'config 未给期望值')
            elif want in found:
                rep.ok('字体: ' + needle + ' = ' + want)
            else:
                rep.fail('字体: ' + needle,
                         '期望 %s 实际 %s' % (want, ','.join(sorted(found)) or '无'))

    # --- 10. 标题不加粗 ---
    if cfg.get('requireHeadingsNotBold'):
        offenders = []
        for p in paras:
            sid = D.para_style(p)
            if not is_heading(sid, style_map, heading_extra):
                continue
            if any(active_bold(r) for r in p.findall(D.qn('r'))):
                offenders.append(D.get_text(p).strip()[:20])
        if offenders:
            rep.fail('标题不加粗', '仍加粗 %d 处: %s'
                     % (len(offenders), ' / '.join(offenders[:5])))
        else:
            rep.ok('标题不加粗')
    else:
        rep.skip('标题不加粗', 'config 未要求')

    # --- 11. document.xml 结构不变量 ---
    # 重新打包/整段重写最容易静默丢东西：图片(drawing)、目录书签(bookmarkStart)、
    # 域(fldChar)、表格。用基线计数卡住，任何一项变化都算 FAIL。
    inv_cfg = cfg.get('docCounts') or {}
    if not inv_cfg:
        rep.skip('结构不变量', 'config 未指定')
    else:
        doc_text = doc_bytes.decode('utf-8', 'ignore')
        bad = []
        for needle, want in inv_cfg.items():
            got = doc_text.count(needle)
            if got != want:
                bad.append('%s 期望 %d 实际 %d' % (needle, want, got))
        if bad:
            rep.fail('结构不变量', '; '.join(bad))
        else:
            rep.ok('结构不变量', '共 %d 项' % len(inv_cfg))

    print('')
    failed = rep.failed()
    if failed:
        print('验收结论: 存在 %d 项 FAIL' % len(failed))
        sys.exit(1)
    print('验收结论: 全部通过（含 SKIP）')
    sys.exit(0)


if __name__ == '__main__':
    main()
