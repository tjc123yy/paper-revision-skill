# -*- coding: utf-8 -*-
"""归纳：把语料记录聚合为「专业 × 批注类型」矩阵与候选技巧，重写派生文档。

这是"技能自动进化"的第二步——完全确定性、可重复。它只重写
references/learned-patterns.md 里 AUTO:START/END 之间的区块，block 之外的手写内容不动。

第三个输出是"进化队列"：未分类批注样本。把它们的新说法补进 corpus/taxonomy.json，
下一轮分类覆盖率就上升——分类法自身也是这样长大的。

用法：
    python evolve.py [--min-count 3] [--top-majors 10] [--dry-run]
"""
import argparse
import datetime
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import corpus_lib as C

OUT_MD = os.path.join(SKILL_DIR, 'references', 'learned-patterns.md')
AUTO_START = '<!-- AUTO:START -->'
AUTO_END = '<!-- AUTO:END -->'

HEADER = chr(10).join([
    '# 自动归纳：专业 × 批注类型',
    '',
    '> 本文件由 `scripts/evolve.py` 自动生成维护。**AUTO 区块内的内容不要手工编辑**，',
    '> 下次运行会被覆盖；区块之外可以自由补写人工结论。',
    '',
    '数据来源：`corpus/records/*.json`（工作区论文批注 + 只读导入的平台台账）。',
    '',
])


def truncate(text, n=70):
    text = (text or '').replace(chr(10), ' ').replace(chr(13), ' ').strip()
    return text if len(text) <= n else text[:n] + '…'


def aggregate(records, min_count, top_majors, origins_filter=None):
    """聚合语料。

    分类**不信任记录里存的 category**（那是写入时的快照），而是用当前 taxonomy.json
    对原始文本重新分类——这样改完 taxonomy 重跑 evolve 就立刻生效，无需重建记录。
    记录里显式设置了 category_fixed 的（人工纠正）优先。

    origins_filter：只统计指定来源渠道的意见。用来剔除"平台模型生成的改写要求"，
    只比较导师/评审的人类意见，否则专业对比会被渠道差异污染。
    """
    taxonomy = C.load_taxonomy()
    papers = [r for r in records if (r.get('comment_count') or 0) > 0]
    total_comments = sum(len(r.get('comments') or []) for r in papers)

    by_source = {}
    origins = {}
    per_major = {}
    per_cat = {}
    cat_records = {}
    matrix = {}
    samples = {}
    uncategorized = []

    for rec in papers:
        rid = rec.get('id') or 'unknown-record'
        major = rec.get('major') or 'unknown'
        kept = []
        for c in rec.get('comments') or []:
            origin = c.get('origin') or 'unknown'
            if origins_filter and origin not in origins_filter:
                continue
            kept.append((c, origin))
        if not kept:
            continue
        by_source[rec.get('source') or 'unknown'] = \
            by_source.get(rec.get('source') or 'unknown', 0) + 1
        m = per_major.setdefault(major, {'papers': 0, 'comments': 0})
        m['papers'] += 1
        for c, origin in kept:
            origins[origin] = origins.get(origin, 0) + 1
            if c.get('category_fixed'):
                cat = c['category_fixed']
            else:
                cat = C.classify(c.get('text', ''), taxonomy)[0]
            m['comments'] += 1
            per_cat[cat] = per_cat.get(cat, 0) + 1
            cat_records.setdefault(cat, {})
            cat_records[cat][rid] = cat_records[cat].get(rid, 0) + 1
            matrix.setdefault(cat, {})
            matrix[cat][major] = matrix[cat].get(major, 0) + 1
            samples.setdefault(cat, [])
            t = truncate(c.get('text'))
            if t and t not in samples[cat] and len(samples[cat]) < 3:
                samples[cat].append(t)
            if cat == C.UNCATEGORIZED:
                uncategorized.append(t)

    total_comments = sum(v['comments'] for v in per_major.values())
    majors = sorted(per_major, key=lambda k: -per_major[k]['comments'])
    cats = sorted(per_cat, key=lambda k: -per_cat[k])

    # 单篇主导检测：某类型若绝大多数条目都来自同一篇论文，那它是"这篇论文的意见"，
    # 不能当成"这个专业的规律"。这是防止把 1 篇的证据读成专业结论的关键防线。
    dominated = []
    for cat in cats:
        total = per_cat[cat]
        if total < 3:
            continue
        recs = cat_records.get(cat) or {}
        if not recs:
            continue
        top_rid, top_n = max(recs.items(), key=lambda kv: kv[1])
        share = top_n / total
        if share >= 0.67:
            dominated.append({'category': cat, 'records': len(recs),
                              'top_share': share, 'top_count': top_n,
                              'total': total, 'top_record': top_rid})

    return {
        'papers': sum(by_source.values()),
        'by_source': by_source,
        'origins': origins,
        'origins_filter': sorted(origins_filter) if origins_filter else None,
        'total_comments': total_comments,
        'majors': majors,
        'cats': cats,
        'cat_records': cat_records,
        'dominated': dominated,
        'top_majors': majors[:top_majors],
        'per_major': per_major,
        'per_cat': per_cat,
        'matrix': matrix,
        'samples': samples,
        'uncategorized': uncategorized,
        'min_count': min_count,
    }


def render_markdown(agg, records):
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    L = []
    L.append(AUTO_START)
    L.append('')
    L.append('_生成时间：%s；语料：%d 条记录，其中 %d 条含批注。_'
             % (now, len(records), agg['papers']))
    L.append('')

    # 概况
    un = len(agg['uncategorized'])
    rate = (un / agg['total_comments'] * 100) if agg['total_comments'] else 0
    L.append('## 语料概况')
    L.append('')
    if agg.get('origins_filter'):
        L.append('> **本视图已按来源渠道过滤**，仅统计：%s。'
                 '用于剔除模型生成内容、只比较人类意见。'
                 % '、'.join(agg['origins_filter']))
        L.append('')
    L.append('| 记录来源 | 条数 |')
    L.append('|---|---|')
    label = {'workspace-docx': '工作区论文（docx 批注）',
             'opinion-intake': '文件夹意见（txt/图片/意见docx）',
             'platform-db': '平台台账（只读导入）'}
    for src in sorted(agg['by_source'], key=lambda k: -agg['by_source'][k]):
        L.append('| %s | %d |' % (label.get(src, src), agg['by_source'][src]))
    L.append('| **合计可学习记录** | **%d** |' % agg['papers'])
    L.append('')
    L.append('| 指标 | 值 |')
    L.append('|---|---|')
    L.append('| 修改意见总数 | %d 条 |' % agg['total_comments'])
    L.append('| 覆盖专业 | %d 个 |' % len([m for m in agg['majors'] if m != 'unknown']))
    L.append('| 未分类意见 | %d 条（%.1f%%） |' % (un, rate))
    L.append('')
    if agg['origins']:
        L.append('| 意见来源渠道 | 条数 |')
        L.append('|---|---|')
        for o in sorted(agg['origins'], key=lambda k: -agg['origins'][k]):
            L.append('| %s | %d |' % (o, agg['origins'][o]))
        L.append('')
        if agg['total_comments']:
            top_o, top_n = max(agg['origins'].items(), key=lambda kv: kv[1])
            share = top_n / agg['total_comments'] * 100
            if share > 50:
                L.append('> ⚠️ **「%s」占全部意见的 %.0f%%**，统计重心被单一渠道主导。' % (top_o, share))
                L.append('> 该渠道多为**模型生成的改写要求**，与导师原始批注性质不同；'
                         '解读"专业差异"前应先按 origin 拆分，不要直接混算。')
                L.append('')
    if rate > 20:
        L.append('> ⚠️ **未分类占比 %.1f%% 偏高**：把下方"进化队列"里的新说法补进 '
                 '`corpus/taxonomy.json`，再重跑 evolve。' % rate)
        L.append('')

    # 专业覆盖
    L.append('## 专业覆盖')
    L.append('')
    L.append('| 专业 | 记录数 | 意见数 | 主要批注类型 |')
    L.append('|---|---|---|---|')
    for major in agg['majors']:
        info = agg['per_major'][major]
        cat_counts = [(c, agg['matrix'][c].get(major, 0)) for c in agg['cats']]
        cat_counts = [x for x in cat_counts if x[1] > 0]
        cat_counts.sort(key=lambda kv: -kv[1])
        tops = '、'.join('%s(%d)' % (c, n) for c, n in cat_counts[:3]) or '—'
        L.append('| %s | %d | %d | %s |'
                 % (major, info['papers'], info['comments'], tops))
    L.append('')

    # 类型总览
    L.append('## 批注类型总览')
    L.append('')
    L.append('| 类型 | 条数 | 占比 | 涉及专业数 | 涉及记录数 |')
    L.append('|---|---|---|---|---|')
    for cat in agg['cats']:
        n = agg['per_cat'][cat]
        majors_hit = len([m for m, v in agg['matrix'][cat].items() if v > 0])
        recs_hit = len((agg.get('cat_records') or {}).get(cat) or {})
        L.append('| %s | %d | %.1f%% | %d | %d |'
                 % (cat, n, n / agg['total_comments'] * 100 if agg['total_comments'] else 0,
                    majors_hit, recs_hit))
    L.append('')

    # 矩阵
    L.append('## 专业 × 类型 矩阵')
    L.append('')
    cols = agg['top_majors']
    L.append('| 类型 | ' + ' | '.join(cols) + ' | 合计 |')
    L.append('|---' * (len(cols) + 2) + '|')
    for cat in agg['cats']:
        row = [str(agg['matrix'][cat].get(m, 0) or '.') for m in cols]
        L.append('| %s | %s | %d |' % (cat, ' | '.join(row), agg['per_cat'][cat]))
    L.append('')

    # 候选技巧
    L.append('## 候选技巧（条数 ≥ %d）' % agg['min_count'])
    L.append('')
    candidates = [c for c in agg['cats']
                  if agg['per_cat'][c] >= agg['min_count'] and c != C.UNCATEGORIZED]
    if not candidates:
        L.append('_暂无：语料还太少。继续按 record_task.py 记录任务即可累积。_')
        L.append('')
    for cat in candidates:
        n = agg['per_cat'][cat]
        majors_hit = [m for m in agg['majors'] if agg['matrix'][cat].get(m)]
        L.append('### %s（%d 条，涉及 %d 个专业）' % (cat, n, len(majors_hit)))
        L.append('')
        L.append('- 涉及专业：%s' % '、'.join(majors_hit))
        L.append('- 代表批注：')
        for s in agg['samples'][cat]:
            L.append('  - %s' % s)
        L.append('')

    # 单篇主导
    L.append('## 单篇主导的类型（不足以代表专业）')
    L.append('')
    dom = agg.get('dominated') or []
    if not dom:
        L.append('_无：每个类型（≥3 条）的证据都分散在 2 篇以上记录。_')
    else:
        L.append('以下类型的绝大多数条目来自**同一篇论文**。它们只能算'
                 '"这篇论文的意见"，**不能**当成"这个专业的规律"：')
        L.append('')
        L.append('| 类型 | 条数 | 来自记录数 | 最集中的一篇占比 |')
        L.append('|---|---|---|---|')
        for d in dom:
            L.append('| %s | %d | %d | %.0f%%（%s） |'
                     % (d['category'], d['total'], d['records'],
                        d['top_share'] * 100, d['top_record']))
    L.append('')

    # 覆盖缺口
    thin = [(m, agg['per_major'][m]['papers']) for m in agg['majors']
            if agg['per_major'][m]['papers'] < 3 and m != 'unknown']
    L.append('## 覆盖缺口（结论尚不稳定）')
    L.append('')
    if not thin:
        L.append('_所有专业都已有 ≥3 篇带批注论文。_')
    else:
        L.append('以下专业样本不足 3 篇，**目前还不能据此下专业特有结论**：')
        L.append('')
        for m, n in thin:
            L.append('- %s：%d 篇' % (m, n))
    L.append('')

    # 进化队列
    L.append('## 进化队列：未分类批注样本')
    L.append('')
    if not agg['uncategorized']:
        L.append('_无。分类覆盖率 100%。_')
    else:
        L.append('把这些说法（或其共同词）补进 `corpus/taxonomy.json` 的对应 categories[].keywords，')
        L.append('再重跑 `evolve.py` 即可提高覆盖率：')
        L.append('')
        seen = set()
        shown = 0
        for t in agg['uncategorized']:
            if t in seen:
                continue
            seen.add(t)
            L.append('- %s' % t)
            shown += 1
            if shown >= 20:
                break
        if len(seen) > shown:
            L.append('- _(还有 %d 条，见 corpus/stats.json)_' % (len(seen) - shown))
    L.append('')
    L.append(AUTO_END)
    return chr(10).join(L)


def write_markdown(body, out_path=None):
    out_path = out_path or OUT_MD
    old = ''
    if os.path.isfile(out_path):
        with open(out_path, 'r', encoding='utf-8') as f:
            old = f.read()
    if AUTO_START in old and AUTO_END in old:
        pre = old.split(AUTO_START)[0]
        post = old.split(AUTO_END, 1)[1]
        new = pre + body + post
    else:
        new = HEADER + chr(10) + body + chr(10)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(new)
    return out_path


def main():
    ap = argparse.ArgumentParser(description='归纳语料，重写 learned-patterns.md')
    ap.add_argument('--min-count', type=int, default=3)
    ap.add_argument('--top-majors', type=int, default=10)
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--records-dir', default=None)
    ap.add_argument('--origin', default=None,
                    help='只统计指定来源渠道，逗号分隔，如 docx-comment,image-sidecar')
    ap.add_argument('--out', default=None, help='输出到另一个 md 文件（不覆盖主视图）')
    args = ap.parse_args()

    origins_filter = None
    if args.origin:
        origins_filter = set(s for s in args.origin.split(',') if s)

    records = C.load_records(args.records_dir)
    agg = aggregate(records, args.min_count, args.top_majors, origins_filter)

    # 用 --out 生成附加视图时，统计写到同名 sidecar，避免覆盖主视图的 stats.json
    stats_path = C.STATS_FILE if not args.out else os.path.splitext(args.out)[0] + '.stats.json'
    with open(stats_path, 'w', encoding='utf-8') as f:
        json.dump({
            'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
            'records_total': len(records),
            'learnable_records': agg['papers'],
            'by_source': agg['by_source'],
            'origin_channels': agg['origins'],
            'total_comments': agg['total_comments'],
            'per_major': agg['per_major'],
            'per_category': agg['per_cat'],
            'matrix': agg['matrix'],
            'uncategorized_samples': agg['uncategorized'],
        }, f, ensure_ascii=False, indent=2)

    body = render_markdown(agg, records)
    if args.dry_run:
        print(body)
    else:
        print('已写出:', write_markdown(body, args.out))
    print('已写出:', stats_path)

    print('')
    print('== 归纳结果 ==')
    print('  可学习记录:', agg['papers'], '(%s)'
          % ', '.join('%s=%d' % kv for kv in sorted(agg['by_source'].items())))
    print('  意见总数  :', agg['total_comments'])
    if agg['origins']:
        print('  来源渠道  :', ', '.join('%s=%d' % kv
                                          for kv in sorted(agg['origins'].items())))
    print('  专业数    :', len([m for m in agg['majors'] if m != 'unknown']))
    un = len(agg['uncategorized'])
    rate = (un / agg['total_comments'] * 100) if agg['total_comments'] else 0
    print('  未分类    :', un, '(%.1f%%)' % rate)
    if rate > 20:
        print('  !! 未分类占比偏高，请把新说法补进 corpus/taxonomy.json 后重跑')
    thin = [m for m in agg['majors']
            if agg['per_major'][m]['papers'] < 3 and m != 'unknown']
    if thin:
        print('  !! 样本不足 3 篇的专业:', '、'.join(thin))


if __name__ == '__main__':
    main()
