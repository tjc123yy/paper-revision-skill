# -*- coding: utf-8 -*-
"""捕获：把一次论文修改任务写成一条结构化语料记录。

这是"技能自动进化"的第一步。每次改完论文都必须执行，否则语料不增长、归纳也无从谈起。

用法：
    python record_task.py <docx路径> [--major 专业] [--student 姓名]
                          [--level 普通本科] [--ptype 问题对策型]
                          [--notes "本次任务要点"] [--date YYYY-MM-DD]
                          [--verify-exit 0]

专业不传时按 证据强度 自动推断：major_map.json → 封面「专业：」→ 文件名 → 平台台账；
全部失败记为 unknown，并在输出里提示补 major_map.json。
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import corpus_lib as C


def main():
    ap = argparse.ArgumentParser(description='捕获一条论文修改任务语料记录')
    ap.add_argument('docx')
    ap.add_argument('--major', default=None)
    ap.add_argument('--student', default=None)
    ap.add_argument('--level', default=None)
    ap.add_argument('--ptype', default=None)
    ap.add_argument('--notes', default='')
    ap.add_argument('--date', default=None)
    ap.add_argument('--verify-exit', type=int, default=None)
    ap.add_argument('--opinions', default=None,
                    help='额外的修改意见来源（文件或目录）：txt/md/意见docx/图片')
    ap.add_argument('--out-dir', default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print('文件不存在:', args.docx)
        sys.exit(2)

    platform = C.load_platform()
    verification = None
    if args.verify_exit is not None:
        verification = {'exit': args.verify_exit,
                        'passed': args.verify_exit == 0}

    opinions = []
    pending = []
    if args.opinions:
        res = C.collect_opinions(args.opinions)
        opinions = res['opinions']
        pending = res['pending']
        print('== 额外意见来源 ==')
        for name, n, kind in res['report']:
            print('  %-46s %-4s %s' % (name[:46], n, kind))
        print('  采集到 %d 条意见' % len(opinions))
        if pending:
            print('  !! %d 个文件需要 agent 视觉/转写后才能纳入：' % len(pending))
            for p in pending:
                print('     [%s] %s -> 转写写到 %s'
                      % (p['reason'], p['name'], p['sidecar_expected']))
        print('')

    rec = C.build_record(
        args.docx,
        student=args.student,
        major=args.major,
        level=args.level,
        ptype=args.ptype,
        date_str=args.date,
        notes=args.notes,
        verification=verification,
        platform=platform,
        opinions=opinions,
    )
    if pending:
        rec['pending'] = pending
    path = C.save_record(rec, args.out_dir)

    print('== 已捕获语料记录 ==')
    print('  记录 id :', rec['id'])
    print('  学生    :', rec['student'])
    print('  专业    :', rec['major'], '(来源:' + str(rec['major_source']) + ')')
    print('  意见条数:', rec['comment_count'])
    print('  来源分布:', rec.get('origin_counts'))
    print('  写出    :', path)

    if rec['comment_count']:
        dist = {}
        for c in rec['comments']:
            dist[c['category']] = dist.get(c['category'], 0) + 1
        print('  分类分布:')
        for name, n in sorted(dist.items(), key=lambda kv: -kv[1]):
            print('    %-12s %d' % (name, n))
        un = dist.get(C.UNCATEGORIZED, 0)
        if un:
            print('  !! 有 %d 条未分类，建议把新说法补进 corpus/taxonomy.json' % un)
    if not rec['major'] or rec['major'] == 'unknown':
        print('  !! 专业未能推断，请在 corpus/major_map.json 的 by_student/by_file 里补一条')


if __name__ == '__main__':
    main()
