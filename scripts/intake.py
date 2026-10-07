# -*- coding: utf-8 -*-
"""修改意见采集入口：支持 docx 批注之外的多种来源。

支持的来源：
    * docx 批注（word/comments.xml）
    * docx 正文（导师单独发的"修改意见.docx"，正文即意见）
    * .txt / .md（按编号行、空行分段或逐行自动切分）
    * 图片（批注截图/拍照）—— 脚本只登记待办，由 agent 用 read_image 转写
    * .pdf —— 有同名 .txt 侧车则用，否则登记待办

图片处理流程（agent 在环）：
    1. intake.py pending <根目录>      列出待转写的图片
    2. agent 对每张图用 read_image 看图，把文字写到同名 .txt（侧车文件）
    3. 重跑 intake.py commit/scan      侧车文本自动纳入语料

用法：
    python intake.py scan    <目标文件或目录>
    python intake.py commit  <目标> [--paper 论文.docx] [--student 姓名] [--major 专业]
                                     [--level 普通本科] [--ptype 问题对策型] [--date YYYY-MM-DD]
    python intake.py pending [<根目录>]
    python intake.py dirs    [<根目录>]     查找约定名的意见目录
"""
import argparse
import datetime
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import corpus_lib as C

PENDING_FILE = os.path.join(C.CORPUS_DIR, 'pending_vision.json')
SKIP_DIRS = {'.preview', '.runtime', '123-repo', 'paper-revision',
             '_task', '_task_学生Q', '_work', 'tpl', 'output', 'corpus'}


def print_report(res):
    print('== 采集结果 ==')
    for name, n, kind in res['report']:
        print('  %-46s %-4s %s' % (name[:46], n, kind))
    print('  合计意见条数:', len(res['opinions']))
    origins = C.count_origins(res['opinions'])
    if origins:
        print('  来源分布:', ', '.join('%s=%d' % kv for kv in sorted(origins.items())))
    cats = {}
    for o in res['opinions']:
        cats[o.get('category')] = cats.get(o.get('category'), 0) + 1
    if cats:
        print('  分类分布:')
        for k, v in sorted(cats.items(), key=lambda kv: -kv[1]):
            print('    %-14s %d' % (k, v))
        un = cats.get(C.UNCATEGORIZED, 0)
        if un:
            print('  !! %d 条未分类，建议补进 corpus/taxonomy.json' % un)
    if res['pending']:
        print('')
        print('== 待 agent 视觉/转写（%d 条）==' % len(res['pending']))
        for p in res['pending']:
            print('  [%s] %s  -> 转写后写到: %s'
                  % (p['reason'], p['name'], p['sidecar_expected']))


def save_pending(pending):
    os.makedirs(C.CORPUS_DIR, exist_ok=True)
    with open(PENDING_FILE, 'w', encoding='utf-8') as f:
        json.dump(pending, f, ensure_ascii=False, indent=2)
    return PENDING_FILE


def cmd_scan(args):
    res = C.collect_opinions(args.target, split_mode=args.split)
    C.classify_comments(res['opinions'], C.load_taxonomy())
    print_report(res)
    if res['pending']:
        print('')
        print('已写出待办:', save_pending(res['pending']))
    return 0


def cmd_commit(args):
    res = C.collect_opinions(args.target, split_mode=args.split)
    C.classify_comments(res['opinions'], C.load_taxonomy())
    print_report(res)

    date_str = args.date or datetime.date.today().isoformat()
    platform = C.load_platform()

    if args.paper:
        if not os.path.isfile(args.paper):
            print('论文不存在:', args.paper)
            return 2
        rec = C.build_record(args.paper, student=args.student, major=args.major,
                             level=args.level, ptype=args.ptype, date_str=date_str,
                             notes=args.notes, platform=platform,
                             opinions=res['opinions'])
    else:
        opinions = res['opinions']
        if not opinions:
            print('没有采集到任何意见，未写出记录。')
            return 2
        stem = os.path.basename(os.path.abspath(args.target))
        major = args.major
        major_source = 'explicit'
        if not major:
            major, major_source = C.infer_major(args.target, student=args.student,
                                                platform_majors=platform.get('majors'))
        rec = {
            'id': C.record_id(date_str, args.student, 'opinions-' + stem),
            'date': date_str,
            'student': args.student,
            'major': major or 'unknown',
            'major_source': major_source,
            'level': args.level,
            'ptype': args.ptype,
            'source': 'opinion-intake',
            'docx': None,
            'comment_count': len(opinions),
            'comments': opinions,
            'origin_counts': C.count_origins(opinions),
            'pending': res['pending'],
            'verification': None,
            'notes': args.notes,
        }

    if res['pending']:
        rec['pending'] = res['pending']
        save_pending(res['pending'])

    # 记录意见目录的相对路径，供 backfill 判断"这个目录已有人工采集过"，避免重复入库
    if os.path.isdir(args.target):
        try:
            rec['opinion_dir'] = os.path.relpath(os.path.abspath(args.target),
                                                 C.workspace_root())
        except Exception:
            rec['opinion_dir'] = os.path.abspath(args.target)

    path = C.save_record(rec, args.out_dir)
    print('')
    print('== 已写入语料记录 ==')
    print('  id      :', rec['id'])
    print('  学生    :', rec['student'])
    print('  专业    :', rec['major'], '(来源:' + str(rec['major_source']) + ')')
    print('  意见条数:', rec['comment_count'])
    print('  来源分布:', rec.get('origin_counts'))
    print('  写出    :', path)
    return 0


def cmd_pending(args):
    """只扫「约定名的意见目录」，绝不整盘扫描工作区。

    工作区里 _work/_task 等目录有大量任务过程渲染图（页面截图、media 抽取图），
    整盘扫描会把它们全列成"待转写意见"，几百条噪音。
    """
    root = args.root or C.workspace_root()
    found = C.find_opinion_targets(root, skip_dirs=SKIP_DIRS)
    targets = list(found['dirs']) + list(found['files'])
    print('意见目标: %d（目录 %d / 文件 %d）'
          % (len(targets), len(found['dirs']), len(found['files'])))
    for d in found['dirs']:
        print('  [dir] ', os.path.relpath(d, root))
    for f in found['files']:
        print('  [file]', os.path.relpath(f, root))

    # 显式 --target 才额外采集指定目录/文件
    if getattr(args, 'target', None):
        targets.append(args.target)

    all_pending = []
    for t in targets:
        res = C.collect_opinions(t, split_mode=args.split)
        for p in res['pending']:
            p['from_dir'] = os.path.relpath(t, root)
        all_pending.extend(res['pending'])

    if all_pending:
        print('')
        print('== 待转写 %d 条 ==' % len(all_pending))
        for p in all_pending:
            print('  [%s] %s' % (p['reason'], os.path.relpath(p['source_file'], root)))
            print('        侧车写到: %s' % os.path.basename(p['sidecar_expected']))
        print('')
        print('已写出:', save_pending(all_pending))
    else:
        print('无待转写项。')
        save_pending([])
    return 0


def cmd_dirs(args):
    root = args.root or C.workspace_root()
    found = C.find_opinion_targets(root, skip_dirs=SKIP_DIRS)
    print('== 意见目录 %d 个 ==' % len(found['dirs']))
    for d in found['dirs']:
        print('  [dir] ', os.path.relpath(d, root))
    print('== 意见文件 %d 个 ==' % len(found['files']))
    for f in found['files']:
        print('  [file]', os.path.relpath(f, root))
    return 0


def main():
    ap = argparse.ArgumentParser(description='修改意见采集（多来源）')
    sub = ap.add_subparsers(dest='cmd')

    p1 = sub.add_parser('scan', help='只侦查，不写记录')
    p1.add_argument('target')
    p1.add_argument('--split', default='auto',
                    choices=['auto', 'line', 'blank', 'number', 'whole'])
    p1.set_defaults(func=cmd_scan)

    p2 = sub.add_parser('commit', help='采集并写入语料记录')
    p2.add_argument('target')
    p2.add_argument('--paper', default=None, help='关联的论文 docx（意见与其合为一条记录）')
    p2.add_argument('--student', default=None)
    p2.add_argument('--major', default=None)
    p2.add_argument('--level', default=None)
    p2.add_argument('--ptype', default=None)
    p2.add_argument('--date', default=None)
    p2.add_argument('--notes', default='')
    p2.add_argument('--out-dir', default=None)
    p2.add_argument('--split', default='auto',
                    choices=['auto', 'line', 'blank', 'number', 'whole'])
    p2.set_defaults(func=cmd_commit)

    p3 = sub.add_parser('pending', help='列出待视觉转写的图片/pdf（只扫意见目录）')
    p3.add_argument('root', nargs='?', default=None)
    p3.add_argument('--target', default=None, help='额外指定一个意见目录或文件')
    p3.add_argument('--split', default='auto')
    p3.set_defaults(func=cmd_pending)

    p4 = sub.add_parser('dirs', help='查找约定名的意见目录')
    p4.add_argument('root', nargs='?', default=None)
    p4.set_defaults(func=cmd_dirs)

    args = ap.parse_args()
    if not getattr(args, 'func', None):
        ap.print_help()
        return 2
    return args.func(args)


if __name__ == '__main__':
    sys.exit(main())
