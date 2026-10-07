# -*- coding: utf-8 -*-
"""回填：扫描工作区全部论文 + 只读导入一体化平台台账，生成初始语料。

默认只记录「有批注」的论文（没有批注就没有修改意见可学）；--all 可记录全部。
平台 DB 为只读导入：把 revisions.opinion 与 changelog.items 作为"平台来源批注"入库。

用法：
    python backfill.py [--root 工作区] [--all] [--force] [--skip-platform]
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import corpus_lib as C
import docxlib as D

# 非学生论文目录
SKIP_DIRS = {'.preview', '.runtime', '123-repo', 'paper-revision',
             '_task', '_task_学生Q', '_work', 'tpl', 'output'}


def iter_papers(root):
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        top = rel.split(os.sep)[0]
        if top in SKIP_DIRS:
            dirnames[:] = []
            continue
        dirnames[:] = [d for d in dirnames
                       if not d.startswith('.') and d not in SKIP_DIRS]
        for fn in sorted(filenames):
            if fn.lower().endswith('.docx') and not fn.startswith('~$'):
                yield os.path.join(dirpath, fn)


def existing_hashes(records_dir=None):
    seen = {}
    for rec in C.load_records(records_dir):
        h = (rec.get('docx') or {}).get('sha256')
        if h:
            seen[h] = rec.get('id')
    return seen


def purge_by_source(records_dir, source):
    """删除某来源的旧记录（记录 id 里含学生名，学生名推断改进后会变化，前缀匹配不可靠）。

    Windows 陷阱：必须先关闭文件句柄再 os.remove，在 with open(...) 里删会抛
    PermissionError（文件被占用），被 except 吞掉后表现为"静默不删"。
    """
    d = records_dir or C.RECORDS_DIR
    n = 0
    if not os.path.isdir(d):
        return 0
    for name in os.listdir(d):
        if not name.endswith('.json'):
            continue
        p = os.path.join(d, name)
        try:
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            continue
        if data.get('source') != source:
            continue
        try:
            os.remove(p)
            n += 1
        except OSError as exc:
            print('  !! 删除失败 %s: %s' % (name, exc))
    return n


def backfill_docx(root, keep_all=False, force=False, records_dir=None):
    taxonomy = C.load_taxonomy()
    platform = C.load_platform(verbose=True)
    seen = {} if force else existing_hashes(records_dir)
    if force:
        purged = purge_by_source(records_dir, 'workspace-docx')
        if purged:
            print('  清理上一轮论文记录 %d 条' % purged)
    added, skipped, failed = 0, 0, 0
    for path in iter_papers(root):
        try:
            z, _root, doc_bytes = D.open_docx(path)
            n_comments = len(D.list_comments(z))
        except Exception as exc:
            print('ERR 无法解析 %s: %s' % (os.path.basename(path), exc))
            failed += 1
            continue
        if not keep_all and n_comments == 0:
            skipped += 1
            continue
        h = D.sha256_file(path)
        if h in seen:
            skipped += 1
            continue
        rec = C.build_record(path, platform=platform)
        C.save_record(rec, records_dir)
        seen[h] = rec['id']
        added += 1
        print('  + %-52s 批注=%-3d 专业=%s(%s)'
              % (os.path.basename(path), rec['comment_count'],
                 rec['major'], rec['major_source']))
    return added, skipped, failed


def purge_pattern(records_dir, prefix):
    """删除某一前缀的旧记录文件，避免上一次运行留下的陈旧记录污染统计。

    只在确定要重新导入时才调用（先确认来源可用，再删）。
    """
    d = records_dir or C.RECORDS_DIR
    n = 0
    if os.path.isdir(d):
        for name in os.listdir(d):
            if name.startswith(prefix) and name.endswith('.json'):
                os.remove(os.path.join(d, name))
                n += 1
    return n


def backfill_platform(records_dir=None):
    """把平台 revisions 导入为语料记录（只读，不改平台）。

    两道去重：
      1. 内容签名去重——平台里有成批的冒烟测试记录（同样的 opinion + changelog 重复 6 次），
         不去重会把个别批注类型灌成几百条，统计直接失真。
      2. 重新导入前清掉上一轮的 platform-rev*.json，避免陈旧记录残留
         （平台的 revision id 会变，只覆盖同 id 文件是不够的）。
    """
    platform = C.load_platform(verbose=True)
    if not platform['revisions']:
        return 0
    purged = purge_pattern(records_dir, 'platform-rev')
    if purged:
        print('  清理上一轮平台记录 %d 条' % purged)

    works = platform.get('works_by_id') or {}
    seen_sig = set()
    added, dup = 0, 0
    for rev in platform['revisions']:
        raw_items = [it['text'] for it in rev['items']]
        sig = (rev['opinion'].strip(), tuple(sorted(raw_items)))
        if sig in seen_sig:
            dup += 1
            continue
        seen_sig.add(sig)

        work = works.get(rev['work_id']) or {}
        major = work.get('major') or 'unknown'
        student = work.get('name')

        items = []
        if rev['opinion'].strip():
            items.append({'id': 'opinion', 'author': 'platform', 'origin': 'platform-opinion',
                          'text': rev['opinion'].strip(), 'section': None,
                          'source_file': os.path.basename(rev['source_path'] or '') or None})
        for i, it in enumerate(rev['items']):
            items.append({'id': 'item%d' % i, 'author': 'platform', 'origin': 'platform-changelog',
                          'text': it['text'], 'section': it.get('section'),
                          'source_file': os.path.basename(rev['source_path'] or '') or None})
        if not items:
            continue
        C.classify_comments(items, C.load_taxonomy())
        rec = {
            'id': 'platform-rev%d' % rev['id'],
            'date': None,
            'student': student,
            'major': major,
            'major_source': 'platform:works' if major != 'unknown' else 'unknown',
            'level': None,
            'ptype': None,
            'source': 'platform-db',
            'docx': {'name': os.path.basename(rev['source_path'] or '') or None,
                     'bytes': None, 'sha256': None},
            'comment_count': len(items),
            'comments': items,
            'origin_counts': C.count_origins(items),
            'commentMarkers': None,
            'verification': None,
            'notes': '由平台 revisions#%s 只读导入' % rev['id'],
        }
        C.save_record(rec, records_dir)
        added += 1
    if dup:
        print('  去重丢弃 %d 条重复的冒烟记录' % dup)
    return added


def backfill_opinions(root, records_dir=None, split_mode='auto'):
    """采集意见目录 + 散落在学生文件夹里的意见文件（如 学生D\\修改意见.txt）。

    图片/pdf 会在记录里留下 pending 条目，等 agent 用 read_image 转写后再重跑。
    """
    found = C.find_opinion_targets(root, skip_dirs=SKIP_DIRS)
    targets = list(found['dirs']) + list(found['files'])
    if targets:
        purged = purge_pattern(records_dir, 'opinions-')
        if purged:
            print('  清理上一轮意见记录 %d 条' % purged)

    # 已被 intake.py commit 人工采集过的目标不再自动入库，否则同一批意见会有两份记录、
    # 在 evolve 里被重复统计。
    covered = set()
    for r in C.load_records(records_dir):
        d = r.get('opinion_dir')
        if d:
            covered.add(os.path.normcase(d.replace('/', os.sep)))

    platform = C.load_platform()
    added = 0
    for t in targets:
        rel = os.path.relpath(t, root)
        if os.path.normcase(rel) in covered:
            print('  = %-44s 已有采集记录，跳过' % rel)
            continue
        res = C.collect_opinions(t, split_mode=split_mode)
        if not res['opinions'] and not res['pending']:
            continue
        C.classify_comments(res['opinions'], C.load_taxonomy())
        student = rel.split(os.sep)[0] if os.sep in rel else None
        major, major_src = C.infer_major(
            t, student=student, platform_majors=platform.get('majors'))
        if not major or major == 'unknown':
            # 意见文件本身没有专业信息时，用同目录里该生的论文来推断
            m2, s2 = C.sibling_paper_major(t, student=student,
                                           platform_majors=platform.get('majors'))
            if m2:
                major, major_src = m2, s2
        rec = {
            'id': 'opinions-' + C.safe_slug(rel.replace(os.sep, '-')),
            'date': None,
            'student': student,
            'major': major or 'unknown',
            'major_source': major_src,
            'level': None,
            'ptype': None,
            'source': 'opinion-intake',
            'docx': None,
            'comment_count': len(res['opinions']),
            'comments': res['opinions'],
            'origin_counts': C.count_origins(res['opinions']),
            'pending': res['pending'],
            'opinion_dir': rel,
            'verification': None,
            'notes': '来自意见目标 %s' % rel,
        }
        C.save_record(rec, records_dir)
        added += 1
        print('  + %-44s 意见=%-3d 专业=%s(%s) 待转写=%d'
              % (rel, len(res['opinions']), rec['major'], major_src,
                 len(res['pending'])))
    return added, targets


def main():
    ap = argparse.ArgumentParser(description='回填论文批注语料')
    ap.add_argument('--root', default=None)
    ap.add_argument('--all', action='store_true', help='连没有批注的论文一起记录')
    ap.add_argument('--force', action='store_true', help='忽略 sha256 去重，强制重建')
    ap.add_argument('--skip-platform', action='store_true')
    ap.add_argument('--skip-opinions', action='store_true')
    ap.add_argument('--out-dir', default=None)
    args = ap.parse_args()

    root = args.root or C.workspace_root()
    print('工作区根:', root)
    print('')
    print('== 扫描工作区 docx ==')
    added, skipped, failed = backfill_docx(root, keep_all=args.all,
                                           force=args.force, records_dir=args.out_dir)
    print('  新增 %d，跳过 %d（无批注或已存在），失败 %d' % (added, skipped, failed))

    oa = 0
    if not args.skip_opinions:
        print('')
        print('== 扫描意见目录 / 意见文件（txt/图片/意见docx）==')
        oa, targets = backfill_opinions(root, records_dir=args.out_dir)
        if not targets:
            print('  未发现意见目录（%s）或意见文件（文件名含 %s）'
                  % ('、'.join(C.OPINION_DIR_NAMES), '、'.join(C.OPINION_FILE_KEYWORDS)))
        else:
            print('  意见目标 %d 个，新增记录 %d 条' % (len(targets), oa))

    pa = 0
    if not args.skip_platform:
        print('')
        print('== 只读导入平台台账 ==')
        pa = backfill_platform(records_dir=args.out_dir)
        print('  导入 %d 条平台修改记录' % pa)

    total = len(C.load_records(args.out_dir))
    print('')
    print('语料记录总数:', total)


if __name__ == '__main__':
    main()
