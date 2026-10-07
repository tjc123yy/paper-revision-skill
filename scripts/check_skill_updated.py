# -*- coding: utf-8 -*-
"""技能回写自检（每次论文修改收尾时跑一遍）。

用户要求：每改完一篇论文，必须把"新发现的问题、怎么解决的、造了什么新方法/新工具"
总结进技能。本脚本就是这条要求的**判定器**——不靠记忆，靠文件事实。

检查五项（对应 SKILL.md「维护与进化」收尾四件事）：

1. `references/evolution-log.md` 有**今天**的任务条目（## YYYY-MM-DD 开头）；
2. `SKILL.md` 在任务窗口内被改过（陷阱/验证清单/执行顺序至少要动一处）；
3. `corpus/records/` 有今天、该学生的语料记录，且 `--docx` 的 sha256 与记录一致；
4. `references/learned-patterns.md`、`corpus/stats.json` 在记录之后被重算过（`evolve.py`）；
5. **新增/改动的通用脚本登记在 SKILL.md 的「资源索引」里**（把"临时脚本"固化成工具的那一步）。

退出码：0 = 五项齐全；1 = 有缺项（缺什么、怎么补都会打印出来）。

用法：
    python check_skill_updated.py --student 学生A
    python check_skill_updated.py --student 学生A --docx "学生A/xxx_已修改.docx"
    python check_skill_updated.py --student 学生A --since 2026-10-06T16:30
"""
import argparse
import datetime as dt
import hashlib
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                     # paper-revision/
SKILL = os.path.join(ROOT, 'SKILL.md')
LOG = os.path.join(ROOT, 'references', 'evolution-log.md')
PATTERNS = os.path.join(ROOT, 'references', 'docx-xml-patterns.md')
LEARNED = os.path.join(ROOT, 'references', 'learned-patterns.md')
STATS = os.path.join(ROOT, 'corpus', 'stats.json')
RECORDS = os.path.join(ROOT, 'corpus', 'records')


def mtime(path):
    return os.path.getmtime(path) if os.path.exists(path) else None


def fmt(ts):
    return dt.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S') if ts else '-'


def main():
    ap = argparse.ArgumentParser(description='技能回写自检')
    ap.add_argument('--student', required=True, help='本次任务的学生姓名')
    ap.add_argument('--docx', default=None, help='本次成稿路径：校验语料记录里的 sha256')
    ap.add_argument('--since', default=None,
                    help='任务开始时间（YYYY-MM-DD 或 YYYY-MM-DDTHH:MM）；'
                         '默认取 evolution-log.md 的修改时间往前 12 小时')
    ap.add_argument('--date', default=None, help='任务日期 YYYY-MM-DD，默认今天')
    args = ap.parse_args()

    today = args.date or dt.date.today().strftime('%Y-%m-%d')
    msgs, problems = [], []

    # ---- 时间窗口 ----
    if args.since:
        since = dt.datetime.fromisoformat(args.since)
    else:
        base = mtime(LOG)
        since = dt.datetime.fromtimestamp(base) - dt.timedelta(hours=12) if base \
            else dt.datetime.now() - dt.timedelta(days=1)
    since_ts = since.timestamp()

    print('== 技能回写自检 ==')
    print('学生 %s | 任务日期 %s | 时间窗口起点 %s' % (args.student, today, fmt(since_ts)))
    print('')

    # ---- 1. evolution-log 今天的条目 ----
    log_txt = open(LOG, encoding='utf-8').read() if os.path.exists(LOG) else ''
    dates = re.findall(r'^##\s+(\d{4}-\d{2}-\d{2})', log_txt, re.M)
    if today in dates:
        msgs.append('evolution-log.md 有 %s 的条目' % today)
    else:
        problems.append('evolution-log.md 缺 %s 的任务条目（## %s …）：'
                        '把"新问题 / 怎么解决 / 新方法"追加进去' % (today, today))

    # ---- 2. SKILL.md 在窗口内被改过 ----
    m_skill = mtime(SKILL)
    if m_skill and m_skill >= since_ts:
        msgs.append('SKILL.md 已更新（%s）' % fmt(m_skill))
    else:
        problems.append('SKILL.md 未在任务窗口内更新（最后 %s）：'
                        '新坑要补进「关键陷阱」+「验证清单」+「脚本执行顺序」' % fmt(m_skill))

    # ---- 3. 语料记录 ----
    recs = [p for p in glob.glob(os.path.join(RECORDS, '*.json'))
            if os.path.basename(p).startswith(today) and args.student in os.path.basename(p)]
    if not recs:
        problems.append('corpus/records/ 缺 %s-%s 的语料记录：'
                        '跑 record_task.py <论文> --opinions <意见> --student %s --major <专业>'
                        % (today, args.student, args.student))
    else:
        rec_path = max(recs, key=os.path.getmtime)
        rec = json.load(open(rec_path, encoding='utf-8'))
        msgs.append('语料记录存在：%s（意见 %s 条，验收退出码 %s）'
                    % (os.path.basename(rec_path), rec.get('comment_count'),
                       (rec.get('verification') or {}).get('exit')))
        if args.docx:
            if not os.path.exists(args.docx):
                problems.append('--docx 指向的文件不存在：%s' % args.docx)
            else:
                h = hashlib.sha256(open(args.docx, 'rb').read()).hexdigest()
                got = (rec.get('docx') or {}).get('sha256')
                if got == h:
                    msgs.append('语料记录 sha256 与成稿一致')
                else:
                    problems.append('语料记录 sha256 与成稿不一致（记录 %s / 实file %s）：'
                                    '成稿在记录之后又被改过，重跑 record_task.py'
                                    % (str(got)[:12], h[:12]))
        # ---- 4. evolve 是否重算过 ----
        m_learn, m_stat = mtime(LEARNED), mtime(STATS)
        if m_learn and m_learn >= os.path.getmtime(rec_path) - 1:
            msgs.append('learned-patterns.md / stats.json 已在记录之后重算')
        else:
            problems.append('evolve.py 未在本条记录之后跑过（learned %s / stats %s）：'
                            '跑 evolve.py 重算归纳' % (fmt(m_learn), fmt(m_stat)))

    # ---- 5. 通用脚本是否登记在资源索引 ----
    skill_txt = open(SKILL, encoding='utf-8').read() if os.path.exists(SKILL) else ''
    new_scripts = [os.path.basename(p) for p in glob.glob(os.path.join(HERE, '*.py'))
                   if os.path.getmtime(p) >= since_ts]
    new_scripts = [n for n in new_scripts
                   if n not in ('check_skill_updated.py',)]
    if new_scripts:
        for n in new_scripts:
            if n in skill_txt:
                msgs.append('新脚本 %s 已登记在 SKILL.md' % n)
            else:
                problems.append('新脚本 scripts/%s 未登记在 SKILL.md「资源索引」：'
                                '补一行说明用途与相关陷阱编号' % n)
    else:
        msgs.append('本次窗口内没有新增通用脚本（若有，记得登记到资源索引）')

    # ---- 代码模式文件 ----
    m_pat = mtime(PATTERNS)
    if m_pat and m_pat >= since_ts:
        msgs.append('docx-xml-patterns.md 已更新')
    else:
        msgs.append('提示：docx-xml-patterns.md 本次未更新（有通用代码模式时应补）')

    print('--- 已做到 ---')
    for m in msgs:
        print('  [OK]   ' + m)
    print('')
    if problems:
        print('--- 还缺 ---')
        for p in problems:
            print('  [TODO] ' + p)
        print('')
        print('结论：技能回写未完成（%d 项待补）' % len(problems))
        sys.exit(1)
    print('结论：技能回写齐全')
    sys.exit(0)


if __name__ == '__main__':
    main()
