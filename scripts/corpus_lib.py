# -*- coding: utf-8 -*-
"""语料层公共逻辑：批注分类、专业推断、记录读写、语料来源。

设计目标：让"技能自我进化"的前两步（捕获、归纳）完全确定性、可重复。
- 分类规则在 corpus/taxonomy.json（数据），不在代码里
- 专业映射在 corpus/major_map.json（数据），不在代码里
- 语料来源在 corpus/sources.json

不含任何学生姓名、专业名、论文题目的硬编码。
"""
import json
import os
import re
import sqlite3

import docxlib as D

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
CORPUS_DIR = os.path.join(SKILL_DIR, 'corpus')
RECORDS_DIR = os.path.join(CORPUS_DIR, 'records')
TAXONOMY_FILE = os.path.join(CORPUS_DIR, 'taxonomy.json')
MAJOR_MAP_FILE = os.path.join(CORPUS_DIR, 'major_map.json')
SOURCES_FILE = os.path.join(CORPUS_DIR, 'sources.json')
STATS_FILE = os.path.join(CORPUS_DIR, 'stats.json')

UNCATEGORIZED = '未分类'

# 封面「专业：XXX」。标签可能是「专业、层次：」这种复合写法，故允许标签后跟少量其它字。
MAJOR_COVER_RE = re.compile(r'^专\s*业[^\n：:]{0,8}[：:]\s*(.+)$')
MAJOR_LABEL_RE = re.compile(r'^专\s*业$')

# 目录里不该出现在专业位置的行
MAJOR_STOPWORDS = ('名称', '代码', '介绍', '英语', '课程')

# 只在封面附近找专业字段；正文里也会出现「专业」二字（如岗位、专业字段名），
# 越过这个行数就停止，避免把正文里的东西当成专业。
MAJOR_SCAN_LINES = 120


def _has_cjk(s, n=2):
    return len(re.findall('[\u4e00-\u9fff]', s or '')) >= n


def _cover_lines(text):
    return (text or '').replace('\r\n', chr(10)).replace('\r', chr(10)).split(chr(10))


def _next_value_line(lines, idx, max_skip=3):
    """取标签行之后最近的一个非空行（封面常用表格：标签与值分属不同段落）。"""
    for j in range(idx + 1, min(idx + 1 + max_skip, len(lines))):
        v = (lines[j] or '').strip().strip('　 ')
        if v:
            return v
    return None


def major_from_cover(text):
    """从封面取专业。兼容三种排版：

    1) 同行：`专业：计算机科学与技术`
    2) 表格分段落：一行只有 `专业`，值在下一行——国家开放大学模板就是这种，
       只有 1) 的正则会静默漏掉（学生F的论文就是这样被漏成 unknown 的）。
    3) 复合标签：`专业、层次：计算机科学与技术、本科`——按分隔符取第一段。
    """
    lines = _cover_lines(text)[:MAJOR_SCAN_LINES]
    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line:
            continue
        m = MAJOR_COVER_RE.match(line)
        if m:
            val = m.group(1).strip()
        elif MAJOR_LABEL_RE.match(line):
            val = _next_value_line(lines, i)
        else:
            continue
        if not val:
            continue
        # 值可能还带着别的字段，如「计算机科学与技术、本科」
        for seg in re.split('[、，,；;]', val):
            seg = seg.strip().strip('　 ')
            if not seg:
                continue
            # 必须是含中文的短词：挡掉正文里误命中的英文/拼音（如 Gangwei）
            if not _has_cjk(seg):
                continue
            if any(w in seg for w in MAJOR_STOPWORDS):
                continue
            if 2 <= len(seg) <= 24:
                return seg
    return None


def _read_json(path, default):
    if not os.path.isfile(path):
        return default
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_taxonomy():
    return _read_json(TAXONOMY_FILE, {'categories': []})


def load_major_map():
    return _read_json(MAJOR_MAP_FILE, {'by_student': {}, 'by_file': {}})


def load_sources():
    return _read_json(SOURCES_FILE, {})


def workspace_root(sources=None):
    """工作区根：sources.json 指定优先，否则取技能目录的上一级。"""
    sources = sources if sources is not None else load_sources()
    root = sources.get('workspace_root')
    if root and os.path.isdir(root):
        return root
    return os.path.dirname(SKILL_DIR)


# --------------------------------------------------------------------------
# 批注分类
# --------------------------------------------------------------------------

def classify(text, taxonomy):
    """按关键词长度加权打分分类。返回 (分类名, 分类id, 得分)。

    得分为命中关键词的长度之和——关键词越长越具体，自然压过泛化词。
    """
    low = (text or '').lower()
    best_name, best_id, best_score = UNCATEGORIZED, 'uncategorized', 0
    for cat in taxonomy.get('categories', []):
        score = 0
        for kw in cat.get('keywords', []):
            if kw and kw.lower() in low:
                score += len(kw)
        if score > best_score:
            best_score = score
            best_name = cat.get('name') or cat.get('id')
            best_id = cat.get('id')
    return best_name, best_id, best_score


def classify_comments(comments, taxonomy):
    """就地给每条批注补 category 字段；尊重已有的 category_fixed（人工纠正）。"""
    for c in comments:
        if c.get('category_fixed'):
            c['category'] = c['category_fixed']
            c['category_id'] = c.get('category_fixed_id') or 'fixed'
            c['category_score'] = -1
            continue
        name, cid, score = classify(c.get('text', ''), taxonomy)
        c['category'] = name
        c['category_id'] = cid
        c['category_score'] = score
    return comments


# --------------------------------------------------------------------------
# 专业推断
# --------------------------------------------------------------------------

def extract_doc_text(path):
    """提取 docx 纯文本（段落间换行），用于封面字段推断。"""
    import zipfile
    from lxml import etree
    try:
        with zipfile.ZipFile(path) as z:
            root = etree.fromstring(z.read(D.DOCUMENT))
    except Exception:
        return ''
    lines = []
    for p in root.iter(D.qn('p')):
        s = ''.join(t.text or '' for t in p.iter(D.qn('t'))).strip()
        if s:
            lines.append(s)
    return chr(10).join(lines)


def major_from_filename(path):
    """匹配『学号_姓名_专业_题目…』这类命名：下划线分段且第 3 段像专业名。"""
    stem = os.path.splitext(os.path.basename(path))[0]
    parts = [p for p in stem.split('_') if p]
    if len(parts) >= 4:
        cand = parts[2].strip()
        if 2 <= len(cand) <= 24 and not cand.isdigit():
            return cand
    return None


STUDENT_COVER_RE = re.compile(r'^(?:姓\s*名|学\s*生)\s*[：:]\s*(.+)$')
STUDENT_LABEL_RE = re.compile(r'^(?:姓\s*名|学\s*生)$')

GENERIC_DIRS = ('论文修改', '修改意见', '意见', '批注', 'output', 'templates',
                'desktop', 'users', 'public', 'documents')


def student_from_cover(text):
    """从封面取学生名。同样兼容『姓名：X』与表格分段两种排版。"""
    lines = _cover_lines(text)
    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line:
            continue
        m = STUDENT_COVER_RE.match(line)
        if m:
            val = m.group(1).strip().strip('　 ')
        elif STUDENT_LABEL_RE.match(line):
            val = _next_value_line(lines, i)
        else:
            continue
        if not val or any(ch.isdigit() for ch in val):
            continue
        if 2 <= len(val) <= 12:
            return val
    return None


def student_from_parent_dir(path):
    """学生论文通常放在以姓名命名的文件夹里（论文修改/<姓名>/xxx.docx）。"""
    parent = os.path.basename(os.path.dirname(os.path.abspath(path)))
    if not parent or parent.lower() in GENERIC_DIRS:
        return None
    if parent.startswith('.') or parent.startswith('_'):
        return None
    if 2 <= len(parent) <= 4 and all('\u4e00' <= ch <= '\u9fff' for ch in parent):
        return parent
    return None


def infer_major(path, student=None, text=None, platform_majors=None, mmap=None):
    """按证据强度依次推断专业。返回 (专业, 来源)。"""
    mmap = mmap if mmap is not None else load_major_map()
    by_student = mmap.get('by_student') or {}
    by_file = mmap.get('by_file') or {}

    fname = os.path.basename(path)
    if student and student in by_student:
        return by_student[student], 'map:student'
    if fname in by_file:
        return by_file[fname], 'map:file'

    if text is None:
        text = extract_doc_text(path)
    val = major_from_cover(text)
    if val:
        return val, 'cover'

    val = major_from_filename(path)
    if val:
        return val, 'filename'

    if student and platform_majors and student in platform_majors:
        return platform_majors[student], 'platform'

    return None, 'unknown'


# --------------------------------------------------------------------------
# 语料来源：平台 DB（只读）
# --------------------------------------------------------------------------

def load_platform(verbose=False):
    """只读读取一体化平台台账。返回 {'majors': {...}, 'revisions': [...]}。

    平台不存在或读取失败时返回空结构——技能本地语料不受影响。
    """
    out = {'majors': {}, 'revisions': [], 'works_by_id': {}}
    db_path = (load_sources() or {}).get('platform_db')
    if not db_path or not os.path.isfile(db_path):
        if verbose:
            print('平台 DB 不存在，跳过:', db_path)
        return out
    uri = 'file:%s?mode=ro' % db_path.replace('\\', '/')
    try:
        conn = sqlite3.connect(uri, uri=True)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        for r in cur.execute('SELECT id, name, major FROM works'):
            out['works_by_id'][r['id']] = {'name': r['name'], 'major': r['major']}
            if r['name'] and r['major']:
                out['majors'][r['name']] = r['major']
        for r in cur.execute(
                'SELECT id, work_id, opinion, changelog, source_path, created_at '
                'FROM revisions ORDER BY id'):
            items = []
            try:
                ck = json.loads(r['changelog'] or '[]')
                for entry in ck:
                    if not isinstance(entry, dict):
                        continue
                    for it in (entry.get('items') or []):
                        if it:
                            items.append({'section': entry.get('section'), 'text': it})
            except Exception:
                pass
            out['revisions'].append({
                'id': r['id'],
                'work_id': r['work_id'],
                'opinion': r['opinion'] or '',
                'items': items,
                'source_path': r['source_path'],
            })
        conn.close()
    except Exception as exc:
        if verbose:
            print('平台 DB 读取失败，跳过:', exc)
        return {'majors': {}, 'revisions': [], 'works_by_id': {}}
    if verbose:
        print('平台 DB: works 专业 %d 条, revisions %d 条'
              % (len(out['majors']), len(out['revisions'])))
    return out


# --------------------------------------------------------------------------
# 记录读写
# --------------------------------------------------------------------------

def safe_slug(text, fallback='rec'):
    out = re.sub('[^0-9A-Za-z\u4e00-\u9fff（）()\\-]+', '-', (text or '').strip())
    out = out.strip('-')
    return out[:80] or fallback


def sibling_paper_major(target, student=None, platform_majors=None):
    """意见文件/目录本身推断不出专业时，用同目录（或上一层学生目录）的论文 docx 推断。

    典型场景：`学生F\\修改意见.txt` 里没有专业字段，但同目录的
    `学生F-论文_xxx.docx` 封面写着「专业：计算机科学与技术」。
    """
    base = target if os.path.isdir(target) else os.path.dirname(target)
    for d in (base, os.path.dirname(base)):
        if not d or not os.path.isdir(d):
            continue
        try:
            names = sorted(os.listdir(d))
        except OSError:
            continue
        for n in names:
            if not n.lower().endswith('.docx') or n.startswith('~$'):
                continue
            major, src = infer_major(os.path.join(d, n), student=student,
                                     platform_majors=platform_majors)
            if major and major != 'unknown':
                return major, 'sibling-paper:' + src
    return None, None


def record_id(date_str, student, stem):
    return '%s-%s-%s' % (date_str, safe_slug(student or 'unknown'), safe_slug(stem))


def save_record(record, records_dir=None):
    records_dir = records_dir or RECORDS_DIR
    os.makedirs(records_dir, exist_ok=True)
    path = os.path.join(records_dir, safe_slug(record['id']) + '.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
    return path


def load_records(records_dir=None):
    records_dir = records_dir or RECORDS_DIR
    out = []
    if not os.path.isdir(records_dir):
        return out
    for name in sorted(os.listdir(records_dir)):
        if not name.endswith('.json'):
            continue
        with open(os.path.join(records_dir, name), 'r', encoding='utf-8') as f:
            try:
                out.append(json.load(f))
            except Exception as exc:
                print('记录解析失败，跳过 %s: %s' % (name, exc))
    return out


def build_record(path, student=None, major=None, level=None, ptype=None,
                 date_str=None, notes='', verification=None, platform=None,
                 opinions=None):
    """从一个 docx 构建语料记录（批注 + 分类 + 专业推断）。

    opinions 为 collect_opinions() 的额外产出（txt/图片/docx 正文等非批注意见），
    会与 docx 批注合并到同一份记录的 comments 里，并标注 origin。
    """
    import datetime

    date_str = date_str or datetime.date.today().isoformat()
    platform = platform if platform is not None else {'majors': {}}
    z, root, doc_bytes = D.open_docx(path)
    comments = D.list_comments(z)
    markers = D.count_comment_markers(doc_bytes)
    text = extract_doc_text(path)

    if not student:
        student = student_from_cover(text)
    if not student:
        student = student_from_parent_dir(path)
    if not student:
        stem = os.path.splitext(os.path.basename(path))[0]
        parts = [p for p in stem.split('_') if p]
        for p in parts[:2]:
            if p and not p.isdigit() and len(p) <= 4 and '\u4e00' <= p[0] <= '\u9fff':
                student = p
                break

    major_source = 'explicit'
    if not major:
        major, major_source = infer_major(path, student=student, text=text,
                                          platform_majors=platform.get('majors'),
                                          mmap=load_major_map())
    if not major:
        major = 'unknown'

    for c in comments:
        c.setdefault('origin', 'docx-comment')
        c.setdefault('source_file', os.path.basename(path))
    merged = list(comments) + list(opinions or [])
    classify_comments(merged, load_taxonomy())

    record = {
        'id': record_id(date_str, student, os.path.splitext(os.path.basename(path))[0]),
        'date': date_str,
        'student': student,
        'major': major,
        'major_source': major_source,
        'level': level,
        'ptype': ptype,
        'source': 'workspace-docx',
        'docx': {
            'name': os.path.basename(path),
            'bytes': os.path.getsize(path),
            'sha256': D.sha256_file(path),
        },
        'comment_count': len(merged),
        'comments': merged,
        'origin_counts': count_origins(merged),
        'commentMarkers': markers,
        'verification': verification,
        'notes': notes,
    }
    return record


def count_origins(comments):
    out = {}
    for c in comments or []:
        o = c.get('origin') or 'unknown'
        out[o] = out.get(o, 0) + 1
    return out


# --------------------------------------------------------------------------
# 多来源修改意见采集
#
# 修改意见不一定在 docx 批注里，常见还有：
#   * 文件夹里的 .txt / .md（一段一段或编号列举）
#   * 导师单独发的"修改意见.docx"（正文即意见，不带批注）
#   * 批注截图 / 拍照图片  -> 需要视觉识别，脚本只登记待办
#   * pdf
# 脚本能确定性地处理前两类；图片交给 agent 用 read_image 转写后写到同名 .txt 侧车文件，
# 再重跑采集即可自动纳入。这是"脚本做确定的事、视觉交给模型"的分工。
# --------------------------------------------------------------------------

OPINION_TEXT_EXT = ('.txt', '.md', '.markdown')
OPINION_IMAGE_EXT = D.IMAGE_EXT
OPINION_PDF_EXT = ('.pdf',)
OPINION_DIR_NAMES = ('意见', '_意见', '批注', '修改意见', '审阅意见', '批语', '意见汇总')

# docx 正文被当作意见时的长度上限。导师单独发的"修改意见.docx"通常几百到几千字；
# 毕业论文正文动辄两三万字，必须卡掉，否则整篇论文会被当成意见灌进语料。
# 阈值故意收得比较紧（8000 字）：宁可漏收一份超长的意见文档，也不能收进一篇论文。
MAX_BODY_AS_OPINION = 8000

NUMBER_LINE_RE = re.compile(r'^[ \t]*(?:[0-9]{1,2}[.、)）]|[-*\u00b7])\s*')


def split_opinions(text, mode='auto'):
    """把一段文本切成一条条意见。

    mode: auto | line（每行一条）| blank（空行分段）| number（编号行）| whole（整段一条）
    """
    text = (text or '').replace('\r\n', chr(10)).replace('\r', chr(10))
    if not text.strip():
        return []
    if mode == 'whole':
        return [text.strip()]
    lines = [l.strip() for l in text.split(chr(10))]
    nonempty = [l for l in lines if l]
    if mode == 'line':
        return nonempty
    blocks = [b.strip() for b in re.split(chr(10) + '{2,}', text) if b.strip()]
    if mode == 'blank':
        return [b.replace(chr(10), ' ') for b in blocks]

    numbered = [l for l in nonempty if NUMBER_LINE_RE.match(l)]
    if len(numbered) >= 2 and len(numbered) >= 0.5 * len(nonempty):
        return [NUMBER_LINE_RE.sub('', l).strip() for l in numbered]
    if len(blocks) >= 2:
        return [b.replace(chr(10), ' ') for b in blocks]
    if len(nonempty) > 3 and (sum(len(l) for l in nonempty) / len(nonempty)) < 120:
        return nonempty
    # 单段落内用中文序号分条（"一、…二、…"）时按序号切开——中文意见文档的常见写法
    segs = [s.strip() for s in
            re.split('(?=[一二三四五六七八九十]{1,2}、)', text.strip()) if s.strip()]
    if len(segs) >= 2:
        return segs
    return [text.strip()]


def sidecar_text(path):
    """图片/pdf 的同名 .txt/.md 转写侧车文件。"""
    base = os.path.splitext(path)[0]
    for ext in ('.txt', '.md'):
        p = base + ext
        if os.path.isfile(p):
            with open(p, 'r', encoding='utf-8', errors='replace') as f:
                return p, f.read()
    return None, None


def iter_files(target):
    if os.path.isfile(target):
        yield target
        return
    for dirpath, dirnames, filenames in os.walk(target):
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        for fn in sorted(filenames):
            if fn.startswith('~$'):
                continue
            yield os.path.join(dirpath, fn)


def collect_opinions(target, split_mode='auto', from_body=True):
    """采集 target（文件或目录）下的全部修改意见。

    返回 {'opinions': [...], 'pending': [...], 'report': [...]}
    pending 是需要 agent 视觉/转写介入的条目（图片、pdf）。
    """
    opinions = []
    pending = []
    report = []
    idx = 0

    def add(text, origin, source_file, author=None, section=None):
        nonlocal idx
        t = (text or '').strip()
        if not t:
            return
        idx += 1
        opinions.append({
            'id': 'op%d' % idx, 'author': author, 'text': t,
            'origin': origin, 'source_file': source_file, 'section': section,
        })

    if not os.path.exists(target):
        return {'opinions': [], 'pending': [], 'report': [('(不存在)', 0, target)]}

    files = list(iter_files(target))

    # 先算出哪些 .txt/.md 是图片/pdf 的侧车文件——它们的内容已计入图片，不能再作为
    # 独立 txt 意见重复计一遍（否则条数正好翻倍）。
    sidecar_paths = set()
    for path in files:
        ext = os.path.splitext(path)[1].lower()
        if ext in OPINION_IMAGE_EXT or ext in OPINION_PDF_EXT:
            base = os.path.splitext(path)[0]
            for e in ('.txt', '.md'):
                p = base + e
                if os.path.isfile(p):
                    sidecar_paths.add(os.path.normcase(os.path.abspath(p)))

    for path in files:
        name = os.path.basename(path)
        ext = os.path.splitext(path)[1].lower()
        side_path, side_text = sidecar_text(path)

        if ext in OPINION_IMAGE_EXT or ext in OPINION_PDF_EXT:
            is_img = ext in OPINION_IMAGE_EXT
            # 必须用 side_path is not None 判断，不能用 side_text 的真值：
            # 空侧车文件是"已审阅、无意见"的约定标记，内容为空但不应再列为待转写。
            if side_path is not None:
                items = split_opinions(side_text, split_mode)
                for t in items:
                    add(t, 'image-sidecar' if is_img else 'pdf-sidecar', name)
                report.append((name, len(items), 'sidecar'))
            else:
                pending.append({
                    'source_file': os.path.abspath(path), 'name': name,
                    'kind': ext.lstrip('.'),
                    'reason': 'needs_vision' if is_img else 'needs_text_extraction',
                    'sidecar_expected': os.path.splitext(os.path.abspath(path))[0] + '.txt',
                })
                report.append((name, 0, 'PENDING'))
            continue

        if ext in OPINION_TEXT_EXT:
            if os.path.normcase(os.path.abspath(path)) in sidecar_paths:
                report.append((name, 0, 'sidecar(已计入图片)'))
                continue
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                raw = f.read()
            items = split_opinions(raw, split_mode)
            for t in items:
                add(t, 'txt', name)
            report.append((name, len(items), 'txt'))
            continue

        if ext == '.docx':
            ok, reason = docx_opinion_verdict(path)
            if not ok:
                report.append((name, 0, 'SKIP:' + reason))
                continue
            try:
                z, _root, _bytes = D.open_docx(path)
            except Exception as exc:
                report.append((name, 0, 'ERR:' + str(exc)[:40]))
                continue
            cmts = D.list_comments(z)
            if cmts:
                for c in cmts:
                    add(c.get('text'), 'docx-comment', name, author=c.get('author'))
                report.append((name, len(cmts), 'docx-comment'))
                continue
            if not from_body:
                report.append((name, 0, 'no-comments'))
                continue
            body = extract_doc_text(path)
            items = split_opinions(body, split_mode)
            for t in items:
                add(t, 'docx-body', name)
            report.append((name, len(items), 'docx-body'))
            continue

    return {'opinions': opinions, 'pending': pending, 'report': report}


def find_opinion_dirs(root, skip_dirs=()):
    """在工作区里找约定名的意见目录（意见/批注/修改意见…）。"""
    out = []
    skip = set(skip_dirs)
    for dirpath, dirnames, _filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                       if not d.startswith('.') and d not in skip]
        for d in dirnames:
            if d in OPINION_DIR_NAMES:
                out.append(os.path.join(dirpath, d))
    return out


# 意见类文件名关键词。注意"意见"很宽松，会命中"修改意见.txt""审阅意见.txt"等，
# 这是有意的——学生把意见直接放在自己文件夹下时不会建约定目录。
OPINION_FILE_KEYWORDS = ('修改意见', '意见', '批注', '审阅', '评审', '批语')


def is_opinion_filename(name):
    base = os.path.splitext(name)[0]
    return any(k in base for k in OPINION_FILE_KEYWORDS)


def docx_opinion_verdict(path):
    """判断一个 docx 是"意见文档"还是"论文"。返回 (是否意见文档, 原因)。

    判据：
      * 带批注（word/comments.xml 非空）→ 是**论文**，其批注归 backfill_docx 处理，
        这里必须跳过，否则同一批批注会被算两遍。
      * 正文超过 MAX_BODY_AS_OPINION → 判定为论文正文，跳过。
      * 其余（短、无批注）→ 当作导师单独发的意见文档。
    """
    try:
        z, _root, _bytes = D.open_docx(path)
    except Exception as exc:
        return False, '无法解析(%s)' % type(exc).__name__
    if D.list_comments(z):
        return False, '论文(带批注，归 backfill_docx)'
    body = extract_doc_text(path)
    if len(body) > MAX_BODY_AS_OPINION:
        return False, '论文(正文%d字)' % len(body)
    return True, 'ok'


def find_opinion_files(root, skip_dirs=()):
    """找散落在学生文件夹里的意见类文本文件（如 学生D\\修改意见.txt）。

    注意：文件名含"审阅"的**论文** docx（如"审阅XXX初稿.docx"）也会命中关键词，
    因此 docx 必须再经过 docx_opinion_verdict 判定，不能只看文件名。
    """
    out = []
    skip = set(skip_dirs)
    exts = OPINION_TEXT_EXT + ('.docx',) + OPINION_PDF_EXT + OPINION_IMAGE_EXT
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                       if not d.startswith('.') and d not in skip]
        for fn in filenames:
            if fn.startswith('~$'):
                continue
            if not is_opinion_filename(fn):
                continue
            if not fn.lower().endswith(exts):
                continue
            path = os.path.join(dirpath, fn)
            if fn.lower().endswith('.docx'):
                ok, _reason = docx_opinion_verdict(path)
                if not ok:
                    continue
            out.append(path)
    return out


def find_opinion_targets(root, skip_dirs=()):
    """意见目录 + 散落的意见文件。已在意见目录内的文件不重复列出。"""
    dirs = find_opinion_dirs(root, skip_dirs=skip_dirs)
    dirs_norm = [os.path.normcase(os.path.abspath(d)) + os.sep for d in dirs]
    files = []
    for f in find_opinion_files(root, skip_dirs=skip_dirs):
        fa = os.path.normcase(os.path.abspath(f))
        if any(fa.startswith(d) for d in dirs_norm):
            continue
        files.append(f)
    return {'dirs': dirs, 'files': files}
