# -*- coding: utf-8 -*-
"""docx 通用操作库 —— 论文修改技能（参数化，无硬编码）

依赖：lxml（保留命名空间前缀；不要用标准库 ElementTree 写回，会重命名前缀并破坏 mc:Ignorable）。
解释器解析见同目录 run.ps1（本环境 python 不在 PATH 上）。

约定：
    Q = chr(167) 是"方括号加数字"引用标记的占位符。
    正文内容里写：'……提出了改进对策' + Q + '1' + Q + '。'
    运行时由 add_text_runs() 输出为带方括号编号的上标 run。

本库不含任何具体论文、学生、章节、样式号或关键字的硬编码。
"""
import copy
import hashlib
import os
import re
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

DOCUMENT = 'word/document.xml'
SETTINGS = 'word/settings.xml'
STYLES = 'word/styles.xml'
COMMENTS = 'word/comments.xml'

Q = chr(167)
SUP_RE = re.compile(Q + '([0-9]+)' + Q)

COMMENT_TAGS = ('commentRangeStart', 'commentRangeEnd', 'commentReference')


def qn(tag):
    """限定到 w: 命名空间的标签名。"""
    return '{%s}%s' % (W, tag)


# --------------------------------------------------------------------------
# 读
# --------------------------------------------------------------------------

def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def open_docx(path):
    """返回 (ZipFile, document.xml 的 lxml 根, document.xml 原始字节)。"""
    z = zipfile.ZipFile(path)
    data = z.read(DOCUMENT)
    root = etree.fromstring(data)
    return z, root, data


def get_text(p):
    """段落可见文本。"""
    return ''.join(t.text or '' for t in p.findall('.//' + qn('t')))


def para_style(p):
    """段落的 pStyle 值（无则为空串）。"""
    pPr = p.find(qn('pPr'))
    if pPr is None:
        return ''
    ps = pPr.find(qn('pStyle'))
    return ps.get(qn('val')) if ps is not None else ''


def iter_paras(root):
    return root.iter(qn('p'))


def find_para(root, prefix, skip_styles=None, contains=False):
    """按文本前缀（或包含）查找段落，可跳过指定样式（如目录样式）。"""
    for p in iter_paras(root):
        txt = get_text(p).strip()
        hit = (prefix in txt) if contains else txt.startswith(prefix)
        if not hit:
            continue
        if skip_styles and para_style(p) in skip_styles:
            continue
        return p
    return None


def find_para_exact(root, text, skip_styles=None):
    """按完整文本精确查找段落。"""
    for p in iter_paras(root):
        if get_text(p).strip() != text:
            continue
        if skip_styles and para_style(p) in skip_styles:
            continue
        return p
    return None


def read_styles(z):
    """返回 {styleId: 样式名} 映射，用于反查标题样式号。"""
    if STYLES not in z.namelist():
        return {}
    sroot = etree.fromstring(z.read(STYLES))
    out = {}
    for st in sroot.iter(qn('style')):
        sid = st.get(qn('styleId'))
        nm = st.find(qn('name'))
        out[sid] = nm.get(qn('val')) if nm is not None else ''
    return out


def list_comments(z):
    """返回批注列表 [{id, author, date, text}]（无 comments.xml 则为空）。"""
    if COMMENTS not in z.namelist():
        return []
    croot = etree.fromstring(z.read(COMMENTS))
    out = []
    for c in croot.iter(qn('comment')):
        out.append({
            'id': c.get(qn('id')),
            'author': c.get(qn('author')),
            'date': c.get(qn('date')),
            'text': get_text(c),
        })
    return out


def count_comment_markers(doc_bytes):
    """document.xml 字节中批注标记的出现次数。"""
    return {tag: doc_bytes.count(tag.encode('utf-8')) for tag in COMMENT_TAGS}


def superscript_runs(root):
    """按文档顺序返回 [(文本, 是否空)]，只含 vertAlign=superscript 的 run。

    不要用字节正则匹配上标：Word 只在需要时写 xml:space="preserve"，
    实际序列化常见为 <w:vertAlign w:val="superscript"/></w:rPr><w:t>[1]</w:t>，
    任何固定属性串的正则都会漏配。用 XML 树判定才是可靠的。
    """
    out = []
    for r in root.iter(qn('r')):
        rPr = r.find(qn('rPr'))
        if rPr is None:
            continue
        va = rPr.find(qn('vertAlign'))
        if va is None or va.get(qn('val')) != 'superscript':
            continue
        txt = ''.join(t.text or '' for t in r.findall(qn('t')))
        out.append((txt, txt == ''))
    return out


def superscript_numbers(doc):
    """返回 (上标编号列表, 空文本上标数)。

    doc 可以是 document.xml 的字节，也可以是已解析的根节点。
    """
    root = etree.fromstring(doc) if isinstance(doc, (bytes, bytearray)) else doc
    nums = []
    empty = 0
    for txt, is_empty in superscript_runs(root):
        if is_empty:
            empty += 1
            continue
        digits = re.findall('[0-9]+', txt)
        if digits:
            nums.append(int(digits[0]))
    return nums, empty


# --------------------------------------------------------------------------
# 写：run 与段落
# --------------------------------------------------------------------------

def first_rpr(p):
    """段内第一个带 rPr 的 run 的 rPr，作为格式模板。"""
    for r in p.findall(qn('r')):
        rPr = r.find(qn('rPr'))
        if rPr is not None:
            return rPr
    return None


def append_run(p, text, rpr=None, superscript=False):
    """向段落追加一个 run；superscript=True 时加上标。"""
    if not text:
        return None
    r = etree.SubElement(p, qn('r'))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if superscript:
        rPr = r.find(qn('rPr'))
        if rPr is None:
            rPr = etree.SubElement(r, qn('rPr'))
        va = etree.SubElement(rPr, qn('vertAlign'))
        va.set(qn('val'), 'superscript')
    t = etree.SubElement(r, qn('t'))
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return r


def add_text_runs(p, text, rpr=None):
    """把含 Q 占位符的文本写入段落：占位符处生成方括号编号的上标 run。

    注意：必须用 finditer 交替写入；不要用 SUP_RE.split() 后对片段 match
    （捕获组是纯数字，匹配会失败）。
    """
    pos = 0
    for m in SUP_RE.finditer(text):
        append_run(p, text[pos:m.start()], rpr, False)
        append_run(p, '[' + m.group(1) + ']', rpr, True)
        pos = m.end()
    append_run(p, text[pos:], rpr, False)


# 允许紧跟在上标之后的字符：句末/句内标点，或同一句里的下一个角标。
SUP_FOLLOW_OK = '。，；：！？、[]'
# 角标之前**不允许**出现的字符（出现即说明角标落到了标点之后）。
SUP_PRECEDE_BAD = '。，；：！？、'


def _nonempty_texts(p):
    """按文档顺序返回段落内非空 w:t 的 [(节点, 相对整段文本起点)]。

    空 w:t 必须排除：它占列表位置但不占文本位置，直接在 list(iter(t)) 上累加长度
    会让后续所有偏移错位（实测踩中，角标被插到段首第一个字后面）。
    """
    out = []
    pos = 0
    for t in p.iter(qn('t')):
        txt = t.text or ''
        if not txt:
            continue
        out.append((t, pos))
        pos += len(txt)
    return out


def _split_t(t, local_off):
    """把 w:t 在 local_off 处拆开：原节点留前半，克隆节点放后半并插在其后。返回后半。"""
    text = t.text or ''
    if not (0 < local_off < len(text)):
        raise ValueError('拆分位置越界: local_off=%d len=%d' % (local_off, len(text)))
    t.text = text[:local_off]
    t.set(XML_SPACE, 'preserve')
    new_t = copy.deepcopy(t)
    new_t.text = text[local_off:]
    new_t.set(XML_SPACE, 'preserve')
    t.addnext(new_t)
    return new_t


def _run_level_anchor(node):
    """把插入锚点提升到"段落的直属子元素"层级。

    **这是踩过大坑的一步**：<w:t> 的父节点往往就是 <w:r> 本身，直接在 w:t 旁边
    addprevious/addnext 会把上标 <w:r> 插进正文 <w:r> **内部**，生成
    `<w:r>…<w:t>…</w:t><w:r><w:vertAlign .../></w:r><w:t>。</w:t></w:r>`
    这种 schema 非法结构。LibreOffice 能渲染（看起来对），但 Word/WPS 会判定文件
    需修复，并把角标重新排到标点之后 —— 本次用户看到的"没改过来"就是这个。
    正确做法：沿祖先链上溯到"父节点是 <w:p> 的那一层"，上标作为它的兄弟插入。
    """
    cur = node
    while cur is not None:
        parent = cur.getparent()
        if parent is None:
            break
        if parent.tag == qn('p'):
            return cur
        cur = parent
    return node


def _set_run_text(run, text):
    """把 run 的文本整体设为 text：清掉原有 w:t，在首个 w:t 位置写回。"""
    ts = run.findall(qn('t'))
    if ts:
        keep = ts[0]
        for t in ts[1:]:
            run.remove(t)
        keep.text = text
        keep.set(XML_SPACE, 'preserve')
        return keep
    t = etree.SubElement(run, qn('t'))
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return t


def _split_run_for_anchor(run, anchor_el):
    """把 run 拆成"插入点之前"和"插入点及之后"两条 run，返回后一条。

    上标随后插在两条 run 之间，因此它始终是**段落的直属子元素**。

    做法是把 run 的文本合并后按 anchor 的起点切开，两条 run 各写一段文本，
    而不是搬动 w:t 子节点 —— 搬节点会让 w:t 在文档顺序里错位（实测会读成
    "。[1]前段文字…"这种文本重复/乱序）。
    """
    all_texts = [t for t in run.findall(qn('t'))]
    if anchor_el not in all_texts:
        raise ValueError('anchor w:t 不属于该 run')
    head = ''.join(t.text or '' for t in all_texts)
    cut = 0
    for t in all_texts:
        if t is anchor_el:
            break
        cut += len(t.text or '')
    if cut >= len(head):
        raise ValueError('拆 run 位置越界: cut=%d 文本长=%d' % (cut, len(head)))
    clone = copy.deepcopy(run)
    run.addnext(clone)
    _set_run_text(run, head[:cut])
    _set_run_text(clone, head[cut:])
    return clone


def insert_superscript_citation(p, anchor_text, num, rpr=None,
                               east_asia=None, ascii_font=None, hansi_font=None):
    """在段落 p 内把上标 [num] 插到 anchor_text **之后、其后的标点之前**。

    规范（论文排版硬要求）：角标必须在句末标点之前。
        规范：……转向主动预防的必然选择[1]。
        不规范：……转向主动预防的必然选择。[1]
    因此传入的 anchor_text 应是标点之前的那段实词（如 '必然选择'），
    本函数把插入点定在 anchor_text 的**末尾**，并在返回前断言其后紧跟标点。

    实现要点（都是踩过的坑）：
      * 定位角标不能用 find('')、也不能在含空 w:t 的列表上按长度累加——两者都会错位。
        这里用 _nonempty_texts() 建立"节点 -> 段落内偏移"，按 offset ∈ [pos, pos+len) 命中。
      * 在 w:t 内部拆开后，上标必须插在拆出的**后半之前**（addprevious）。
        写成后半之后（addnext）会让角标越过标点落到句末，正是被用户退回的那个 bug。
      * 插入点 = anchor 起始 offset + len(anchor)，不是 anchor 起始 offset
        （写成起始会让角标插进词中间："必然[1]选择。"）。

    rpr 为 None 时取该 run 自身的 rPr（并复制）；east_asia/ascii_font/hansi_font
    非空时顺带显式设置角标字体。返回段落 p。定位或规范校验失败时抛 ValueError。
    """
    nodes = _nonempty_texts(p)
    full = ''.join(t.text or '' for t, _ in nodes)
    if anchor_text not in full:
        raise ValueError('段落内未找到定位片段 %r（段落=%.30s）' % (anchor_text, full))
    start = full.index(anchor_text)
    at = start + len(anchor_text)              # 插入点：实词末尾，标点之前
    if at >= len(full):
        raise ValueError(
            'anchor_text %r 已到段落末尾，之后没有标点可"插在其前"，'
            '或 anchor_text 本身把标点也包含进去了。\n'
            '请传标点之前的最后一个实词（如 anchor_text="必然选择"，而不是'
            '"必然选择。"）。' % anchor_text)
    follow = full[at:at + 1]
    if follow not in SUP_FOLLOW_OK:
        raise ValueError(
            '角标位置不符合规范: %r 之后紧跟 %r，不是标点。\n'
            '角标必须放在句末标点之前，请把 anchor_text 取成标点前的最后一个实词。'
            % (anchor_text, follow))

    hit_t = None
    local = 0
    for t, pos in nodes:
        length = len(t.text or '')
        if pos <= at < pos + length:
            hit_t, local = t, at - pos
            break
    if hit_t is None:
        raise ValueError('定位失败: anchor=%r at=%d 段长=%d' % (anchor_text, at, len(full)))
    hit_r = hit_t.getparent()
    run_texts = [t for t in hit_r.findall(qn('t'))]
    if local == 0:
        anchor_el = hit_t                     # 插入点在此 w:t 之前
        prev_texts = []
        for t in run_texts:
            if t is anchor_el:
                break
            prev_texts.append(t)
    else:
        anchor_el = _split_t(hit_t, local)    # 拆开后，插入点在此后半 w:t 之前
        prev_texts = [t for t in run_texts if t is not anchor_el]

    base_rpr = rpr if rpr is not None else hit_r.find(qn('rPr'))
    if base_rpr is None:
        base_rpr = etree.Element(qn('rPr'))
        hit_r.insert(0, base_rpr)
    sup_rpr = copy.deepcopy(base_rpr)
    if east_asia or ascii_font or hansi_font:
        set_east_asia_font(sup_rpr, east_asia or '宋体',
                           ascii_font=ascii_font, hansi_font=hansi_font)
    va = etree.SubElement(sup_rpr, qn('vertAlign'))
    va.set(qn('val'), 'superscript')
    sup = etree.Element(qn('r'))
    sup.append(sup_rpr)
    t = etree.SubElement(sup, qn('t'))
    t.text = '[%d]' % num
    t.set(XML_SPACE, 'preserve')
    # 关键：上标必须是"段落的直属子元素"，放在两条 run 之间，
    # 绝不能嵌进正文 <w:r> 内部（schema 非法 -> Word/WPS 会修复并把角标挪到标点后）。
    if prev_texts:
        # 本 run 内含多个 w:t（或已在 w:t 内部拆过），要先把 run 在插入点处拆成两条
        tail_run = _split_run_for_anchor(hit_r, anchor_el)
        tail_run.addprevious(sup)
    else:
        hit_r.addprevious(sup)
    return p


def check_nested_runs(root):
    """检查非法嵌套：<w:r> 里面不能有 <w:r>（会破坏上标位置，见 _place_sup_before_run）。"""
    bad = 0
    for r in root.iter(qn('r')):
        for child in r:
            if child.tag == qn('r'):
                bad += 1
    return bad


def check_citation_placement(root, skip_contains=('[J]',)):
    """全局核对每个上标的落点与结构，返回问题列表（空列表 = 全部规范）。

    判据（三条都要满足）：
        1. 上标 <w:r> 必须是段落的直属子元素 —— 不能嵌在别的 <w:r> 里。
           嵌套时 LibreOffice 仍能渲染成"看起来正确"，但 Word/WPS 会判文件需修复
           并把角标挪到标点之后（本次交付被用户退回就是这个原因）。
        2. 角标之后第一个字符 ∈ SUP_FOLLOW_OK（标点或同一句的下一个角标）
        3. 角标之前第一个字符 ∉ SUP_PRECEDE_BAD（否则角标落在标点之后）
    实现必须按"展平的 (类型, 文本) 序列 + 下标"取前后字符。**不要用 find('')**：
    空串 find 恒返回 0，会让检查恒真（本次真实踩中，构建脚本报了 PASS 却是错的）。
    """
    problems = []
    n_nested = check_nested_runs(root)
    if n_nested:
        problems.append('存在 %d 处非法嵌套 <w:r>（上标 run 被嵌进正文 run 内部）' % n_nested)
    for r in root.iter(qn('r')):
        rPr = r.find(qn('rPr'))
        va = rPr.find(qn('vertAlign')) if rPr is not None else None
        if va is not None and va.get(qn('val')) == 'superscript':
            if r.getparent() is not None and r.getparent().tag not in (qn('p'), qn('hyperlink')):
                problems.append('上标 run 的父节点是 <%s>，不是段落直属（结构非法）'
                                % r.getparent().tag.split('}')[1])
    for p in root.iter(qn('p')):
        ptxt = get_text(p)
        if any(s in ptxt for s in skip_contains):
            continue
        seq = []
        for node in p.iter():
            if node.tag == qn('t') and (node.text or ''):
                par = node.getparent()
                prpr = par.find(qn('rPr')) if par is not None else None
                pva = prpr.find(qn('vertAlign')) if prpr is not None else None
                if pva is not None and pva.get(qn('val')) == 'superscript':
                    continue                      # 角标自带 w:t，单独收集
                seq.append(('t', node.text))
            elif node.tag == qn('r'):
                rpr = node.find(qn('rPr'))
                va = rpr.find(qn('vertAlign')) if rpr is not None else None
                if va is not None and va.get(qn('val')) == 'superscript':
                    seq.append(('sup', ''.join(x.text or '' for x in node.findall(qn('t')))))
        for i, (kind, txt) in enumerate(seq):
            if kind != 'sup' or not txt:
                continue
            prev_ch = ''.join(t for k, t in seq[:i] if k == 't')[-1:]
            nxt_ch = ''
            for k, t in seq[i + 1:]:
                if k == 'sup' and t:
                    nxt_ch = '['
                    break
                if k == 't' and t:
                    nxt_ch = t[0]
                    break
            if nxt_ch not in SUP_FOLLOW_OK:
                problems.append('角标 %s 之后紧跟 %r（应为标点），段=%.20s'
                                % (txt, nxt_ch, ptxt))
            if prev_ch in SUP_PRECEDE_BAD:
                problems.append('角标 %s 之前是标点 %r，角标落在标点之后，段=%.20s'
                                % (txt, prev_ch, ptxt))
    return problems


def set_text(p, text, rpr=None):
    """清空段落所有 run，按模板 rPr 重建文本（保留 pPr）。"""
    rpr = rpr if rpr is not None else first_rpr(p)
    for r in p.findall(qn('r')):
        p.remove(r)
    for child in list(p):
        if child.tag != qn('pPr'):
            p.remove(child)
    add_text_runs(p, text, rpr)


def replace_in_texts(p, old, new, limit=0):
    """在段落内把 old 替换为 new，**只改 w:t 文本节点**，不动任何 run/图片/书签。

    这是"改文字"的首选手段（见 SKILL.md 陷阱十）：`set_text()` 会清空段落除 pPr 外的
    全部子节点，段内的图片、目录书签、域都会一起消失，而且 media 部件仍在 zip 里，
    常规验收（图片数）根本发现不了。本函数只在文本节点上做拼接，结构逐字节不变。

    ⚠️ **`old` 跨多个 w:t 节点时有严重副作用，调用前必须保证 `old` 在段内唯一**
    （2026-10-05 学生J任务实测，见 references/evolution-log.md）：匹配跨多个节点时，
    命中区间内的**第一个节点会被写成整段新文本**，区间内**其余节点被整体清空**
    （不是"只删被匹配的片段"）。后果：跨"正文 run + 上标 run + 正文 run"的替换会把
    上标压成普通文本，最小复现见 `_task/学生J/bug_repro.py`。

    安全用法（缺一不可）：

        before = ''.join(t.text or '' for t in p.iter(qn('t')))
        assert before.count(old) == 1          # 1) 唯一 -> 循环只走一遍
        assert replace_in_texts(p, old, new) == 1
        # 2) 校验文本恰好是"只替换了这一处"
        assert ''.join(t.text or '' for t in p.iter(qn('t'))) == before.replace(old, new, 1)

    3) 需要"整段文本 -> 整段文本"的重写（引号规范化、占位符转编号等）**不要**用它，
    改为逐个 w:t 改 `t.text`；跨节点的结构性改动请走 `set_text*` 或自建 build_runs。

    limit>0 时最多替换 limit 次；返回实际替换次数。
    """
    ts = [t for t in p.iter(qn('t'))]
    if not ts:
        return 0
    full = ''.join(t.text or '' for t in ts)
    if old not in full:
        return 0
    total = 0
    start = 0
    while True:
        i = full.find(old, start)
        if i < 0 or (limit and total >= limit):
            break
        j = i + len(old)
        spans = []
        pos = 0
        for t in ts:
            txt = t.text or ''
            spans.append((pos, pos + len(txt)))
            pos += len(txt)
        for (s, e), t in zip(spans, ts):
            if e <= i or s >= j:
                continue
            txt = t.text or ''
            ls = max(i, s) - s
            le = min(j, e) - s
            if s <= i < e:
                t.text = txt[:ls] + new + txt[le:]
                if t.text != t.text.strip():
                    t.set(XML_SPACE, 'preserve')
            else:
                t.text = txt[:ls] + txt[le:]
        full = full[:i] + new + full[j:]
        total += 1
        start = i + len(new)
    return total


def replace_in_texts_global(root, pairs):
    """对文档内全部段落（含表格、文本框）按顺序套用 [(old, new), ...]。

    返回 {old: 替换次数}。用于"SpringBoot→Spring Boot"这类全局术语统一。
    """
    counts = {o: 0 for o, _ in pairs}
    for p in root.iter(qn('p')):
        for old, new in pairs:
            counts[old] += replace_in_texts(p, old, new)
    return counts


def has_protected_content(p, extra=()):
    """段落内是否含图片/书签/域/批注等不可重建的节点（set_text 前的守卫）。"""
    tags = ('drawing', 'pict', 'object', 'bookmarkStart', 'bookmarkEnd',
            'commentRangeStart', 'commentRangeEnd', 'commentReference',
            'fldChar', 'instrText') + tuple(extra)
    for t in tags:
        if p.find('.//' + qn(t)) is not None:
            return t
    return None


def set_text_checked(p, text, rpr=None, extra=()):
    """带守卫的 set_text：段内含受保护节点时抛错，绝不静默丢内容。"""
    hit = has_protected_content(p, extra)
    if hit:
        raise ValueError('段落含受保护节点 <%s>，拒绝整段重写' % hit)
    set_text(p, text, rpr)
    return p


def new_para(text, rpr=None, style=None):
    """新建段落（可选 pStyle）。"""
    p = etree.Element(qn('p'))
    pPr = etree.SubElement(p, qn('pPr'))
    if style:
        s = etree.SubElement(pPr, qn('pStyle'))
        s.set(qn('val'), style)
    add_text_runs(p, text, rpr)
    return p


def new_table(rows_data, rpr=None, style=None):
    """按二维列表新建带边框表格。style 为 None 时不套表格样式。"""
    tbl = etree.Element(qn('tbl'))
    tblPr = etree.SubElement(tbl, qn('tblPr'))
    if style:
        ts = etree.SubElement(tblPr, qn('tblStyle'))
        ts.set(qn('val'), style)
    tw = etree.SubElement(tblPr, qn('tblW'))
    tw.set(qn('w'), '0')
    tw.set(qn('type'), 'auto')
    borders = etree.SubElement(tblPr, qn('tblBorders'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = etree.SubElement(borders, qn(edge))
        b.set(qn('val'), 'single')
        b.set(qn('sz'), '4')
        b.set(qn('color'), '000000')
    for row_data in rows_data:
        tr = etree.SubElement(tbl, qn('tr'))
        for cell_text in row_data:
            tc = etree.SubElement(tr, qn('tc'))
            tcPr = etree.SubElement(tc, qn('tcPr'))
            cw = etree.SubElement(tcPr, qn('tcW'))
            cw.set(qn('w'), '0')
            cw.set(qn('type'), 'auto')
            p = etree.SubElement(tc, qn('p'))
            append_run(p, cell_text, rpr)
    return tbl


def new_three_line_table(rows_data, widths=None, header=True, sz=21,
                         east_asia='宋体', ascii_font='Times New Roman',
                         line='300', center_cols=None):
    """新建**三线表**（学术论文规范表）：顶线/底线加粗，表头下一条细线，其余无框线。

    为什么要单独一个函数（2026-10-07 学生K任务）：
    `new_table()` 给六条边都加单线，得到的是"全网格表"，不符合论文三线表规范；
    而三线表的坑在于 tblBorders 与 tcBorders 要分别设置、且 `w:val="none"` 的边
    也必须显式写出 sz/color，否则 Word 会回退到默认框线。

    参数：
        rows_data:   二维列表，rows_data[0] 为表头
        widths:      每列宽度（twips 字符串），None 时按列数均分页面可用宽度
        header:      是否有表头行（有则在首行加细底线并设 tblHeader 重复标题行）
        sz:          单元格字号（半磅）；五号 = 21
        center_cols: 需要居中的列号集合，默认表头与所有列都居中；
                     传 (0, 1) 表示只有前两列居中、其余左对齐
    返回 <w:tbl> 元素（尚未挂到文档树上，由调用方 addnext 插入）。
    """
    ncol = len(rows_data[0])
    if widths is None:
        widths = [str(9400 // ncol)] * ncol
    if center_cols is None:
        center_cols = set(range(ncol))

    tbl = etree.Element(qn('tbl'))
    tblPr = etree.SubElement(tbl, qn('tblPr'))
    tw = etree.SubElement(tblPr, qn('tblW'))
    tw.set(qn('w'), '0')
    tw.set(qn('type'), 'auto')
    jc = etree.SubElement(tblPr, qn('jc'))
    jc.set(qn('val'), 'center')
    borders = etree.SubElement(tblPr, qn('tblBorders'))
    for edge in ('top', 'bottom'):
        b = etree.SubElement(borders, qn(edge))
        b.set(qn('val'), 'single')
        b.set(qn('sz'), '12')          # 粗线
        b.set(qn('color'), '000000')
    for edge in ('left', 'right', 'insideH', 'insideV'):
        b = etree.SubElement(borders, qn(edge))
        b.set(qn('val'), 'none')       # 必须显式 none，否则 Word 回退默认框线
        b.set(qn('sz'), '0')
        b.set(qn('color'), 'auto')
    layout = etree.SubElement(tblPr, qn('tblLayout'))
    layout.set(qn('type'), 'fixed')

    for ri, row_data in enumerate(rows_data):
        tr = etree.SubElement(tbl, qn('tr'))
        trPr = etree.SubElement(tr, qn('trPr'))
        if header and ri == 0:
            etree.SubElement(trPr, qn('tblHeader'))
        for ci, cell_text in enumerate(row_data):
            tc = etree.SubElement(tr, qn('tc'))
            tcPr = etree.SubElement(tc, qn('tcPr'))
            cw = etree.SubElement(tcPr, qn('tcW'))
            cw.set(qn('w'), widths[ci])
            cw.set(qn('type'), 'dxa')
            if header and ri == 0:
                tcb = etree.SubElement(tcPr, qn('tcBorders'))
                bb = etree.SubElement(tcb, qn('bottom'))
                bb.set(qn('val'), 'single')
                bb.set(qn('sz'), '6')  # 表头下细线
                bb.set(qn('color'), '000000')
            va = etree.SubElement(tcPr, qn('vAlign'))
            va.set(qn('val'), 'center')
            p = etree.Element(qn('p'))
            pPr = etree.SubElement(p, qn('pPr'))
            sp = etree.SubElement(pPr, qn('spacing'))
            sp.set(qn('line'), line)
            sp.set(qn('lineRule'), 'auto')
            j = etree.SubElement(pPr, qn('jc'))
            j.set(qn('val'), 'center' if ci in center_cols else 'left')
            rPr = etree.Element(qn('rPr'))
            rf = etree.SubElement(rPr, qn('rFonts'))
            rf.set(qn('hint'), 'eastAsia')
            rf.set(qn('ascii'), ascii_font)
            rf.set(qn('hAnsi'), ascii_font)
            rf.set(qn('eastAsia'), east_asia)
            s = etree.SubElement(rPr, qn('sz'))
            s.set(qn('val'), str(sz))
            scs = etree.SubElement(rPr, qn('szCs'))
            scs.set(qn('val'), str(sz))
            append_run(p, cell_text, rPr)
            tc.append(p)
    return tbl


# --------------------------------------------------------------------------
# 写：整篇级处理
# --------------------------------------------------------------------------

def strip_bold(root):
    """两级标题去加粗：把所有生效的 w:b 置 0。返回处理数量。"""
    n = 0
    for b in root.iter(qn('b')):
        if b.get(qn('val')) in ('1', 'true') or b.get(qn('val')) is None:
            b.set(qn('val'), '0')
            n += 1
    return n


def set_east_asia_font(rpr, east_asia, ascii_font=None, hansi_font=None):
    """在给定 rPr 上设置中文字体（其余可选）。"""
    rf = rpr.find(qn('rFonts'))
    if rf is None:
        rf = etree.SubElement(rpr, qn('rFonts'))
    rf.set(qn('eastAsia'), east_asia)
    if ascii_font:
        rf.set(qn('ascii'), ascii_font)
    if hansi_font:
        rf.set(qn('hAnsi'), hansi_font)
    return rf


def insert_section_break(root, anchor_prefix, fmts, skip_styles=None):
    """在 anchor_prefix 段之前插入分节符，并为各节设置页码格式。

    fmts: [{'fmt': 'upperRoman', 'start': '1'}, {'fmt': 'decimal', 'start': '1'}]
          第一个用于插入点之前的节，第二个用于之后的节。
    返回插入的段落，找不到锚点时返回 None。
    """
    anchor = find_para(root, anchor_prefix, skip_styles=skip_styles)
    if anchor is None:
        return None
    sect = etree.Element(qn('p'))
    pPr = etree.SubElement(sect, qn('pPr'))
    sectPr = etree.SubElement(pPr, qn('sectPr'))
    pg = etree.SubElement(sectPr, qn('pgNumType'))
    pg.set(qn('fmt'), fmts[0]['fmt'])
    pg.set(qn('start'), str(fmts[0]['start']))
    anchor.addprevious(sect)

    body = root.find(qn('body'))
    if body is not None:
        tail = body.find(qn('sectPr'))
        if tail is not None:
            tpg = tail.find(qn('pgNumType'))
            if tpg is None:
                tpg = etree.SubElement(tail, qn('pgNumType'))
            tpg.set(qn('fmt'), fmts[1]['fmt'])
            tpg.set(qn('start'), str(fmts[1]['start']))
    return sect


def enable_update_fields(settings_bytes):
    """在 settings.xml 根节点下加 updateFields，使 Word 打开时刷新目录。

    返回 (新字节, 是否新增)。已存在时不重复添加。
    """
    root = etree.fromstring(settings_bytes)
    for el in root.iter(qn('updateFields')):
        el.set(qn('val'), 'true')
        return etree.tostring(
            root, xml_declaration=True, encoding='UTF-8', standalone=True
        ), False
    el = etree.Element(qn('updateFields'))
    el.set(qn('val'), 'true')
    root.insert(0, el)
    return etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    ), True


_RPR_ORDER = ('rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
              'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
              'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
              'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
              'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
              'eastAsianLayout', 'specVanish', 'oMath')


def _place(parent, el, order):
    """把 el 插入 parent，位置满足 order 给出的子元素顺序。

    rPr 的子元素顺序在 OOXML 里是 schema 规定的：顺序错了 Word 会拒绝打开
    整个文件。追加到末尾在多数情况下会踩这个坑，所以按 order 找插入点。
    """
    try:
        want = order.index(etree.QName(el).localname)
    except ValueError:
        parent.append(el)
        return el
    for i, child in enumerate(parent):
        try:
            ci = order.index(etree.QName(child).localname)
        except ValueError:
            continue
        if ci > want:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def get_or_add(parent, tag, order=None):
    """取子元素，没有就建。order 给定时按 schema 顺序插入。"""
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = etree.Element(qn(tag))
    if order is None:
        parent.append(el)
    else:
        _place(parent, el, order)
    return el


def ensure_pPr(p):
    return get_or_add(p, 'pPr')


def ensure_rPr(r):
    return get_or_add(r, 'rPr', _RPR_ORDER)


def count_in_doc_children(root, tag):
    """body 直属子元素中 tag 的个数（不含表格/文本框内部）。"""
    body = root.find(qn('body'))
    if body is None:
        return 0
    return sum(1 for c in body if c.tag == qn(tag))


def para_index_map(root):
    """body 直属 p 元素 -> 序号（与遍历顺序一致）。"""
    body = root.find(qn('body'))
    out = {}
    i = 0
    for c in body:
        if c.tag == qn('p'):
            out[c] = i
            i += 1
    return out


def strip_comments(root, rels_bytes=None):
    """从 document.xml 删除全部批注标记；可选从 rels 删除对应关系。

    返回 (删除的标记数, 新的 rels 字节或 None)。
    """
    removed = 0
    for tag in COMMENT_TAGS:
        for el in list(root.iter(qn(tag))):
            el.getparent().remove(el)
            removed += 1
    new_rels = None
    if rels_bytes is not None:
        rroot = etree.fromstring(rels_bytes)
        for rel in list(rroot):
            if 'comments' in (rel.get('Type') or '') or 'people' in (
                rel.get('Target') or ''
            ):
                rroot.remove(rel)
        new_rels = etree.tostring(
            rroot, xml_declaration=True, encoding='UTF-8', standalone=True
        )
    return removed, new_rels


# --------------------------------------------------------------------------
# 打包
# --------------------------------------------------------------------------

COMMENT_PARTS = (
    'word/comments.xml',
    'word/commentsExtended.xml',
    'word/commentsIds.xml',
    'word/commentsExtensible.xml',
    'word/people.xml',
)

IMAGE_EXT = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff',
             '.emf', '.wmf', '.svg')


def save_docx(out_path, base_path, replacements, remove=(), drop_comments=False):
    """以 base_path 为底，替换/删除若干部件，写出 out_path（原子替换）。

    replacements: {部件名: bytes}
    remove:       要删除的部件名集合
    drop_comments: 顺带删除全部批注相关部件
    """
    remove = set(remove)
    if drop_comments:
        remove.update(COMMENT_PARTS)
    tmp = out_path + '.tmp'
    with zipfile.ZipFile(base_path) as zin, zipfile.ZipFile(
        tmp, 'w', zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            name = item.filename
            if name in remove:
                continue
            data = replacements.get(name)
            zout.writestr(item, data if data is not None else zin.read(name))
    os.replace(tmp, out_path)


def save_root(out_path, base_path, root, extra=None, remove=(), drop_comments=False):
    """把修改后的 lxml 根写回 word/document.xml 并打包。"""
    xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    repl = {DOCUMENT: xml}
    if extra:
        repl.update(extra)
    save_docx(out_path, base_path, repl, remove=remove, drop_comments=drop_comments)


def replace_media(out_path, src_path, media_map):
    """按 {部件名: bytes} 替换图片字节（在格式脚本之后执行）。"""
    save_docx(out_path, src_path, media_map)


def list_media(z):
    """列出 word/media 下的图片部件及字节数。

    排除目录条目（zip 里存在名为 'word/media/' 的 0 字节条目）与非图片文件。
    """
    out = []
    for n in z.namelist():
        if not n.startswith('word/media/') or n.endswith('/'):
            continue
        if not n.lower().endswith(IMAGE_EXT):
            continue
        out.append({'name': n, 'bytes': z.getinfo(n).file_size})
    return sorted(out, key=lambda d: d['name'])


def looks_like_caption(txt):
    """判断段落是否像表题注：以 表+编号 或 Table 开头，且足够短。

    正文里的"如表4.2所示"不以"表"开头，因此不会被误判为题注。
    """
    if not txt or len(txt) > 40:
        return False
    if re.match('表[ \t]*[0-9]', txt):
        return True
    return txt.lower().startswith('table')


def iter_body_tables(root):
    """按文档顺序产出 body 直属的 (表格, 上方最近题注) 对。

    inspect 与 verify 共用同一套定位逻辑，保证索引一致。
    """
    body = root.find(qn('body'))
    if body is None:
        return
    recent = None
    for child in body:
        if child.tag == qn('p'):
            txt = get_text(child).strip()
            if looks_like_caption(txt):
                recent = txt
        elif child.tag == qn('tbl'):
            yield child, recent
            recent = None
