# docx XML 操作代码模式

所有代码均在本环境验证可用。docx 本质是 zip 包，正文在 word/document.xml，批注在 word/comments.xml。

> **优先用 `scripts/docxlib.py`**：下列模式已固化为可 import 的函数（`D.open_docx`、`D.list_comments`、`D.set_text`、`D.add_text_runs`、`D.new_table`、`D.strip_bold`、`D.insert_section_break`、`D.enable_update_fields`、`D.strip_comments`、`D.replace_media`、`D.save_root`…）。本文件保留为参考实现与原理说明；新脚本请 `import docxlib as D`，不要复制粘贴。

## 公共头

```python
import zipfile, re
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def qn(tag):
    return '{%s}%s' % (W, tag)
def get_text(p):
    return ''.join(t.text or '' for t in p.findall('.//' + qn('t')))
```

## 1. 读取全部批注

```python
zin = zipfile.ZipFile(path)
if 'word/comments.xml' in zin.namelist():
    croot = etree.fromstring(zin.read('word/comments.xml'))
    for c in croot.iter(qn('comment')):
        text = ''.join(t.text or '' for t in c.findall('.//' + qn('t')))
        author = c.get(qn('author'))
        print(author, text[:100])
```

批注锚点：document.xml 中 commentRangeStart/End 包裹被批注内容，commentReference 指向批注 id。

## 2. 查找段落

按文本前缀查找，可跳过目录样式（如 style '15' 目录行、'19' 目录一级）：

```python
def find_para(root, prefix, skip_styles=None, contains=False):
    for p in root.iter(qn('p')):
        txt = get_text(p).strip()
        hit = txt.startswith(prefix) if not contains else prefix in txt
        if hit:
            style = ''
            pPr = p.find(qn('pPr'))
            if pPr is not None:
                ps = pPr.find(qn('pStyle'))
                style = ps.get(qn('val')) if ps is not None else ''
            if skip_styles and style in skip_styles:
                continue
            return p
    return None
```

## 3. 重写段落文本（保留原格式）

清空该段所有 run，按第一个 run 的 rPr 模板重建：

```python
def first_rpr(p):
    for r in p.findall(qn('r')):
        rPr = r.find(qn('rPr'))
        if rPr is not None:
            return rPr
    return None

def set_text(p, text, rpr=None):
    rpr = rpr or first_rpr(p)
    for r in p.findall(qn('r')):
        p.remove(r)
    for child in list(p):
        if child.tag != qn('pPr'):
            p.remove(child)
    add_text_runs(p, text, rpr)
```

### 3a. 更保守的写法：只改 `w:t`，不重建 run（推荐用于"整段换文"）

`set_text()` 会删掉原 run 再按模板重建，遇到**段内有多个不同格式的 run**
（封面「学    号：学生E」= `学` / 四个空格 / `号：学生E` 三段）时，差异会被抹平；
而且如果调用方还假设"段内只有一个 run"，多 run 的段落会直接失败。
换文场景优先用这个：**第一个 `w:t` 写全文，其余 `w:t` 置空**，pPr 与每个 run 的 rPr 原样保留。

```python
def set_para_wt(p, text):
    """整段换文：不改 run 结构，只写第一个 w:t、清空其余 w:t。"""
    ts = []
    for r in p.findall(qn('r')):
        ts.extend(r.findall(qn('t')))
    if not ts:
        return False
    ts[0].text = text
    for rest in ts[1:]:
        rest.text = ''
    return True
```

**保留 XML 空格**：`w:t` 文本首尾带空格时记得 `t.set(qn('xml:space'), 'preserve')`，
否则 Word 会吞掉空格（改封面这类含对齐空格的字段时尤其要当心）。

### 3b. 最保守的写法：只在 `w:t` 内做子串替换（2026-09-30 起，改文字的首选）

`set_text()` / `set_para_wt()` 都是"整段换文"，会**清空段落除 pPr 外的全部子节点**。
真实论文里段落远不止装文字：

- **图片可能与题注同段**：学生F论文 `图5-13 交流主题模块管理` 那一段的子节点是
  `pPr, r(含 w:drawing), r(题注文字), bookmarkStart×4`。整段重写后图片不再被引用，
  但 `word/media/*.png` 还在 zip 里，`inspect_docx.py` 的"图片数"照样 PASS——**完全静默**。
- **段落尾部常挂目录书签**：`bookmarkStart name="_Toc..."`。重写会删掉 start 留下 end，
  书签失衡（本次 142 → 136）。

所以只要改动是"把 A 换成 B"，就用这个：

```python
def replace_in_texts(p, old, new):
    """只在 w:t 文本节点内替换；段内其他节点（图片/书签/域）原样保留。"""
    n = 0
    for t in p.findall('.//' + qn('t')):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            n += 1
    return n
```

确需整段重写时，**先加守卫**，宁可脚本报错也不要静默丢内容：

```python
GUARD = ('drawing', 'pict', 'object', 'bookmarkStart', 'bookmarkEnd',
         'commentRangeStart', 'commentRangeEnd', 'fldChar')

def set_para_text(p, text):
    for tag in GUARD:
        if p.find('.//' + qn(tag)) is not None:
            raise SystemExit('段落内含 w:%s，禁止整段重写: %s' % (tag, get_text(p)[:40]))
    set_text(p, text)
```

**验收必须配结构计数**（`verify_revision.py` 的 `docCounts`）：
`<w:drawing>` / `<w:bookmarkStart>` / `<w:bookmarkEnd>` / `<w:fldChar>` 与基线一致，
并拿改动前的备份跑同一份配置做反向自测（应报 FAIL）。

## 4. 上标引用（关键模式）

### 4a. 生成上标

内容中的引用标记用占位符 `Q = chr(167)`，运行时识别并输出"方括号加编号"的上标文本。
（`Q` 现在是可选的便利写法，不再是"Write 工具会删方括号"的规避手段——见 SKILL.md 陷阱一/二/三。）

```python
Q = chr(167)
SUP_RE = re.compile(Q + '([0-9]+)' + Q)

def append_run(p, text, rpr, superscript):
    import copy
    if not text:
        return
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
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def add_text_runs(p, text, rpr):
    pos = 0
    for m in SUP_RE.finditer(text):
        append_run(p, text[pos:m.start()], rpr, False)
        append_run(p, '[' + m.group(1) + ']', rpr, True)
        pos = m.end()
    append_run(p, text[pos:], rpr, False)
```

**不要**用 `SUP_RE.split(text)` 后对片段 `match`——捕获组是纯数字，匹配会失败。

内容定义示例（脚本文件里）：

```python
CH12_P1 = ('……提出了改进对策' + Q + '1' + Q + '。王岩针对……')
```

### 4b. 读取上标（踩坑修正：必须用 XML 树，不要用字节正则）

早期写法是字节正则
`rb'w:val="superscript"/></w:rPr><w:t xml:space="preserve">(.*?)</w:t>'`。
**这个正则在真实论文上会返回 0 个上标**：Word 只在文本需要保留空格时才写
`xml:space="preserve"`，实际序列化通常是

```xml
<w:vertAlign w:val="superscript"/></w:rPr><w:t>[1]</w:t>
```

正确做法（`docxlib.superscript_runs()`）：

```python
def superscript_runs(root):
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
```

空文本上标 `('', True)` 要单独计数——它是"上标 run 建了但文字丢了"的故障信号。

### 4c. 表格定位（踩坑修正：不要用首行关键字）

数据库设计类表格首行几乎都是「字段名|类型|是否为空|说明」，按首行关键字匹配会**全部命中第一张表**。
用 `docxlib.iter_body_tables()` 按 body 直属子节点顺序拿到 `(表格, 上方最近题注)`，
配置里用 `index` 定位；题注判定为「以 `表+数字` 或 `Table` 开头且长度 ≤ 40」，
避免正文里的「如表4.2所示」被当成题注。

## 5. 插入新段落

```python
def new_para(text, rpr, style=None):
    p = etree.Element(qn('p'))
    pPr = etree.SubElement(p, qn('pPr'))
    if style:
        s = etree.SubElement(pPr, qn('pStyle'))
        s.set(qn('val'), style)
    add_text_runs(p, text, rpr)
    return p

ref_p.addprevious(new_para(text, rpr))  # 插到目标段之前
ref_p.addnext(new_para(text, rpr))      # 插到目标段之后
```

### 5a. 补题注/新段落优先"克隆既有同类段落"（2026-09-30）

新建段落要自己猜字体字号行距，容易与全文不一致。**同一文档里已有同类段落时直接深拷贝再改文字**，
版式逐字节一致：

```python
import copy
proto = find_body_para(root, '表6-2\u3000登录测试用例表', exact=True)   # 既有题注段
tbl = table_after_para(anchor)                                       # 锚点段之后的第一张表
new_p = copy.deepcopy(proto)
replace_in_texts(new_p, '表6-2', '表6-5')                            # 或 set_para_text(new_p, 新题注)
tbl.addprevious(new_p)                                               # 中文体例：表题在表上方
```

要点：题注判定用"以 `表N-M` / `图N-M` 开头"，**表格定位用"锚点正文段之后第一张 `<w:tbl>`"**
（允许中间夹空段落），不要按 index 硬编码——段落索引在插入题注后就会整体位移。

## 6. 新建表格（含边框）

```python
def new_table(rows_data, rpr):
    tbl = etree.Element(qn('tbl'))
    tblPr = etree.SubElement(tbl, qn('tblPr'))
    tblStyle = etree.SubElement(tblPr, qn('tblStyle'))
    tblStyle.set(qn('val'), 'a3')
    tblW = etree.SubElement(tblPr, qn('tblW'))
    tblW.set(qn('w'), '0'); tblW.set(qn('type'), 'auto')
    tblBorders = etree.SubElement(tblPr, qn('tblBorders'))
    for edge in ('top','left','bottom','right','insideH','insideV'):
        b = etree.SubElement(tblBorders, qn(edge))
        b.set(qn('val'), 'single'); b.set(qn('sz'), '4'); b.set(qn('color'), '000000')
    for row_data in rows_data:
        tr = etree.SubElement(tbl, qn('tr'))
        for cell_text in row_data:
            tc = etree.SubElement(tr, qn('tc'))
            tcPr = etree.SubElement(tc, qn('tcPr'))
            tcW = etree.SubElement(tcPr, qn('tcW'))
            tcW.set(qn('w'), '0'); tcW.set(qn('type'), 'auto')
            p = etree.SubElement(tc, qn('p'))
            r = etree.SubElement(p, qn('r'))
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = etree.SubElement(r, qn('t'))
            t.text = cell_text
    return tbl
```

## 7. 两级标题去加粗

遍历含 w:b 的 run，把 b 的 val 置 0：

```python
for b in root.iter(qn('b')):
    if b.get(qn('val')) in ('1', 'true') or b.get(qn('val')) is None:
        b.set(qn('val'), '0')
```

## 8. 分节页码

在第1章标题前插入带 sectPr 的段落（分节符），前缀节 upperRoman 从 1 起，正文节 decimal 从 1 起：

```python
# 前缀节（含目录）：
sectPr = etree.Element(qn('sectPr'))
pg = etree.SubElement(sectPr, qn('pgNumType'))
pg.set(qn('fmt'), 'upperRoman'); pg.set(qn('start'), '1')
# 正文节：
pg.set(qn('fmt'), 'decimal'); pg.set(qn('start'), '1')
```

注意两节各有独立 sectPr（sectPr 数量应为 2）。

## 9. 目录自动更新

settings.xml 根节点下加：

```xml
<w:updateFields w:val="true"/>
```

## 10. 删除批注

- 从 zip 包移除 comments.xml、commentsExtended.xml、people.xml
- 从 word/_rels/document.xml.rels 删除对应 Relationship
- 从 document.xml 删除 commentRangeStart/End、commentReference 元素

## 11. 字体设置（如致谢仿宋）

```python
rPr = etree.SubElement(r, qn('rPr'))
rf = etree.SubElement(rPr, qn('rFonts'))
rf.set(qn('eastAsia'), '仿宋_GB2312')
rf.set(qn('ascii'), 'Times New Roman')
rf.set(qn('hAnsi'), 'Times New Roman')
```

## 11b. 改文字：只动 w:t，不要整段重写（2026-10-04 新增实现）

`D.replace_in_texts(p, old, new)` 已在 `docxlib.py` 里实现（SKILL.md 陷阱十早就点名要它，
但此前只有会清空整段的 `set_text()`）。要点：

```python
ts = [t for t in p.iter(D.qn('t'))]          # 段内所有文本节点，含跨 run
full = ''.join(t.text or '' for t in ts)
i = full.find(old)                            # 跨 run 也能命中
# 把全文展平成"节点区间 -> 字符区间"，替换结果写进第一个命中的节点，
# 其余节点只删掉被匹配的片段（空 run 与全部 rPr 原样保留）
```

三条使用纪律：

1. **old 里不要包含上标标记**。`…实现路径[20]。` 整个当 old 传进去，`[20]` 会被并入普通 run，
   上标格式就丢了。只替换上标之前的正文（`old='…实现路径'`），让 `[20]` 留在原 run。
2. **整段重写前加守卫**：`D.set_text_checked(p, text)` 会在段内含
   drawing/pict/object/bookmarkStart/bookmarkEnd/commentRange/fldChar/instrText 时抛错。
   图片与题注同段、段尾挂 `_Toc` 书签是常见形态（见陷阱十七），整段重写会静默吃掉它们。
3. **补新段落用 `copy.deepcopy(已有同型段落)` 再改文字**，版式逐字节一致；
   `p.addnext(newp)` 插在正文后、空段落前，不要去 `body.append`。

全局术语统一（如 `SpringBoot → Spring Boot`）用 `D.replace_in_texts_global(root, pairs)`：
它对 `root.iter(qn('p'))` 的每个段落套用，返回每条 old 的命中次数，便于与探查结果对账
（本次命中 8 次 == 探查到 8 处）。

## 12. 图片替换
- 图片在 word/media/imageN.png
- 生成新图（image_edit/image_gen），下载为本地文件，zip 方式替换字节：

```python
with open(fig_path, 'rb') as f:
    new_img = f.read()
zin = zipfile.ZipFile(OUT)
zout = zipfile.ZipFile(OUT + '.tmp', 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    if item.filename == 'word/media/image5.png':
        zout.writestr(item, new_img)
    else:
        zout.writestr(item, zin.read(item.filename))
zin.close(); zout.close()
os.replace(OUT + '.tmp', OUT)
```

图替换必须在格式脚本之后执行（格式脚本会重新打包）。

## 13. 打包写回

```python
out_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
zin2 = zipfile.ZipFile(BAK)  # 从备份读其余文件
for item in zin2.infolist():
    if item.filename == 'word/document.xml':
        zout.writestr(item, out_xml)
    else:
        zout.writestr(item, zin2.read(item.filename))
```

## 环境执行注意

- Windows PowerShell：命令用分号连接，不要用 `&&`。
- **不要直接用 `python`**：本机 `python` 不在 PATH 上（只有 WindowsApps Store 别名），且 lxml 只在独立运行时里。统一用 `scripts\run.ps1 <脚本> [参数...]`，它会探测解释器并校验 `import lxml`。
- 脚本以 `-X utf8` 运行（`run.ps1` 已内置）避免中文编码问题。
- `\d`、`\w`、`\n` 与字面 `[1]` 在当前 DSH 工具下均可正常写入（2026-09-29 字节级复核，见 SKILL.md"陷阱一/二/三"）；旧的占位符 `Q = chr(167)` 与 `[0-9]+` 写法保留为可选便利手段，不再是强制规避项。

## 12. 图片加高后同步显示尺寸（2026-10-04 学生H）

给截图顶部补品牌栏会让图片变高。若不同步，Word 按原 `cy` 拉伸会把内容压扁。
同一个 `w:drawing` 里通常有 **两处** 尺寸：`wp:extent` 与 `pic:spPr/a:xfrm/a:ext`，都要改。

```python
A  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
R  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

rels = {rel.get('Id'): rel.get('Target')
        for rel in ET.fromstring(z.read('word/_rels/document.xml.rels'))}

for dr in root.iter(D.qn('drawing')):
    blip = dr.find('.//{%s}blip' % A)
    if blip is None:
        continue
    part = 'word/' + rels[blip.get('{%s}embed' % R)].lstrip('/')
    k = scale.get(part)                      # k = h_new / h_old
    if not k:
        continue
    for ext in dr.iter('{%s}extent' % WP):
        ext.set('cy', str(round(int(ext.get('cy')) * k)))
    for ext in dr.iter('{%s}ext' % A):
        ext.set('cy', str(round(int(ext.get('cy')) * k)))
```

要点：只改 `cy`，`cx` 不动 —— 这样"每像素显示比例"与改前一致，正文不会被压扁；
改完必须重跑 `verify_revision.py`（`docCounts` 与 `media` 两项会同时受图片与尺寸影响）。

## 索引自检 + 段落级 diff（防"改错段"）

按 body 段落序号批量改写时，**序号错一位就是静默覆盖**。两个必备件：

```python
# 1) 改写前断言：段号 -> 该段独有的一小截原文
EXPECT = {555: '后端对订单数据进行分库分表优化',  ...}
for i, mk in EXPECT.items():
    if mk not in D.get_text(rows[i]):
        raise SystemExit('INDEX CHECK FAILED [%d]: %r' % (i, D.get_text(rows[i])[:70]))

# 2) 收尾 diff：改前/改后段落序列做 difflib，逐块人工确认
sm = difflib.SequenceMatcher(a=[t for _, t in base_items],
                             b=[t for _, t in build_items], autojunk=False)
for tag, i1, i2, j1, j2 in sm.get_opcodes():
    if tag != 'equal':
        print(tag, base_items[i1:i2], build_items[j1:j2])
```

## 图片段落居中（判据是 ind/@firstLine）

```python
_PPR_ORDER = ('pStyle', 'keepNext', ..., 'spacing', 'ind', 'contextualSpacing', ...,
              'jc', 'textDirection', ..., 'rPr', 'sectPr', 'pPrChange')

for p in rows:
    if not p.findall('.//' + D.qn('drawing')) or D.get_text(p).strip():
        continue                      # 无图 / 签名图（带文字）
    pPr = D.ensure_pPr(p)
    ind = pPr.find(D.qn('ind'))
    if ind is not None and ind.get(D.qn('firstLine')):
        del ind.attrib[D.qn('firstLine')]      # 关键：清掉首行缩进，否则 Word 里图会右移
    jc = pPr.find(D.qn('jc'))
    if jc is None:
        jc = etree.Element(D.qn('jc'))
        D._place(pPr, jc, _PPR_ORDER)          # 按 schema 顺序插入，不能 append
    jc.set(D.qn('val'), 'center')
```

## 克隆既有段落/表格来新增内容（格式逐字节一致）

```python
def clone_para(template_p, text):
    el = copy.deepcopy(template_p)
    for tag in ('bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd',
                'commentReference', 'fldChar', 'instrText'):
        for e in list(el.iter(D.qn(tag))):
            e.getparent().remove(e)             # 克隆体不要带走书签/域，否则计数失衡
    D.set_text_checked(el, text)
    return el

anchor.addprevious(clone_para(caption_tpl, '表4-9 公告信息表'))
anchor.addprevious(make_table_from_existing(tbl_with_same_rows, data))
```

克隆表前先确认行数一致（`assert len(trs) == len(rows_data)`）；
改单元格文字用 `set_cell(tr, ci, text)`（内部对带书签的单元格退回 `replace_in_texts`）。

## 带 vMerge 的纵向合并表追加分组

组首行 `vMerge w:val="restart"`、组内 `vMerge`（无 val）、整表末行另带
`<w:bottom w:sz="12"/>` 边框 —— 三种模板各克隆一份，末行模板只用于**新的最后一行**。

## 删除段落时保住 _Toc 书签

```python
for el in (p2, p1):                      # 段内可能挂 _TocXXXX 书签
    prev = el.getprevious()
    for tag in ('bookmarkStart', 'bookmarkEnd'):
        for e in list(el.iter(D.qn(tag))):
            e.getparent().remove(e)
            while prev is not None and prev.tag != D.qn('p'):
                prev = prev.getprevious()
            if prev is not None:
                prev.append(e)           # 搬到前一段末尾，保持 bookmark 配平
    el.getparent().remove(el)
```

## 按实测分页回写目录（TOC）缓存页码

目录每个条目的页码是 `PAGEREF` 域的**缓存结果**，是普通 `<w:t>`，域本身不会自动重算
（`updateFields` 只作用于 Word 打开时）。完整实现见 `scripts/fix_toc_pages.py`，
核心状态机如下：

```python
# 在目录段落（pStyle = toc 1/2/3）内，定位 PAGEREF 域的结果节点
for r in p.iter(D.qn('r')):
    ins = r.find(D.qn('instrText'))
    fld = r.find(D.qn('fldChar'))
    if ins is not None and 'PAGEREF' in (ins.text or ''):
        name = re.search(r'PAGEREF\s+(\S+)', ins.text).group(1)   # 书签名
        st = 'ref'
        continue
    if fld is not None:
        ft = fld.get(D.qn('fldCharType'))
        if ft == 'separate' and st == 'ref':
            st = 'num'          # 下一个带 w:t 的 run 就是缓存页码
        elif ft == 'end' and st == 'num':
            st = None
        continue
    if st == 'num':
        ts = r.findall(D.qn('t'))
        if ts:
            ts[0].text = new_page_number      # 只改文字，rPr 原样保留
            st = None
```

配套两个判据缺一不可：

1. **行首自动编号必须先剥掉再等值比较**。标题的“一、”“（一）”“1．”由 `numPr` 生成，
   不落在 XML 文字里（PDF 里却渲染出来了）。用"包含即可"会让
   “图3.17 蓝牙模块实物图”命中标题“蓝牙模块”，页码错到第 3 章。
2. **做完要验证不动点**：回写 → 重新转 PDF → 再跑一遍 `--check`，必须报 0 条待更新。

## 段尾追加文字的正确写法

“不动结构、逐段就地改写”时不要用 `set_text()`（会清空段内图片/书签/域，见陷阱十）。
向段落追加一句话，用 `append_run` 并把 rPr 克隆自**段内最后一个非上标正文 run**：

```python
rpr = None
for r in reversed(p.findall(D.qn('r'))):
    rp = r.find(D.qn('rPr'))
    if rp is None:
        continue
    va = rp.find(D.qn('vertAlign'))
    if va is not None and va.get(D.qn('val')) == 'superscript':
        continue                                  # 不要抄角标的格式
    if ''.join(t.text or '' for t in r.findall(D.qn('t'))).strip():
        rpr = rp
        break
D.append_run(p, text, rpr=rpr or D.first_rpr(p))
```

注意 `CT_P` 的 `EG_PContent` 是 `maxOccurs="unbounded"` 的 choice，所以把 `<w:r>`
追加在段尾（即便段尾有 `bookmarkEnd`）**是 schema 合法的**，只是会延长书签范围，无副作用。

计数的处理见陷阱四十一：`r`/`t` 按"基线 + 追加次数"核对，其余载体计数必须逐项不变。

## 脚注：插入、按 id 重编号、保证渲染不错位（2026-10-06 学生A任务）

脚注由**两个部件**共同定义，任何一边不一致都会静默出错：

* `word/document.xml`：`<w:r><w:footnoteReference w:id="N"/></w:r>`，位置就是**角标落点**；
* `word/footnotes.xml`：`<w:footnote w:id="N">`，里面 `w:footnoteRef` 出角标、后面的 `w:t` 是注释正文。

### 插入一条新脚注引用（落在"实词末尾、标点之前"）

复用 `insert_superscript_citation` 的定位算法，但先落一个临时上标再整条替换成脚注 run
（该函数内部把编号写死成 `[%d]`，直接传脚注 run 不行）：

```python
D.insert_superscript_citation(p, frag, 0)            # 临时标记 [0]
runs = list(p.findall(D.qn('r')))
old = next(r for r in runs
           if ''.join(t.text or '' for t in r.iter(D.qn('t'))) == '[0]')
new = copy.deepcopy(fn_tmpl_run)                     # 段内已有的脚注 run 作模板
for x in new.iter(D.qn('footnoteReference')):
    x.set(D.qn('id'), str(fid))
old.addprevious(new)
p.remove(old)
```

`frag` 必须取自**批注圈定的那句话**的结尾实词（见陷阱四十五）。

### 按"文档引用顺序"重编号（渲染不错位的必要条件）

`w:id` 只是**标识**，页面上显示的 ①②③ 由**同页脚注的出现顺序**决定；
但 LibreOffice/WPS 分配显示编号时读的是 `footnotes.xml` 的**定义顺序**。
一旦"定义顺序 ≠ 文档引用顺序"（例如把新脚注直接 append 到末尾），
新脚注会被整片渲染到**文末最后几页**——而 docx 里引用落点完全正确，只跑 XML 验收查不出来。

```python
# 1) 收集首次引用顺序
old_order = []
for p in body.iter(D.qn('p')):
    for fr in p.iter(D.qn('footnoteReference')):
        if fr.get(D.qn('id')) not in old_order:
            old_order.append(fr.get(D.qn('id')))
# 2) 引用侧改成 0..N-1
NEWID = {old: str(n) for n, old in enumerate(old_order)}
for el in body.iter(D.qn('footnoteReference')):
    el.set(D.qn('id'), NEWID[el.get(D.qn('id'))])
# 3) 定义侧按同一顺序重排；separator/continuationSeparator 先摘出、最后放回
by_old = {f.get(D.qn('id')): f for f in fnroot.iter(D.qn('footnote'))
          if f.get(D.qn('type')) is None}
parent = next(iter(by_old.values())).getparent()
seps = [el for el in list(parent) if el.get(D.qn('type')) is not None]
for el in seps + [el for el in list(parent) if el.get(D.qn('type')) is None]:
    parent.remove(el)                                 # 先全摘掉，避免边遍历边删
for n, old in enumerate(old_order):
    f = by_old[old]
    f.set(D.qn('id'), str(n))
    parent.append(f)
for el in seps:
    parent.append(el)                                 # separator 必须在内容脚注之后
```

三条硬性注意：

1. **`separator` / `continuationSeparator` 不能丢**（用 `w:type` 属性识别），否则 Word 判文件需修复；
2. **遍历中不要边遍历边 `remove`** —— 会死循环（本次 build 卡在后台被强杀），一律 `for el in list(parent)`；
3. 克隆脚注段落时 `copy.deepcopy` 会连 `w14:paraId` 一起复制，**必须重新赋值唯一值**：
   `el.set('{http://schemas.microsoft.com/office/word/2010/wordml}paraId', '%08X' % random.getrandbits(32))`。

按页重新编号是 `sectPr` 上的一行配置（两个元素须在 `w:pgSz` **之前**）：

```xml
<w:footnotePr>
  <w:numFmt w:val="decimalEnclosedCircle"/>
  <w:numRestart w:val="eachPage"/>
</w:footnotePr>
```

核对一律用 `scripts/check_footnotes.py`（id 连续性 + 定义顺序 + 特殊脚注 + 批注锚点对照），
并**转 PDF 确认脚注出现在被引用的那一页**。

## 重建目录（TOC 域）并回填实测页码（2026-10-06 学生A任务）

### 域标记绝不能落在"目录"标题段里

把 `fldChar(begin) / instrText(" TOC \\o \"1-2\" \\h \\z \\u ") / separate` 追加到"目录"标题段
末尾，标题段就成了域的开头 —— Word 更新域时会把**标题本身一起替换掉**。
必须新建一个**独立段落**承载域标记：

```python
toc_field_p = etree.Element(_attr('p'))       # 该段只有域标记 + 一句占位提示
toc_field_p.append(f_run(kind='begin'))
toc_field_p.append(f_run(instr=' TOC \\o "1-2" \\h \\z \\u '))
toc_field_p.append(f_run(kind='separate'))
toc_field_p.append(f_run(text='（按 F9 更新域即按当前排版重建目录）'))
toc_field_p.append(f_run(kind='end'))
toc_head.addnext(toc_field_p)                 # 标题段保持原样，不动
```

### 条目段落：套回 `toc 1`/`toc 2` 样式号

条目名用**普通文本 + `PAGEREF` 域**出页码；**不要**套 `HYPERLINK \l` 域——
实测 LibreOffice 会把条目的显示文字整段丢掉，只渲染出页码。
段落仍要带 `pStyle=13/15`（toc 1 / toc 2），否则 `check_toc.py` 报"目录条目 0"。

### 页码必须重测

条目多了会让目录自身多占一页，正文首页整体后移（本次 25 页 → 26 页，全部条目 +1）。
流程固定为：**先写条目 → 转 PDF 量分页 → 回填 `MEASURED_PAGES` → 再转 PDF 复核**，
最后用 `check_toc_pages.py` 那种"目录页码 vs 正文实测页码"逐条比对（须 0 处不符）。

## 包完整性：新增部件与悬空 Override（2026-10-06 学生A任务）

`docxlib.save_docx()` 只替换**基线 zip 里已有的**部件；新键不会进包、也不报错。
一份 safe 的写法是覆盖它，在写完基线条目后补写剩下的新部件：

```python
_orig = D.save_docx
def save_docx_newparts(out_path, base_path, replacements, remove=(), drop_comments=False):
    tmp = out_path + '.tmp2'
    written = set()
    with zipfile.ZipFile(base_path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in remove:
                continue
            data = replacements.get(item.filename)
            zout.writestr(item, data if data is not None else zin.read(item.filename))
            written.add(item.filename)
        for name, data in replacements.items():       # 补写新部件
            if name not in written and name not in remove and not name.endswith('/'):
                zout.writestr(name, data)
    os.replace(tmp, out_path)
D.save_docx = save_docx_newparts
```

`[Content_Types].xml` 也要同步：**新加**的部件要有 Override（同名旧 Override 先删避免重复），
**删掉**的部件（批注/people 等）的 Override 必须一并删除，否则留下悬空 Override，Word 判需修复。

一行自查：

```python
parts = re.findall(r'<Override PartName="([^"]+)"', ct)
assert len(parts) == len(set(parts))                                    # 无重复
assert not [p for p in parts if p.lstrip('/') not in z.namelist()]      # 无悬空
```

## 表格单元格内重建段落序列（2026-10-06 学生M任务 · 开题报告）

开题报告/任务书的正文全装在 `body/tbl[0]` 的 10 行里，**按段落序号定位必然错位**。
正确姿势：按行 → 单元格 → 保留前 N 段 → 用原单元格里的段落当模板重建其余段落。
要点是"模板段落 `deepcopy` 之后再删原段落"，这样字体/字号/行距/缩进逐项继承。

```python
from lxml import etree
import copy, zipfile, os
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_NS = '{http://www.w3.org/XML/1998/namespace}'
def qn(t): return W + t

root = etree.fromstring(zipfile.ZipFile(src).read('word/document.xml'))
tbl  = [e for e in root.find(qn('body')) if e.tag == qn('tbl')][0]
rows = tbl.findall(qn('tr'))
# 用"行内文本包含小节标签"定位，比行号稳
row  = [r for r in rows if '研究的主要内容和思路' in
        ''.join(t.text or '' for t in r.iter(qn('t')))][0]
cell = row.findall(qn('tc'))[0]
ps   = cell.findall(qn('p'))

indent_tmpl = copy.deepcopy(ps[1])   # 有 <w:ind firstLineChars="200"> 的缩进段
flush_tmpl  = copy.deepcopy(ps[0])   # 无 <w:ind> 的顶格段
anchor = ps[0]
for p in ps[1:]:
    cell.remove(p)

def make(tmpl, text):
    p = copy.deepcopy(tmpl)
    for ch in list(p):                       # 清空除 pPr 外的全部子节点
        if ch.tag != qn('pPr'):
            p.remove(ch)
    r = etree.SubElement(p, qn('r'))
    rPr = tmpl.find(qn('r')).find(qn('rPr'))
    if rPr is not None:
        r.append(copy.deepcopy(rPr))
    t = etree.SubElement(r, qn('t'))
    t.text = text
    t.set(XML_NS + 'space', 'preserve')      # 首尾空格必须保
    return p

for line in spec:                            # spec: [(indent, text), ...]
    anchor.addnext(make(indent_tmpl if line[0] else flush_tmpl, line[1]))
    anchor = anchor.getnext()
```

重打包时**逐部件保留原始压缩方式与属性**，否则 WPS 写的"近乎不压缩的大图"会被重压
（体积骤降，见 SKILL.md 陷阱四十四）：

```python
with zipfile.ZipFile(src) as z, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
    for item in z.infolist():
        data = out_xml if item.filename == 'word/document.xml' else z.read(item.filename)
        zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
        zi.compress_type = item.compress_type
        zi.external_attr = item.external_attr
        zo.writestr(zi, data)
```

配套断言（少一条就可能静默改错或静默无改动）：

* `--find` 命中的**行数必须恰好 1**；
* `--replace` 的每处替换**命中次数必须恰好 1**（0 次=定位串写错，>1 次=误伤）；
* 重建后 `cell` 内每个 `w:r` 的父节点必须是 `w:p`（无嵌套 run）；
* 改前/改后逐段比对：`A[:k] == B[:k]`、尾部逐段相等，且 `tr/tc/drawing/bookmarkStart/fldChar` 不变。

通用实现见 `scripts/rebuild_cell.py`，配套探测器 `scripts/dump_table_row.py`。

## 三线表：`none` 边线必须显式写（2026-10-07 学生K任务）

论文里的数据表要的是**三线表**（顶线、表头下线、底线），而 `docxlib.new_table()`
给六条边都加单线，是全网格表。新增的 `docxlib.new_three_line_table()` 按下面的模式生成：

```python
# tblPr：居中 + 固定布局 + 只保留上下粗线，其余显式 none
tblPr
  tblW(w=0, type=auto)
  jc(val=center)
  tblBorders
    top    (val=single, sz=12, color=000000)
    bottom (val=single, sz=12, color=000000)
    left/right/insideH/insideV (val=none, sz=0, color=auto)   # ← 三处都不能省
  tblLayout(type=fixed)

# 表头行：trPr/tblHeader（跨页重复）+ 每个 tc 的 tcBorders/bottom 加一条细线
tr[0]/trPr/tblHeader
tr[0]/tc/tcPr/tcBorders/bottom(val=single, sz=6, color=000000)

# 每个单元格：tcW(type=dxa, w=<该列宽度>) + vAlign(center) + 段落 jc
```

要点：

* `left/right/insideH/insideV` **只写 `val="none"` 不够**，必须同时给 `sz=0`/`color=auto`，
  否则 Word 回退到默认框线，"三线表"变成"全网格表"；
* 列宽用 `tcW type="dxa"`，各列之和 ≤ 正文节可用宽度
  （`pgSz/@w − pgMar/@left − pgMar/@right`；本次 `11907−1134×2=9639`，取 9400）；
* 题注段在表格**上方**、资料来源段在表格**下方**，都用普通 body 段落承载：
  `anchor.addnext(cap); cap.addnext(tbl); tbl.addnext(src)`；
* 正文必须有"见表 N-M"的引用，否则 `check_captions.py` 报"表有题注但正文未引用"；
* 单元格字号按学校规范（五号 = `sz=21`），标签过长会把列撑成"一个字一行"，
  渲染出来看一眼再定列宽。

插入后会改变 `docCounts`：`<w:tbl` 会从 0 变大（注意它是**子串计数**，
一个表格会命中 `<w:tbl` `w:tblPr` `w:tblW` `w:tblBorders` `w:tblLayout` `w:tblHeader` 六次）。
所以 `docCounts` 里**不要放 `<w:tbl`**，表格存在性交给 `expectTables` 按题注核对。

## 已有网格表改成三线表：必须同时清掉行级与单元格级边框（2026-10-07 学生D任务）

`new_three_line_table()` 只能用于**新建**表。若要在**原有表格**上去网格线（降查重时最省力的一层
差异化），只改 `tblPr/tblBorders` 是没用的 —— 渲染出来还是网格。原因是有两处覆盖：

```python
# 1) 每一行都可能带 <w:tblPrEx>，里面是一整份 tblBorders，直接盖过表级设置
#    → 必须整块删掉（不是改值）
for ex in tbl.findall(q('tblPrEx')):
    tblPr.remove(ex)
for tr in tbl.findall(q('tr')):
    ex = tr.find(q('tblPrEx'))
    if ex is not None:
        tr.remove(ex)
    # 2) 单元格级 tcBorders 也会盖过表级设置 → 整块删掉
    for tc in tr.findall(q('tc')):
        tcPr = tc.find(q('tcPr'))
        tcb = tcPr.find(q('tcBorders')) if tcPr is not None else None
        if tcb is not None:
            tcPr.remove(tcb)

# 表级：顶/底粗线，其余四边显式 none（含 sz=0/color=auto）
tblBorders/top    (val=single, sz=12, color=000000)
tblBorders/bottom (val=single, sz=12, color=000000)
tblBorders/left|right|insideH|insideV (val=none, sz=0, color=auto)

# 表头行：tblHeader + tcBorders/bottom(sz=6) + shd(fill=F2F2F2)
```

**插入位置必须按 schema 顺序**，否则 Word 提示"需修复"。用顺序插入器而不是 `append`：

```python
TCPR_ORDER = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd',
              'noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark']
TRPR_ORDER = ['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter',
              'cantSplit','trHeight','tblHeader','tblCellSpacing','jc','hidden']

def inserter(parent, child, order):
    idx = order.index(child.tag.split('}')[-1])
    for i, el in enumerate(parent):
        name = el.tag.split('}')[-1]
        if name in order and order.index(name) > idx:
            parent.insert(i, child); return child
    parent.append(child); return child
```

这一项**抓到过真 bug**：`tblHeader` 被 `insert(0, ...)` 放到 `trHeight` 之前，
而 schema 要求 `trHeight` 在前。判据是遍历所有 `tblPr/trPr/tcPr/pPr/rPr` 看子元素序号是否递增：

```python
seq = [c.tag.split('}')[-1] for c in el]
known = [x for x in seq if x in ORDER[name]]
assert [ORDER[name].index(x) for x in known] == sorted(ORDER[name].index(x) for x in known)
```

## 拆掉 vMerge：跨页时合并标签会错位（2026-10-07 学生D任务）

首列做纵向合并的分组表（"用户注册"跨 3 行），一旦跨页，**LibreOffice 把合并标签整体推到
分页点**，渲染成"第一行空、第二行显示上一组名"，而 docx 里 `vMerge` 的 restart/continue
完全正确 —— 只查 XML 发现不了，必须 `render` 出来看。

```python
# 拆掉 col 列的 vMerge（普通单元格的 vMerge 直接删除即可）
for tr in tbl.findall(q('tr')):
    tcPr = tr.findall(q('tc'))[col].find(q('tcPr'))
    vm = tcPr.find(q('vMerge'))
    if vm is not None:
        tcPr.remove(vm)
# 然后把组名写进每一行（表格逐行自解释，复制粘贴也不丢组名）
```

顺带两条排版判据：

* `cantSplit`（`trPr/cantSplit`，禁止行内断页）能让"表头 + 半行"不再孤零零留在上一页，
  但**会让总页数 +1**（本次 44 → 45）；需要页数不变就别加；
* 题注段带 `keepNext` 只保证"题注 + 表格第一行"同页，**不保证第一行不与表体分离**，
  所以两者要一起用。核对方式：`convert` 成 PDF 后逐页看"题注所在页是否就是表格首页"。

重写表格的内容侧做法（换列结构 / 换表头 / 换字段 / 换说明，一次动四层）见
`scripts/rewrite_tables.py` 与 SKILL.md 陷阱五十八～六十。
