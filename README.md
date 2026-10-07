# paper-revision —— 按批注修改毕业论文的 DSH 技能

一套把"导师批注 → 改动后的 docx"这条链路工程化的技能：**配置驱动、逐项可验收、每次任务沉淀经验**。
原本是给一个真实论文代改场景写的，现在整理成可以直接换论文复用的公开版本。

## 它解决什么问题

毕业论文修改的痛点是"改完说不清改对没有"。本技能把每个判定都写成可执行检查：

- 批注逐条落到对应段落（先打印 `commentRangeStart..commentRangeEnd` 原文，再改）；
- 参考文献全部替换为近 5 年真实可查文献，并逐条核验真实性（学位论文 `[D]` 是编造重灾区）；
- 格式（两级标题去加粗、分节页码、目录更新、字体、题注、三线表）按学校规范文档里的**真实数值**对齐；
- 交付前跑 `verify_revision.py`（退出码必须 0）、结构合法性校验、图表编号与目录双向核对、
  脚注落点核对、转 PDF 目视复核，并对**改前备份**做反向自测（必须 FAIL，证明检查有区分度）。

## 目录结构

```
SKILL.md                     技能正文（工作流 + 60 条实战陷阱 + 验收清单）
scripts/                     参数化脚本（docx 操作库、探查、验收、核对、渲染分页、表格重写）
  docxlib.py                 通用 docx XML 操作库（无硬编码）
  inspect_docx.py            基线探查，并生成 task.json 草稿
  verify_revision.py         配置驱动验收（退出码 0/1）
  check_captions.py          图表编号与引用双向核对
  check_toc.py               目录与正文标题双向核对
  check_footnotes.py         脚注落点、id 连续性与定义顺序核对
  check_balance.py           交付前自查（章节字数均衡 / 直引号 / 上标与文献编号）
  fix_toc_pages.py           按实测分页回写目录缓存页码（含 --check 不动点校验）
  rewrite_tables.py          按 JSON 重写表格（降查重：三线表、拆 vMerge、改列结构）
  screenshot_edit/           系统截图批量精修工具箱（含 Windows OCR 调用）
  run.ps1                    解释器解析与执行包装
references/                  方法论文档（进化日志、XML 代码模式、文献核验方法、归纳统计）
corpus/taxonomy.json         批注分类规则表（数据，可增补）
tests/fixtures/              采集层的最小测试样本
```

## 安装

```bash
git clone https://github.com/tjc123yy/paper-revision-skill.git ~/.dsh/skills/paper-revision
```

DSH 会扫描 `<技能根>/<名称>/SKILL.md`，放好后即可被技能目录识别（也可用目录联接指向克隆位置）。

## 快速开始

```powershell
cd <技能目录>\scripts
# 1. 探查：提取批注、样式号、参考文献区间、表格、字体、图片，生成 task.json 草稿
.\run.ps1 inspect_docx.py "论文.docx"
# 2. 把 suggestedConfig 存为 task.json 并按需修订
# 3. 逐项验收，退出码必须为 0
.\run.ps1 verify_revision.py "论文_已修改.docx" --config task.json
```

`run.ps1` 会自动探测带 `lxml` 的 Python 解释器（脚本不依赖 `python` 在 PATH 上），
可用环境变量 `DSH_PAPER_PYTHON` 强制指定。依赖：`lxml`、`python-docx`、`Pillow`，
渲染核对另需 LibreOffice；`pypdf` / `pypdfium2` 可选。

## 关于脱敏

本仓库**不含**任何真实论文、批注原文、学生姓名或学号：

- 真实学生姓名统一替换为 `学生A`…`学生R`，学号与论文题目隐去；
- 语料记录（`corpus/records/`）、专业映射（`corpus/major_map.json`）、
  语料来源（`corpus/sources.json`）与任务脚本（`tasks/`）均留在作者本地，不发布；
- `scripts/screenshot_edit/` 里示例账号为占位值，非真实账号。

## 许可

MIT，见 [LICENSE](LICENSE)。技能正文里的方法与判据可以自由取用；请自行删除其中的个人信息后再分发。
