# ruanzhu-skill · 软著材料生成 Skill

给 AI Agent（Claude Code / Codex / Cursor / OpenCode 等）用的技能包：读取真实项目，按中国版权保护中心的交存习惯和一套经过多次过审/补正打磨的内部格式规范，生成软件著作权登记的三份材料：

| 材料 | 输出 | 内置规则 |
|---|---|---|
| 申请表 | 填好的 Word 申请表 + 逐项对照的 txt | 21 个字段；主要功能 500–1300 字校验；技术特点 ≤100 字 |
| 操作手册 | Word（封面、目录、页眉、右上角页码、截图+图注） | 除目录 ≥20 页，每页 ≥30 行，截图含标题栏、不打码 |
| 源代码鉴别材料 | Word（每页严格 50 行，连续前 30 + 后 30 页，末页标注） | 首页入口、尾页完整收尾；去空行/注释；页眉"全称+版本号" |

外加一致性自检：三处名称/版本/权利人一字不差、截图占位残留、代码页行数、语言与申请表匹配、手册是否夹带代码。

> 个人在版权中心官网申请通常免费。本 skill 只整理材料，不代替权属判断、原创性审查，不保证过审。提交前请自行核对名称、版本、著作权人、日期和代码是否真实一致。官方入口：中国版权保护中心著作权登记系统，咨询电话 010-61090099。

## 安装

依赖：Python 3.10+、[uv](https://docs.astral.sh/uv/)（脚本用 `uv run --with` 自动装 python-docx / pyyaml）、LibreOffice（`soffice`，仅转 PDF 用）。

```bash
git clone https://github.com/pkm365/ruanzhu-skill.git
```

然后把目录放到（或软链到）你的 Agent 的 skills 目录：

| Agent | 目录 |
|---|---|
| Claude Code | `~/.claude/skills/ruanzhu-materials` |
| Codex | `~/.codex/skills/ruanzhu-materials` |
| Cursor / OpenCode 等 | 各自的 skills 目录，或直接把 `SKILL.md` 内容贴进项目规则 |

```bash
ln -s "$(pwd)/ruanzhu-skill" ~/.claude/skills/ruanzhu-materials
```

## 使用

在项目里对 Agent 说：

> 为当前项目生成软件著作权申请资料

Agent 会按 `SKILL.md` 的 6 步走，在 4 个节点停下等你确认（登记事实 → 业务口径 → 代码文件顺序 → 手册内容与截图），然后生成 Word 并自检。产物在项目根目录 `软件著作权申请资料/`。

也可以不经 Agent 直接跑脚本，见 `SKILL.md` 第 5 步的命令，或 `examples/设备点检系统/README.md` 试跑虚构示例。

## 目录

```
SKILL.md                      Agent 读的工作流与硬规则
references/
  格式要求.md                 版权中心交存格式 + 实操防打回要点（2026.8）
  申请表字段说明.md           21 个字段怎么填，主要功能五段骨架
  操作手册写作规范.md         结构、篇幅、截图、措辞
templates/申请表模板.docx     空白申请表（脚本按单元格填）
examples/设备点检系统/        虚构示例：申请表 yaml、手册 md、代码顺序 txt
scripts/
  collect_sources.py          扫项目 → 候选文件 + 真实源程序量 + 语言
  build_code_docx.py          代码 Word（50 行/页、前30后30、页眉页码、索引）
  build_manual_docx.py        手册 Markdown → Word
  fill_application_form.py    yaml → 申请表 docx + txt
  check_consistency.py        提交前自检
  docx_to_pdf.sh              批量转 PDF
```

## 已知限制

- 手册目录是 Word 域，需用 Word 打开一次更新域再导出 PDF（LibreOffice 不会算页码）；或加 `--static-toc` 出无页码目录。
- 手册 Markdown 只支持标题、段落、编号行、管道表、`![图注](路径)` 图片、`【截图预留】` 占位。
- 代码文档前后 30 页的中间断点在文件中部是正常的，但请核对脚本打印的首页起始与末页结束。
- 格式规则以中国版权保护中心官网当期说明为准；本仓库规则整理于 2026 年 8 月。

## License

MIT
