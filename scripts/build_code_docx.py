#!/usr/bin/env python3
"""
按公司格式生成源代码鉴别材料 Word 文档。

用法:
  uv run --with python-docx python3 build_code_docx.py \
      --root <项目根目录> --files 草稿/代码文件顺序.txt \
      --name "XX系统" --version V1.0 -o 正式资料/

代码文件顺序.txt: 一行一个相对路径，按抽取顺序排列（第 1 个必须是入口文件）。以 # 开头的行忽略。

规则（内置，见 references/格式要求.md）:
  - 去空行、去纯注释行（--keep-comments 可保留注释）；Tab 转 4 空格
  - 每页严格 50 行，每页前强制分页；页眉左"全称+版本号"，右"第 N 页"
  - 总页 > 60：连续前 30 页 + 连续后 30 页，页码 1–60；≤ 60 页：全部
  - 末页不足 50 行：页底标注"本页为源代码最后一页，共 XX 行"
  - A4，上 1.8 / 下 1.6 / 左右 2.4 cm，宋体 9pt，单倍行距
输出:
  <全称>_软件代码.docx 、 代码文档索引.md（每页对应的文件与原始行号，供人工核对）
"""
import argparse, re, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_LINE_SPACING
from docx_common import set_font, set_style_font, setup_header, setup_page, page_break_before, remove_doc_grid

LINES_PER_PAGE = 50
FRONT_PAGES = 30
BACK_PAGES = 30
BLOCK_COMMENT = {"/*": "*/", "<!--": "-->", '"""': '"""', "'''": "'''"}

def clean_lines(text: str, keep_comments: bool):
    """返回 [(原始行号, 文本)]"""
    out = []; in_block = None
    for i, raw in enumerate(text.splitlines(), 1):
        line = raw.replace("\t", "    ").rstrip()
        s = line.strip()
        if not s: continue
        if not keep_comments:
            if in_block:
                if in_block in s: in_block = None
                continue
            if s.startswith(("//", "#", "*", "--")) and not s.startswith("#!"): continue
            opened = False
            for o, c in BLOCK_COMMENT.items():
                if s.startswith(o):
                    if c not in s[len(o):]: in_block = c
                    opened = True; break
            if opened: continue
        out.append((i, line))
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--files", required=True)
    ap.add_argument("--name", required=True, help="软件全称")
    ap.add_argument("--version", required=True, help="版本号，如 V1.0")
    ap.add_argument("-o", "--outdir", default=".")
    ap.add_argument("--keep-comments", action="store_true")
    ap.add_argument("--font", default="宋体")
    ap.add_argument("--font-size", type=float, default=9)
    a = ap.parse_args()

    root = Path(a.root).resolve()
    paths = [l.strip() for l in Path(a.files).read_text(encoding="utf-8").splitlines() if l.strip() and not l.startswith("#")]
    if not paths:
        sys.exit("文件清单为空")

    # 拼接所有代码行: (path, lineno, text)
    all_lines = []
    per_file = []
    for rel in paths:
        p = root / rel
        if not p.exists():
            sys.exit(f"文件不存在: {p}")
        lines = clean_lines(p.read_text(encoding="utf-8", errors="replace"), a.keep_comments)
        per_file.append((rel, len(lines)))
        all_lines.extend((rel, n, t) for n, t in lines)
    total = len(all_lines)
    pages = [all_lines[i:i + LINES_PER_PAGE] for i in range(0, total, LINES_PER_PAGE)]
    n_pages = len(pages)

    if n_pages > FRONT_PAGES + BACK_PAGES:
        selected = pages[:FRONT_PAGES] + pages[-BACK_PAGES:]
        mode = f"前 {FRONT_PAGES} 页 + 后 {BACK_PAGES} 页（原文共 {n_pages} 页，中间删去 {n_pages - 60} 页）"
    else:
        selected = pages
        mode = f"全部 {n_pages} 页（不足 60 页）"

    header_text = f"{a.name}{a.version}"
    doc = Document()
    set_style_font(doc.styles["Normal"], a.font, a.font_size)
    pf = doc.styles["Normal"].paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(a.font_size + 2)
    sec = doc.sections[0]
    setup_page(sec, 1.8, 1.6, 2.4, 2.4)
    remove_doc_grid(sec)
    setup_header(sec, header_text, font_size=a.font_size, font=a.font)

    index = ["# 代码文档索引", "", f"软件：{header_text}", f"抽取方式：{mode}", f"有效代码总行数：{total}", "",
             "## 文件顺序与有效行数", "", "| 序 | 文件 | 有效行 |", "|---|---|---:|"]
    index += [f"| {i+1} | `{rel}` | {n} |" for i, (rel, n) in enumerate(per_file)]
    index += ["", "## 页 → 文件/原始行号", "", "| 页码 | 起 | 止 |", "|---|---|---|"]

    for pi, page in enumerate(selected, 1):
        first = True
        for rel, n, text in page:
            p = doc.add_paragraph()
            if first and pi > 1:
                page_break_before(p)
            first = False
            r = p.add_run(text); set_font(r, a.font, a.font_size)
        s_rel, s_n, _ = page[0]; e_rel, e_n, _ = page[-1]
        index.append(f"| {pi} | `{s_rel}`:{s_n} | `{e_rel}`:{e_n} |")
        if pi == len(selected) and len(page) < LINES_PER_PAGE:
            p = doc.add_paragraph()
            r = p.add_run(f"本页为源代码最后一页，共 {len(page)} 行；以下无代码。"); set_font(r, a.font, a.font_size, bold=True)

    outdir = Path(a.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / f"{a.name}_软件代码.docx"
    doc.save(out)
    (outdir / "代码文档索引.md").write_text("\n".join(index) + "\n", encoding="utf-8")

    print(f"✅ {out}")
    print(f"✅ {outdir / '代码文档索引.md'}")
    print(f"有效行数 {total}，{mode}，输出 {len(selected)} 页")
    print(f"首页起始：{selected[0][0][0]}:{selected[0][0][1]}  → 请确认是程序入口")
    print(f"末页结束：{selected[-1][-1][0]}:{selected[-1][-1][1]}  → 请确认是完整模块收尾")
    if total < 60 * LINES_PER_PAGE:
        print(f"ℹ️  不足 3000 行，按\"不足 60 页交全部\"处理")
    if n_pages > 60:
        cut = pages[FRONT_PAGES - 1][-1]; res = pages[-BACK_PAGES][0]
        print(f"中间断点：第 30 页止于 {cut[0]}:{cut[1]}，第 31 页起于 {res[0]}:{res[1]}")

if __name__ == "__main__":
    main()
