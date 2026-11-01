#!/usr/bin/env python3
"""
统一提取职规赛参赛材料文本（带页码/段落锚点，供 PPT 大纲与证据位取材）。

用法:
    python3 extract_content.py <输入文件> [输出.txt]

支持格式与锚点规范:
    .pptx            → "===== 幻灯片 N ====="（含表格、组合形状、备注）——锚点写作 P N
    .pdf             → "===== 第 N 页 ====="（pdftotext -layout）——PDF 报告/简历/辅助证明通用
    .docx            → 标题行原样保留，正文段落前缀 "PN"，表格按行输出——锚点写作 P N
    .md/.markdown/.txt → 按空行分块，块首行前缀 "PN"——锚点写作 P N

依赖:
    pptx:  pip3 install python-pptx
    docx:  pip3 install python-docx
    pdf:   系统 pdftotext（brew install poppler）
"""
import os
import re
import shutil
import subprocess
import sys


def extract_pptx(path):
    """幻灯片逐页提取；有备注页时附 [备注]，常含真实细节。"""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def iter_text(shape):
        texts = []
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in shape.shapes:
                texts.extend(iter_text(sub))
            return texts
        if getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                texts.append(" | ".join(c.text.strip() for c in row.cells))
            return texts
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                line = "".join(r.text for r in para.runs).strip()
                if line:
                    texts.append(line)
        return texts

    prs = Presentation(path)
    out = []
    for idx, slide in enumerate(prs.slides, 1):
        out.append(f"\n===== 幻灯片 {idx} =====")
        for shape in slide.shapes:
            try:
                out.extend(iter_text(shape))
            except Exception as e:  # 单个形状失败不中断整体提取
                print(f"[slide {idx} shape skipped: {e}]", file=sys.stderr)
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                out.append(f"[备注] {notes}")
    return "\n".join(out)


def extract_pdf(path):
    if not shutil.which("pdftotext"):
        sys.exit("缺少 pdftotext，请先安装: brew install poppler")
    raw = subprocess.run(
        ["pdftotext", "-layout", path, "-"],
        capture_output=True, text=True, check=True,
    ).stdout
    out = []
    for idx, page in enumerate(raw.split("\f"), 1):
        page = page.strip()
        if not page:
            continue
        out.append(f"\n===== 第 {idx} 页 =====")
        # 折叠页内多余空行，减少噪音
        out.append(re.sub(r"\n{3,}", "\n\n", page))
    return "\n".join(out)


def extract_docx(path):
    import docx
    from docx.document import Document as DocClass
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def iter_blocks(parent):
        for child in parent.element.body.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent)
            elif child.tag == qn("w:tbl"):
                yield Table(child, parent)

    doc = docx.Document(path)
    out, para_no, tbl_no = [], 0, 0
    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            if block.style.name.lower().startswith("heading"):
                out.append(f"\n■ {text}")  # 标题作为天然章节锚点
            else:
                para_no += 1
                out.append(f"P{para_no} {text}")
        elif isinstance(block, Table):
            tbl_no += 1
            out.append(f"（表格{tbl_no}）")
            for row in block.rows:
                out.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(out)


def extract_text(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        content = f.read()
    out, block_no = [], 0
    for block in re.split(r"\n\s*\n", content):
        block = block.strip("\n")
        if not block.strip():
            continue
        block_no += 1
        lines = block.splitlines()
        lines[0] = f"P{block_no} {lines[0]}"
        out.append("\n".join(lines))
    return "\n\n".join(out)


EXTRACTORS = {
    ".pptx": extract_pptx,
    ".pdf": extract_pdf,
    ".docx": extract_docx,
    ".md": extract_text,
    ".markdown": extract_text,
    ".txt": extract_text,
}


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    path = sys.argv[1]
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXTRACTORS:
        sys.exit(f"不支持的格式 {ext}；支持: {', '.join(EXTRACTORS)}")
    text = EXTRACTORS[ext](path)
    header = f"# 提取自: {os.path.basename(path)}（格式 {ext}）\n"
    result = header + text
    if len(sys.argv) >= 3:
        with open(sys.argv[2], "w", encoding="utf-8") as f:
            f.write(result)
        print(f"OK → {sys.argv[2]}（{len(result)} 字符）", file=sys.stderr)
    else:
        print(result)


if __name__ == "__main__":
    main()
