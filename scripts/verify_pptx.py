#!/usr/bin/env python3
"""zgs-ppt / verify_pptx.py —— 渲染层复核：把 pptx 转成 PDF，用真实文字坐标查越界、压字、过小字号。

为什么不只信 build_pptx.py 的估算：
估算基于"字宽≈字号"的近似，遇到数字与字母混排、全角标点、字体替换时会偏乐观。
PDF 里的 bbox 是排版器算完的真实位置，能在交稿前把"看不出来但会被评委看到"的问题捞出来。

用法：
    python3 verify_pptx.py 展示.pptx
    python3 verify_pptx.py 展示.pptx --expect-pages 10 --margin 20
    python3 verify_pptx.py 展示.pptx --json 报告.json

依赖：soffice（LibreOffice）、pdftotext（poppler）。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

MARGIN_PT = 20.0      # 安全边距（pt），约 0.28 英寸
MIN_FONT_PT = 9.0     # 低于此高度视为过小
OVERLAP_RATIO = 0.26  # 两词框交叠超过较小框的该比例，判为压字
MIN_OVERLAP_PT = 3.5  # 两个方向的交叠都不小于此值才判（避免描边级误报）


def to_pdf(pptx: str) -> tuple[str | None, str]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None, "未找到 soffice，无法渲染复核（只能依赖 build_pptx.py 的估算）"
    outdir = os.path.dirname(os.path.abspath(pptx))
    pdf = os.path.splitext(os.path.abspath(pptx))[0] + ".pdf"
    proc = subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                           "--outdir", outdir, os.path.abspath(pptx)],
                          capture_output=True, text=True)
    if not os.path.exists(pdf):
        return None, f"PDF 转换失败：{proc.stderr.strip()[:200]}"
    return pdf, ""


def word_boxes(pdf: str):
    pdftotext = shutil.which("pdftotext")
    if not pdftotext:
        raise SystemExit("未找到 pdftotext（pip/brew 安装 poppler）")
    xml = subprocess.run([pdftotext, "-bbox", "-q", pdf, "-"],
                         capture_output=True, text=True).stdout
    root = ET.fromstring(xml)

    def tag(el):
        return el.tag.split("}")[-1]

    pages = []
    for page in [el for el in root.iter() if tag(el) == "page"]:
        w, h = float(page.get("width")), float(page.get("height"))
        words = []
        for el in [x for x in page.iter() if tag(x) == "word"]:
            txt = (el.text or "").strip()
            if not txt:
                continue
            words.append({
                "text": txt,
                "x0": float(el.get("xMin")), "y0": float(el.get("yMin")),
                "x1": float(el.get("xMax")), "y1": float(el.get("yMax")),
            })
        pages.append({"w": w, "h": h, "words": words})
    return pages


def area(b) -> float:
    return max(0.0, b["x1"] - b["x0"]) * max(0.0, b["y1"] - b["y0"])


def inter(a, b) -> float:
    dx = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
    dy = min(a["y1"], b["y1"]) - max(a["y0"], b["y0"])
    if dx < MIN_OVERLAP_PT or dy < MIN_OVERLAP_PT:
        return 0.0
    return dx * dy


def check_page(page, margin: float):
    words = page["words"]
    out = {"words": len(words), "outside": [], "overlaps": [], "tiny": []}
    for wd in words:
        if (wd["x0"] < margin or wd["y0"] < margin
                or wd["x1"] > page["w"] - margin or wd["y1"] > page["h"] - margin):
            out["outside"].append(wd["text"][:20])
        # 页脚（组别、页码）本来就小，不算"过小字号"
        if (wd["y1"] - wd["y0"]) < MIN_FONT_PT and wd["y0"] < page["h"] - 26:
            out["tiny"].append(wd["text"][:20])
    # 压字：先按 y 分桶，减少无谓两两比较
    buckets: dict[int, list] = {}
    for wd in words:
        buckets.setdefault(int(wd["y0"] // 12), []).append(wd)
    seen = set()
    for wd in words:
        key = int(wd["y0"] // 12)
        near = []
        for k in range(key - 2, key + 3):
            near.extend(buckets.get(k, []))
        for cand in near:
            if cand is wd:
                continue
            pair = tuple(sorted((id(wd), id(cand))))
            if pair in seen:
                continue
            seen.add(pair)
            a = inter(wd, cand)
            if a <= 0:
                continue
            small = min(area(wd), area(cand))
            if small > 0 and a / small > OVERLAP_RATIO:
                out["overlaps"].append(f"{wd['text'][:12]} × {cand['text'][:12]}")
    return out


def main():
    ap = argparse.ArgumentParser(description="PPTX 渲染层复核（越界/压字/过小字号）")
    ap.add_argument("pptx")
    ap.add_argument("--expect-pages", type=int, default=0)
    ap.add_argument("--margin", type=float, default=MARGIN_PT)
    ap.add_argument("--json", default=None)
    args = ap.parse_args()

    pdf, err = to_pdf(args.pptx)
    if not pdf:
        print(f"! {err}", file=sys.stderr)
        sys.exit(0)  # 复核不可用不算失败，提示即可

    pages = word_boxes(pdf)
    report = {"file": os.path.basename(args.pptx), "pdf": os.path.basename(pdf),
              "pages": len(pages), "margin_pt": args.margin, "problems": []}
    bad_pages = 0
    print(f"渲染复核：{os.path.basename(args.pptx)} → {len(pages)} 页")
    print()
    print("| 页 | 词块 | 越界 | 压字 | 过小字号 |")
    print("| --: | --: | --: | --: | --: |")
    for i, page in enumerate(pages, 1):
        r = check_page(page, args.margin)
        if r["outside"] or r["overlaps"] or r["tiny"]:
            bad_pages += 1
        print(f"| {i} | {r['words']} | {len(r['outside'])} | {len(r['overlaps'])} | {len(r['tiny'])} |")
        if r["outside"]:
            report["problems"].append({"page": i, "type": "outside",
                                       "detail": r["outside"][:6]})
        if r["overlaps"]:
            report["problems"].append({"page": i, "type": "overlap",
                                       "detail": r["overlaps"][:6]})
        if r["tiny"]:
            report["problems"].append({"page": i, "type": "tiny",
                                       "detail": r["tiny"][:6]})
    print()
    if args.expect_pages and len(pages) != args.expect_pages:
        print(f"! 页数 {len(pages)} 与预期 {args.expect_pages} 不一致", file=sys.stderr)
        report["problems"].append({"type": "page-count",
                                   "detail": [f"实际 {len(pages)} 页，预期 {args.expect_pages} 页"]})
    if report["problems"]:
        print(f"✗ {bad_pages} 页有问题，共 {len(report['problems'])} 处：")
        for p in report["problems"]:
            print(f"  - P{p.get('page', '?')} [{p['type']}] {'；'.join(p['detail'])}")
    else:
        print("✓ 未发现越界、压字、过小字号")
    if args.json:
        json.dump(report, open(args.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"✓ 明细 {args.json}")


if __name__ == "__main__":
    main()
