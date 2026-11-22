#!/usr/bin/env python3
"""zgs-ppt / check_outline.py —— 核「outline.json（唯一事实源）」与「-大纲.md（人读视图）」是否同源。

为什么要有这个脚本：SKILL.md §3.1 原来把这段检查内联成一段 heredoc，模型粘贴执行时，
如果 md 是旧格式（正则 0 匹配）就会输出「全部对齐」——在最该拦的时候发绿灯。
把「0 匹配」当硬失败，比宽容解析更安全。

用法：
    python3 check_outline.py outline.json 姓名-赛道展示-大纲.md
    python3 check_outline.py outline.json 姓名-赛道展示-大纲.md --quiet

退出码：
    0  对齐
    2  不一致或 md 格式不符（附逐条差异）
    1  文件读不到 / JSON 解析失败
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

FIELD_RE = re.compile(r"^- \*\*页面标题\*\*：(.+?)（一级）\s*(.+?)（重点句）\s*$", re.M)
MAIN_RE = re.compile(r"^- \*\*主标题\*\*：(.+?)\s*$", re.M)


def load_json(path: str) -> dict:
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"[E] 读不到或解析失败：{path}：{e}", file=sys.stderr)
        sys.exit(1)


def main() -> int:
    ap = argparse.ArgumentParser(description="outline.json 与 大纲.md 一致性检查")
    ap.add_argument("outline")
    ap.add_argument("md")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    try:
        md = open(args.md, encoding="utf-8-sig").read()
    except OSError as e:
        print(f"[E] 读不到 {args.md}：{e}", file=sys.stderr)
        return 1
    data = load_json(args.outline)
    slides = data.get("slides") or []
    if not slides:
        print("[E] outline.json 没有 slides", file=sys.stderr)
        return 2

    content = FIELD_RE.findall(md)
    mains = MAIN_RE.findall(md)
    problems: list[str] = []

    n_slides = len(slides)
    n_content_expected = max(0, n_slides - 2)
    if len(mains) != 2:
        problems.append(
            f"封面/封底主标题应各有 1 条「- **主标题**：…」，实际 {len(mains)} 条"
            "——md 可能不是 0.4.3 格式（旧格式用 `## P1 封面 · cover` + `- 文案：…`）"
        )
    if len(content) != n_content_expected:
        problems.append(
            f"内容页「- **页面标题**：维度（一级） 论点句（重点句）」应有 "
            f"{n_content_expected} 条（{n_slides} 页 − 封面封底），实际 {len(content)} 条"
        )
    if len(content) > n_content_expected:
        problems.append("md 内容页比 JSON 多，后续比对会错位：先删 md 的多余页，别改 JSON 去迁就")

    if len(mains) == 2:
        if slides[0].get("title") != mains[0].strip():
            problems.append(f"P1 主标题：JSON {slides[0].get('title')!r} ≠ md {mains[0].strip()!r}")
        if slides[-1].get("title") != mains[1].strip():
            problems.append(f"P{n_slides} 封底主标题：JSON {slides[-1].get('title')!r} "
                            f"≠ md {mains[1].strip()!r}")

    for idx, (dim, title) in enumerate(content, 2):
        if idx - 1 > n_slides - 1:
            break
        page = slides[idx - 1]
        if (page.get("dim") or "").strip() != dim.strip():
            problems.append(f"P{idx} 维度：JSON {page.get('dim')!r} ≠ md {dim.strip()!r}")
        if (page.get("title") or "").strip() != title.strip():
            problems.append(f"P{idx} 论点：JSON {page.get('title')!r} ≠ md {title.strip()!r}")

    if not args.quiet:
        print(f"JSON {n_slides} 页 ｜ 大纲：主标题 {len(mains)} 条、内容页 {len(content)} 条")
    if problems:
        print(f"✗ 需修（{len(problems)} 项）：", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        print("  规矩：改内容只改 JSON，md 是它的视图；改完两边页序与标题必须保持一致。",
              file=sys.stderr)
        return 2
    if not args.quiet:
        print("✓ 全部对齐")
    return 0


if __name__ == "__main__":
    sys.exit(main())
