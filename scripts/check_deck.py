#!/usr/bin/env python3
"""zgs-ppt / check_deck.py —— 交付前的文件级检查：页数、尺寸、备注是否写入口播。

用法：
    python3 check_deck.py 展示.pptx --expect 13
    python3 check_deck.py 展示.pptx --expect-min 12
"""
from __future__ import annotations

import argparse
import sys

from pptx import Presentation


def main():
    ap = argparse.ArgumentParser(description="PPTX 文件级检查")
    ap.add_argument("pptx")
    ap.add_argument("--expect", type=int, default=0, help="期望页数（精确）")
    ap.add_argument("--expect-min", type=int, default=0, help="期望最少页数")
    ap.add_argument("--no-notes", action="store_true", help="不检查备注")
    args = ap.parse_args()

    p = Presentation(args.pptx)
    n = len(p.slides)
    w = round(p.slide_width / 914400, 2)
    h = round(p.slide_height / 914400, 2)
    notes = sum(1 for s in p.slides if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip())
    print(f"{args.pptx}: {n} 页, {w}×{h} in, 有备注 {notes}/{n} 页")

    bad = []
    if args.expect and n != args.expect:
        bad.append(f"页数 {n} ≠ 期望 {args.expect}")
    if args.expect_min and n < args.expect_min:
        bad.append(f"页数 {n} < 下限 {args.expect_min}（本技能要求 ≥12 页）")
    if (w, h) not in ((13.33, 7.5), (10.0, 7.5)):
        bad.append(f"页面尺寸 {w}×{h} 不是 16:9 或 4:3")
    if not args.no_notes and notes < n:
        bad.append(f"{n - notes} 页没有备注（口播要点应写进备注栏）")
    for b in bad:
        print("✗", b, file=sys.stderr)
    if bad:
        sys.exit(1)
    print("✓ 文件级检查通过")


if __name__ == "__main__":
    main()
