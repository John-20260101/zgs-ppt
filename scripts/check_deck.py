#!/usr/bin/env python3
"""zgs-ppt / check_deck.py —— 交付前的文件级检查：页数、尺寸、备注是否写入口播。

用法：
    python3 check_deck.py 展示.pptx --group Z5
    python3 check_deck.py 展示.pptx --expect 18
    python3 check_deck.py 展示.pptx --expect-min 15
    python3 check_deck.py 展示.pptx --allow-size 10x7.5     # 模板为 4:3 时显式放行
"""
from __future__ import annotations

import argparse
import sys

from pptx import Presentation

# 与 build_pptx.py 的 GROUPS 保持同一条区间：所有组别 15–20 页（0.3.0 起的硬约束）
PAGE_RANGE = {"Z1": (15, 20), "Z2": (15, 20), "Z3": (15, 20),
              "Z4": (15, 20), "Z5": (15, 20)}
DEFAULT_RANGE = (15, 20)


def main():
    ap = argparse.ArgumentParser(description="PPTX 文件级检查")
    ap.add_argument("pptx")
    ap.add_argument("--group", default=None, help="Z1–Z5，用该组别的页数区间作默认下限/上限")
    ap.add_argument("--expect", type=int, default=0, help="期望页数（精确）")
    ap.add_argument("--expect-min", type=int, default=0, help="期望最少页数（覆盖组别下限）")
    ap.add_argument("--allow-size", default=None,
                    help="额外允许的页面尺寸（英寸，如 10x7.5 表示 4:3 模板）")
    ap.add_argument("--no-notes", action="store_true", help="不检查备注")
    args = ap.parse_args()

    p = Presentation(args.pptx)
    n = len(p.slides)
    w = round(p.slide_width / 914400, 2)
    h = round(p.slide_height / 914400, 2)
    notes = sum(1 for s in p.slides if s.has_notes_slide
                and s.notes_slide.notes_text_frame.text.strip())
    print(f"{args.pptx}: {n} 页, {w}×{h} in, 有备注 {notes}/{n} 页")

    lo, hi = PAGE_RANGE.get(args.group or "", DEFAULT_RANGE)
    if args.expect_min:
        lo = args.expect_min
    bad = []
    if args.expect and n != args.expect:
        bad.append(f"页数 {n} ≠ 期望 {args.expect}")
    if n < lo:
        bad.append(f"页数 {n} < 下限 {lo}（{'组别 ' + args.group if args.group else '本技能要求'} "
                   f"{lo}–{hi} 页；不足时按 SKILL.md §2 的四类必补页补："
                   "岗位调研实录 → 人岗匹配度 → 二次证据 → 风险预案）")
    if n > hi:
        bad.append(f"页数 {n} > 上限 {hi}：正文超时，合并同类页而不是每页拉长")
    size_ok = (w, h) in ((13.33, 7.5), (10.0, 7.5))
    if args.allow_size:
        try:
            aw, ah = (float(x) for x in args.allow_size.lower().split("x"))
            size_ok = size_ok or (abs(w - aw) <= 0.02 and abs(h - ah) <= 0.02)
        except ValueError:
            bad.append(f"--allow-size 格式应为 WxH（英寸），收到 {args.allow_size!r}")
    if not size_ok:
        bad.append(f"页面尺寸 {w}×{h} 不是 16:9 或 4:3；模板比例特殊时用 --allow-size 显式放行")
    if not args.no_notes and notes < n:
        bad.append(f"{n - notes} 页没有备注（口播要点应写进备注栏）")
    for b in bad:
        print("✗", b, file=sys.stderr)
    if bad:
        sys.exit(1)
    print("✓ 文件级检查通过")


if __name__ == "__main__":
    main()
