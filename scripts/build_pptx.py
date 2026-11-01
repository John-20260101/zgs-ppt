#!/usr/bin/env python3
"""zgs-ppt / build_pptx.py —— 把 outline.json 渲染成职规赛生涯发展展示 / 求职综合展示 PPTX。

设计要点：
1. 先在 13.333×7.5 in 的画布上排版，再按目标模板的页面尺寸等比缩放，因此 4:3 模板也能用。
2. 每段文字在落版前先估算行数与高度，超出就自动缩字号；缩到下限仍溢出则记 warning，绝不静默裁字。
3. 输出 pptx + 审计报告（溢出、页数、评分维度覆盖、素材缺口），让"看不见的画面"也能被检查。

用法：
    python3 build_pptx.py outline.json -o 展示.pptx
    python3 build_pptx.py outline.json -o 展示.pptx --template 学校模板.pptx --pdf
    python3 build_pptx.py outline.json --audit-only
"""
from __future__ import annotations

import argparse
import json
import math
import os
import shutil
import subprocess
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- 画布与主题

CW, CH = 13.333, 7.5          # 设计画布（英寸，16:9）
M = 0.72                       # 页边距
TITLE_Y = 0.62
BODY_Y = 1.92
BODY_H = 4.55
FOOT_Y = 6.92

DEFAULT_THEMES = {
    "academic": {
        "ink": "#132133", "accent": "#1F4E79", "accent2": "#C08A2E",
        "muted": "#6C7A8C", "band": "#F1F4F9", "line": "#D8DFE9", "bg": "#FFFFFF",
    },
    "tech": {
        "ink": "#0B1622", "accent": "#0E7C86", "accent2": "#F2A93B",
        "muted": "#5F7183", "band": "#EDF4F6", "line": "#CFDDE2", "bg": "#FFFFFF",
    },
    "warm": {
        "ink": "#2B2118", "accent": "#A8432F", "accent2": "#7A8C5A",
        "muted": "#7C6E63", "band": "#F7F1EA", "line": "#E6D8C9", "bg": "#FFFDFA",
    },
}

LAYOUT_ORDER = ["cover", "agenda", "section", "bullets", "detail", "cards",
                "kpi", "timeline", "process", "compare", "matrix", "table",
                "evidence", "twocol", "image", "quote", "closing"]

GROUPS = {
    "Z1": ("成长赛道·高教组", ["职业目标", "学习实践行动", "优化改进"], (12, 16)),
    "Z2": ("成长赛道·职教组", ["职业目标", "学习实践行动", "优化改进"], (12, 16)),
    "Z3": ("就业赛道·本科组", ["职业目标", "岗位胜任力", "发展潜力"], (12, 15)),
    "Z4": ("就业赛道·研究生组", ["职业目标", "岗位胜任力", "发展潜力"], (12, 15)),
    "Z5": ("就业赛道·职教组", ["职业目标", "岗位胜任力", "发展潜力"], (12, 15)),
}


DEFAULT_FONT = "Microsoft YaHei"   # 学生多在 Windows 上打开与放映；mac 装了 Office 也有


def cjk_font() -> str:
    """默认用跨机器最稳的中文字体名；可用 outline 的 theme_spec.font 或 --font 覆盖。"""
    return DEFAULT_FONT


class Theme:
    def __init__(self, spec: dict):
        self.name = spec.get("name", "academic")
        for k in ("ink", "accent", "accent2", "muted", "band", "line", "bg"):
            setattr(self, k, hexc(spec.get(k, DEFAULT_THEMES["academic"][k])))
        self.font = spec.get("font") or cjk_font()


def hexc(s: str) -> RGBColor:
    s = s.strip().lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


# ---------------------------------------------------------------- 文本度量

def text_units(s: str) -> float:
    """把字符串折算成"全角字宽"数量：CJK=1，拉丁/数字=0.55。"""
    n = 0.0
    for ch in s:
        n += 1.0 if ord(ch) > 0x2E80 else 0.55
    return n


def est_lines(text: str, width_in: float, size_pt: float) -> int:
    per_line = max(1.0, width_in * 72 / (size_pt * 1.10))
    return max(1, math.ceil(text_units(text) / per_line))


def est_block_height(paras: list[str], width_in: float, size_pt: float,
                     gap_pt: float = 0.35) -> float:
    total = 0.0
    for p in paras:
        lines = est_lines(p, width_in, size_pt)
        total += lines * size_pt * 1.32 / 72 + gap_pt * size_pt / 72
    return total


def fit_size(paras: list[str], width_in: float, height_in: float,
             base: float, minimum: float = 13.0) -> tuple[float, float]:
    """返回 (可用字号, 填充率)。填充率 >1 表示缩到下限仍溢出。"""
    size = base
    while size >= minimum:
        h = est_block_height(paras, width_in, size)
        if h <= height_in:
            return size, h / height_in
        size -= 1
    h = est_block_height(paras, width_in, minimum)
    return minimum, h / height_in


# ---------------------------------------------------------------- 绘图基元

class Ctx:
    def __init__(self, prs, theme: Theme, group: str, total: int):
        self.prs = prs
        self.theme = theme
        self.sx = prs.slide_width / Emu(int(CW * 914400))
        self.sy = prs.slide_height / Emu(int(CH * 914400))
        self.fs = min(self.sx, self.sy)
        self.group = group
        self.total = total
        self.dir = os.path.dirname(os.path.abspath(__file__))
        self.boxes: list[tuple[float, float, float, float, str]] = []

    def x(self, v: float) -> Emu:
        return Emu(int(v * 914400 * self.sx))

    def y(self, v: float) -> Emu:
        return Emu(int(v * 914400 * self.sy))

    def w(self, v: float) -> Emu:
        return Emu(int(v * 914400 * self.sx))

    def h(self, v: float) -> Emu:
        return Emu(int(v * 914400 * self.sy))

    def pt(self, v: float) -> Pt:
        return Pt(max(8.0, v * self.fs))


def add_text(slide, ctx: Ctx, x, y, w, h, text, *, size=18, bold=False,
             color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.18, space_after=4, font=None, italic=False,
             char_spacing=None):
    ctx.boxes.append((x, y, w, h, text if isinstance(text, str) else " ".join(text)))
    box = slide.shapes.add_textbox(ctx.x(x), ctx.y(y), ctx.w(w), ctx.h(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = ctx.pt(space_after)
        run = p.add_run()
        run.text = line
        f = run.font
        f.size = ctx.pt(size)
        f.bold = bold
        f.italic = italic
        f.color.rgb = color or ctx.theme.ink
        name = font or ctx.theme.font
        f.name = name
        rpr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set("typeface", name)
        if char_spacing is not None:
            rpr.set("spc", str(int(char_spacing * 100)))
    return box


def add_rect(slide, ctx, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = slide.shapes.add_shape(shape, ctx.x(x), ctx.y(y), ctx.w(w), ctx.h(h))
    s.shadow.inherit = False
    if color is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = color
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.9)
    s.text_frame.word_wrap = True
    return s


def add_line(slide, ctx, x1, y1, x2, y2, color, width=1.2):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                   ctx.x(x1), ctx.y(y1), ctx.x(x2), ctx.y(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


# ---------------------------------------------------------------- 页面通用件

def chrome(slide, ctx: Ctx, kicker: str | None, title: str | None, idx: int):
    t = ctx.theme
    add_rect(slide, ctx, 0, 0, CW, 0.13, t.accent)
    if kicker:
        add_text(slide, ctx, M, 0.34, 6.0, 0.26, kicker, size=11.5,
                 color=t.accent2, bold=True, char_spacing=1.6)
    if title:
        tsize, _ = fit_size([title], CW - 2 * M - 0.2, 0.92, 30, 21)
        add_text(slide, ctx, M, TITLE_Y, CW - 2 * M, 0.95, title, size=tsize,
                 bold=True, color=t.ink, line_spacing=1.08)
        add_rect(slide, ctx, M, TITLE_Y + 0.98, 0.9, 0.055, t.accent2)
    # 页脚
    add_text(slide, ctx, M, FOOT_Y, 7.4, 0.3,
             f"{ctx.group}", size=10, color=t.muted)
    add_text(slide, ctx, CW - M - 2.6, FOOT_Y, 2.6, 0.3,
             f"{idx:02d} / {ctx.total:02d}", size=10, color=t.muted,
             align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- 各页型

def s_cover(slide, ctx, spec, idx, warn):
    t = ctx.theme
    add_rect(slide, ctx, 0, 0, CW, CH, t.bg)
    add_rect(slide, ctx, 0, 0, 0.26, CH, t.accent)
    add_rect(slide, ctx, 0.26, CH - 1.55, CW - 0.26, 1.55, t.band)

    title = spec.get("title", "")
    size, fill = fit_size([title], CW - 2.9, 2.1, 44, 26)
    add_text(slide, ctx, 1.15, 1.75, CW - 2.9, 2.1, title, size=size,
             bold=True, color=t.ink, line_spacing=1.1)

    sub = spec.get("subtitle", "")
    add_text(slide, ctx, 1.15, 4.05, CW - 2.9, 0.9, sub, size=17,
             color=t.accent, line_spacing=1.3)

    meta = spec.get("meta") or []
    if meta:
        add_text(slide, ctx, 1.15, CH - 1.15, CW - 2.9, 0.75,
                 "　·　".join(meta), size=12, color=t.muted)
    else:
        warn.append(("W", "cover-meta", "封面缺少 meta（赛道/组别/学校/姓名/目标岗位）"))
    return {"title": title, "fill": fill}


def s_section(slide, ctx, spec, idx, warn):
    t = ctx.theme
    num = spec.get("kicker", "")
    add_text(slide, ctx, M - 0.08, 1.28, 3.2, 1.55, num, size=78, bold=True,
             color=t.band if t.band != RGBColor(255, 255, 255) else hexc("#EDEFF4"))
    title = spec.get("title", "")
    size, fill = fit_size([title], CW - 2 * M, 1.15, 34, 22)
    add_text(slide, ctx, M, 3.02, CW - 2 * M, 1.15, title, size=size, bold=True,
             color=t.ink)
    note = spec.get("note", "")
    if note:
        add_text(slide, ctx, M, 4.35, CW - 2 * M - 1.2, 0.95, note, size=15,
                 color=t.muted)
    add_rect(slide, ctx, M + 0.02, 4.2, 0.72, 0.05, t.accent2)
    chrome(slide, ctx, None, None, idx)
    return {"title": title, "fill": fill}


def s_bullets(slide, ctx, spec, idx, warn):
    t = ctx.theme
    items = spec.get("bullets") or []
    if not items:
        warn.append(("E", "bullets-empty", "bullets 页没有任何要点"))
        return {"title": spec.get("title", ""), "fill": 0.0}
    if len(items) > 6:
        warn.append(("W", "bullets-count",
                     f"要点 {len(items)} 条，建议 ≤6（当前页：{spec.get('title')}）"))
    norm = []
    for it in items:
        if isinstance(it, dict):
            norm.append({"text": str(it.get("text", "")),
                         "note": str(it.get("note") or ""),
                         "em": bool(it.get("em"))})
        else:
            norm.append({"text": str(it), "note": "", "em": False})
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)

    box_w = CW - 2 * M - 0.52
    n = len(norm)
    gap = min(1.18, BODY_H / n)
    avail = gap - 0.10
    base = 22 if n <= 3 else (20 if n == 4 else 18)
    note_size = 11.5
    size = base

    def item_h(it, sz):
        h = est_block_height([it["text"]], box_w - 0.52, sz)
        if it["note"]:
            h += est_block_height([it["note"]], box_w - 0.52, note_size) + 0.03
        return h

    heights = [item_h(it, size) for it in norm]
    while size > 13 and max(heights) > avail:
        size -= 1
        heights = [item_h(it, size) for it in norm]
    fill = max(heights) / avail if avail > 0 else 0
    if fill > 1:
        warn.append(("W", "bullets-overflow",
                     f"文字缩到下限仍溢出（{fill:.0%}）：{spec.get('title')}"))
    if size < 15:
        warn.append(("W", "bullets-tiny",
                     f"要点字号降到 {size}pt，建议删掉 20% 文字：{spec.get('title')}"))

    start = BODY_Y + 0.06
    for i, (it, h) in enumerate(zip(norm, heights)):
        yy = start + i * gap
        add_rect(slide, ctx, M, yy + 0.10, 0.3, 0.055, t.accent2 if it["em"] else t.line)
        text_h = est_block_height([it["text"]], box_w - 0.52, size)
        add_text(slide, ctx, M + 0.52, yy, box_w - 0.52, text_h + 0.06, it["text"],
                 size=size, bold=it["em"], color=t.accent if it["em"] else t.ink,
                 line_spacing=1.22)
        if it["note"]:
            add_text(slide, ctx, M + 0.52, yy + text_h + 0.05,
                     box_w - 0.52, h - text_h, it["note"], size=note_size,
                     color=t.muted, line_spacing=1.2)
    src = spec.get("source")
    if src:
        add_text(slide, ctx, M, FOOT_Y - 0.42, CW - 2 * M, 0.3, f"数据来源：{src}",
                 size=10, color=t.muted)
    return {"title": spec.get("title", ""), "fill": fill}


def s_kpi(slide, ctx, spec, idx, warn):
    t = ctx.theme
    items = spec.get("items") or []
    if not 2 <= len(items) <= 4:
        warn.append(("W", "kpi-count", f"数字卡 {len(items)} 个，建议 2–4 个"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    n = max(1, len(items))
    gap = 0.34
    cw = (CW - 2 * M - gap * (n - 1)) / n
    inner = cw - 0.56
    top = BODY_Y + 0.3

    # 先按"数字一行放下"定字号；标签行数也参与卡高计算，避免标签折行压住注释
    sizes = []
    for it in items:
        val = str(it.get("value", ""))
        size = 34
        while size > 20 and est_lines(val, inner, size) > 1:
            size -= 1
        lines = est_lines(val, inner, size)
        lab = str(it.get("label", ""))
        lab_lines = est_lines(lab, inner, 14)
        sizes.append((val, size, lines, lab, lab_lines))
        if lines > 1:
            warn.append(("W", "kpi-wrap", f"数字卡数值换行：{val}（建议 ≤8 字）"))
        if lab_lines > 2:
            warn.append(("W", "kpi-label-wrap", f"数字卡标签 {lab_lines} 行，建议 ≤14 字：{lab}"))
    max_lines = max(x[2] for x in sizes)
    max_lab = max(x[4] for x in sizes)
    value_h = max_lines * 34 * 1.30 / 72
    label_h = max_lab * 14 * 1.32 / 72
    has_note = any(it.get("note") for it in items)
    card_h = 0.28 + value_h + 0.20 + label_h + (0.34 if has_note else 0) + 0.20
    card_h = max(2.05, min(2.95, card_h))

    for i, (it, (val, size, lines, lab, lab_lines)) in enumerate(zip(items, sizes)):
        xx = M + i * (cw + gap)
        add_rect(slide, ctx, xx, top, cw, card_h, t.band)
        add_rect(slide, ctx, xx, top, 0.07, card_h, t.accent2)
        vh = lines * size * 1.30 / 72
        add_text(slide, ctx, xx + 0.3, top + 0.26, inner, vh + 0.12, val,
                 size=size, bold=True, color=t.accent, line_spacing=1.18)
        label_y = top + 0.26 + vh + 0.20
        add_text(slide, ctx, xx + 0.3, label_y, inner, label_h + 0.1, lab,
                 size=14, color=t.ink, line_spacing=1.24)
        note = it.get("note")
        if note:
            add_text(slide, ctx, xx + 0.3, label_y + label_h + 0.08, inner, 0.36,
                     str(note), size=10.5, color=t.muted)
    tail = spec.get("note")
    if tail:
        add_text(slide, ctx, M, top + card_h + 0.34, CW - 2 * M, 0.55, tail,
                 size=14, color=t.accent, bold=True)
    src = spec.get("source")
    if src:
        add_text(slide, ctx, M, FOOT_Y - 0.42, CW - 2 * M, 0.3, f"数据来源：{src}",
                 size=10, color=t.muted)
    return {"title": spec.get("title", ""), "fill": None}


def s_timeline(slide, ctx, spec, idx, warn):
    t = ctx.theme
    items = spec.get("items") or []
    if not 3 <= len(items) <= 5:
        warn.append(("W", "timeline-count", f"时间轴节点 {len(items)} 个，建议 3–5 个"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    n = max(1, len(items))
    left, right = M + 0.35, CW - M - 0.35
    span = right - left
    axis_y = 4.02
    add_line(slide, ctx, left, axis_y, right, axis_y, t.line, 2.2)
    step = span / max(1, n - 1) if n > 1 else 0
    for i, it in enumerate(items):
        cx = left + i * step
        add_rect(slide, ctx, cx - 0.085, axis_y - 0.085, 0.17, 0.17, t.accent,
                 shape=MSO_SHAPE.OVAL)
        when = str(it.get("when", ""))
        what = str(it.get("what", ""))
        above = i % 2 == 0
        tw = min(2.55, span / max(1, n) + 0.5)
        tx = min(max(cx - tw / 2, M), CW - M - tw)
        detail = str(it.get("detail", ""))
        if above:
            add_text(slide, ctx, tx, axis_y - 0.52, tw, 0.38, when, size=13,
                     bold=True, color=t.accent, align=PP_ALIGN.CENTER)
            sz, fill = fit_size([what], tw, 0.95, 12.5, 10)
            add_text(slide, ctx, tx, axis_y - 1.72, tw, 0.95, what, size=sz,
                     color=t.ink, align=PP_ALIGN.CENTER, line_spacing=1.22)
            if detail:
                add_text(slide, ctx, tx, axis_y - 0.78, tw, 0.28, detail,
                         size=9.5, color=t.muted, align=PP_ALIGN.CENTER)
        else:
            sz, fill = fit_size([what], tw, 0.95, 12.5, 10)
            add_text(slide, ctx, tx, axis_y + 0.34, tw, 0.95, what, size=sz,
                     color=t.ink, align=PP_ALIGN.CENTER, line_spacing=1.22)
            if detail:
                add_text(slide, ctx, tx, axis_y + 1.32, tw, 0.3, detail,
                         size=9.5, color=t.muted, align=PP_ALIGN.CENTER)
            add_text(slide, ctx, tx, axis_y + 1.66, tw, 0.38, when, size=13,
                     bold=True, color=t.accent, align=PP_ALIGN.CENTER)
        if fill > 1:
            warn.append(("W", "timeline-overflow", f"时间轴节点文字溢出：{when}"))
    return {"title": spec.get("title", ""), "fill": 0.85}


def s_compare(slide, ctx, spec, idx, warn):
    t = ctx.theme
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    def _panel(v):
        if isinstance(v, list):
            return {"head": "", "items": v}
        return v or {}
    panels = [_panel(spec.get("left")), _panel(spec.get("right"))]
    heads: list[str] = []
    pw = (CW - 2 * M - 1.5) / 2
    top = BODY_Y + 0.18
    ph = 3.5
    for i, p in enumerate(panels):
        xx = M + i * (pw + 1.5)
        add_rect(slide, ctx, xx, top, pw, ph, t.band)
        add_rect(slide, ctx, xx, top, pw, 0.55, t.accent if i == 0 else t.accent2)
        add_text(slide, ctx, xx + 0.28, top + 0.12, pw - 0.56, 0.35,
                 str(p.get("head", "")), size=14, bold=True,
                 color=hexc("#FFFFFF"))
        items = [str(x) for x in (p.get("items") or [])]
        if len(items) > 4:
            heads.append(f"{p.get('head') or '对比栏'}：{len(items)} 条")
        size, fill = fit_size([f"· {x}" for x in items], pw - 0.56, ph - 0.85, 14, 10.5)
        add_text(slide, ctx, xx + 0.28, top + 0.78, pw - 0.56, ph - 0.95,
                 "\n".join(f"· {x}" for x in items), size=size, color=t.ink,
                 line_spacing=1.25, space_after=7)
        if fill > 1:
            warn.append(("W", "compare-overflow", f"对比栏溢出：{p.get('head')}"))
    if heads:
        warn.append(("W", "compare-items", "；".join(heads) + "，建议每栏 ≤4 条"))
    cx = M + pw + 0.75
    add_rect(slide, ctx, cx - 0.26, top + ph / 2 - 0.26, 0.52, 0.52, t.accent2,
             shape=MSO_SHAPE.OVAL)
    add_text(slide, ctx, cx - 0.26, top + ph / 2 - 0.16, 0.52, 0.32, "VS",
             size=12, bold=True, color=hexc("#FFFFFF"), align=PP_ALIGN.CENTER)
    return {"title": spec.get("title", ""), "fill": 0.85}


def s_table(slide, ctx, spec, idx, warn):
    t = ctx.theme
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    headers = [str(h) for h in (spec.get("headers") or [])]
    rows = [[str(c) for c in r] for r in (spec.get("rows") or [])]
    if not headers or not rows:
        warn.append(("E", "table-empty", "table 页缺少表头或数据行"))
        return {"title": spec.get("title", ""), "fill": 0.0}
    if len(headers) > 6:
        warn.append(("W", "table-cols", f"表格 {len(headers)} 列，建议 ≤6"))
    if len(rows) > 6:
        warn.append(("W", "table-rows", f"表格 {len(rows)} 行，建议 ≤6"))
    nrow, ncol = len(rows) + 1, len(headers)
    width = CW - 2 * M
    row_h = min(0.62, (BODY_H - 0.2) / nrow)
    shape = slide.shapes.add_table(nrow, ncol, ctx.x(M), ctx.y(BODY_Y + 0.12),
                                   ctx.w(width), ctx.h(row_h * nrow))
    table = shape.table
    table.first_row = True
    col_w = width / ncol
    for j in range(ncol):
        table.columns[j].width = ctx.w(col_w)
    for i in range(nrow):
        table.rows[i].height = ctx.h(row_h)
    size = 13 if ncol <= 4 else 11.5
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = t.accent
        cell.margin_left = cell.margin_right = Emu(int(0.08 * 914400))
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = h
        r.font.size = ctx.pt(size)
        r.font.bold = True
        r.font.color.rgb = hexc("#FFFFFF")
        r.font.name = t.font
    for i, row in enumerate(rows, start=1):
        for j in range(ncol):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = t.bg if i % 2 else t.band
            cell.margin_left = cell.margin_right = Emu(int(0.08 * 914400))
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = row[j] if j < len(row) else ""
            r.font.size = ctx.pt(size)
            r.font.color.rgb = t.ink
            r.font.name = t.font
    note = spec.get("note")
    if note:
        add_text(slide, ctx, M, BODY_Y + 0.12 + row_h * nrow + 0.12,
                 CW - 2 * M, 0.4, note, size=10.5, color=t.muted)
    return {"title": spec.get("title", ""), "fill": None}


def s_image(slide, ctx, spec, idx, warn):
    t = ctx.theme
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    path = spec.get("path", "")
    cap = spec.get("caption", "")
    box_w, box_h = CW - 2 * M, BODY_H - (0.5 if cap else 0.1)
    if path and os.path.exists(path):
        from PIL import Image
        iw, ih = Image.open(path).size
        scale = min(box_w / iw, box_h / ih)
        dw, dh = iw * scale, ih * scale
        x = M + (box_w - dw) / 2
        y = BODY_Y + 0.12 + (box_h - dh) / 2
        slide.shapes.add_picture(path, ctx.x(x), ctx.y(y), ctx.w(dw), ctx.h(dh))
    else:
        add_rect(slide, ctx, M, BODY_Y + 0.2, box_w, box_h - 0.2, t.band,
                 line=t.line)
        add_text(slide, ctx, M, BODY_Y + box_h / 2 - 0.2, box_w, 0.5,
                 "【图片位】待补：实习/现场/作品照片", size=15, color=t.muted,
                 align=PP_ALIGN.CENTER)
        warn.append(("W", "image-missing", f"图片缺失：{path or '未填写 path'}"))
    if cap:
        add_text(slide, ctx, M, BODY_Y + BODY_H - 0.28, box_w, 0.32, cap,
                 size=11, color=t.muted, align=PP_ALIGN.CENTER)
    return {"title": spec.get("title", ""), "fill": None}


def s_quote(slide, ctx, spec, idx, warn):
    t = ctx.theme
    add_rect(slide, ctx, 0, 0, CW, CH, t.bg)
    add_rect(slide, ctx, 0, 0, 0.26, CH, t.accent2)
    text = spec.get("text", "")
    size, fill = fit_size([text], CW - 3.4, 2.6, 32, 18)
    add_text(slide, ctx, 1.5, 2.05, CW - 3.0, 2.6, text, size=size, bold=True,
             color=t.ink, line_spacing=1.25, align=PP_ALIGN.CENTER)
    by = spec.get("by")
    if by:
        add_text(slide, ctx, 1.5, 4.95, CW - 3.0, 0.4, f"— {by}", size=13,
                 color=t.muted, align=PP_ALIGN.CENTER)
    chrome(slide, ctx, None, None, idx)
    return {"title": text[:14], "fill": fill}


def s_closing(slide, ctx, spec, idx, warn):
    t = ctx.theme
    add_rect(slide, ctx, 0, 0, CW, CH, t.bg)
    add_rect(slide, ctx, 0, CH - 1.35, CW, 1.35, t.accent)
    title = spec.get("title", "")
    size, fill = fit_size([title], CW - 2.4, 1.9, 38, 20)
    add_text(slide, ctx, 1.2, 2.35, CW - 2.4, 1.9, title, size=size, bold=True,
             color=t.ink, align=PP_ALIGN.CENTER, line_spacing=1.2)
    sub = spec.get("subtitle", "")
    if sub:
        add_text(slide, ctx, 1.2, 4.45, CW - 2.4, 0.6, sub, size=15,
                 color=t.accent, align=PP_ALIGN.CENTER)
    add_text(slide, ctx, 1.2, CH - 0.95, CW - 2.4, 0.5,
             spec.get("footer", "谢谢聆听 · 恳请各位评委老师批评指正"), size=13,
             color=hexc("#FFFFFF"), align=PP_ALIGN.CENTER)
    return {"title": title, "fill": fill}


def s_agenda(slide, ctx, spec, idx, warn):
    """目录页：给阅读型 PPT 一个"这份材料怎么读"的入口。"""
    t = ctx.theme
    items = spec.get("items") or []
    if not items:
        warn.append(("E", "agenda-empty", "agenda 页没有任何条目"))
        return {"title": spec.get("title", ""), "fill": 0.0}
    if len(items) > 6:
        warn.append(("W", "agenda-count", f"目录 {len(items)} 项，建议 ≤6"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)

    col_w = (CW - 2 * M - 0.62) / 2
    rows = (len(items) + 1) // 2
    row_h = min(1.62, (BODY_H - 0.25) / max(rows, 1))
    for i, it in enumerate(items):
        col, row = i % 2, i // 2
        xx = M + col * (col_w + 0.62)
        yy = BODY_Y + 0.18 + row * row_h
        num = str(it.get("no") or f"{i + 1:02d}")
        add_text(slide, ctx, xx, yy - 0.04, 0.95, 0.62, num, size=32, bold=True,
                 color=t.accent2, line_spacing=1.0)
        add_text(slide, ctx, xx + 0.98, yy - 0.02, col_w - 0.98, 0.5,
                 str(it.get("title", "")), size=18, bold=True, color=t.ink)
        note = it.get("note")
        if note:
            add_text(slide, ctx, xx + 0.98, yy + 0.46, col_w - 0.98, row_h - 0.52,
                     str(note), size=12, color=t.muted, line_spacing=1.25)
    add_rect(slide, ctx, M, BODY_Y - 0.12, CW - 2 * M, 0.03, t.line)
    return {"title": spec.get("title", ""), "fill": None}


def s_detail(slide, ctx, spec, idx, warn):
    """导语 + 2–4 个并列小块的阅读型页：讲清一个判断的几条理由。"""
    t = ctx.theme
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    blocks = spec.get("blocks") or []
    if not 1 <= len(blocks) <= 4:
        warn.append(("W", "detail-count", f"并列块 {len(blocks)} 个，建议 2–4 个"))
    lead_h = 0.0
    lead = spec.get("lead")
    if lead:
        lsize, _ = fit_size([str(lead)], CW - 2 * M - 0.1, 0.9, 17, 13)
        add_text(slide, ctx, M, BODY_Y + 0.02, CW - 2 * M, 0.92, str(lead),
                 size=lsize, color=t.accent, line_spacing=1.34, bold=False)
        lead_h = 1.02
    n = max(1, len(blocks))
    gap = 0.3
    bw = (CW - 2 * M - gap * (n - 1)) / n
    y = BODY_Y + 0.04 + lead_h
    bh = BODY_H - 0.12 - lead_h
    fill = None
    for i, b in enumerate(blocks):
        xx = M + i * (bw + gap)
        add_rect(slide, ctx, xx, y, bw, bh, t.band)
        add_rect(slide, ctx, xx, y, bw, 0.05, t.accent2)
        head = str(b.get("title", ""))
        hsize, _ = fit_size([head], bw - 0.56, 0.55, 16, 12)
        add_text(slide, ctx, xx + 0.28, y + 0.28, bw - 0.56, 0.58, head,
                 size=hsize, bold=True, color=t.accent, line_spacing=1.15)
        body = str(b.get("body", ""))
        bsize, bfill = fit_size([body], bw - 0.56, bh - 1.12, 13.5, 10.5)
        add_text(slide, ctx, xx + 0.28, y + 0.92, bw - 0.56, bh - 1.12, body,
                 size=bsize, color=t.ink, line_spacing=1.3)
        fill = bfill if fill is None else max(fill, bfill)
    return {"title": spec.get("title", ""), "fill": fill}


def s_cards(slide, ctx, spec, idx, warn):
    """卡片页：2–4 张卡，每张一个标签+结论+两三行说明，信息量比纯要点大。"""
    t = ctx.theme
    cards = spec.get("cards") or []
    n = len(cards)
    if not 2 <= n <= 4:
        warn.append(("W", "cards-count", f"卡片 {n} 张，建议 2–4 张"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    gap = 0.34
    if n >= 3:
        cw = (CW - 2 * M - gap * (n - 1)) / n
        ch = BODY_H - 0.42
        cols = n
    else:
        cw = (CW - 2 * M - gap) / 2
        ch = BODY_H - 0.42
        cols = 2
    if n == 4:
        cw = (CW - 2 * M - gap) / 2
        ch = (BODY_H - 0.44 - gap) / 2
        cols = 2
    fill = None
    pad = 0.26
    for i, c in enumerate(cards):
        col, row = i % cols, i // cols
        xx = M + col * (cw + gap)
        yy = BODY_Y + 0.2 + row * (ch + gap)
        add_rect(slide, ctx, xx, yy, cw, ch, t.band)
        add_rect(slide, ctx, xx, yy, cw, 0.05, t.accent2)
        top = yy + pad
        label = c.get("label")
        if label:
            add_text(slide, ctx, xx + pad, top, cw - 2 * pad, 0.28, str(label),
                     size=10.5, bold=True, color=t.accent2, char_spacing=1.2)
            top += 0.30
        head = str(c.get("title", ""))
        hsize, _ = fit_size([head], cw - 2 * pad, 0.6, 16, 12)
        add_text(slide, ctx, xx + pad, top, cw - 2 * pad, 0.62, head, size=hsize,
                 bold=True, color=t.ink, line_spacing=1.14)
        top += 0.64
        body = str(c.get("body", ""))
        avail = yy + ch - top - 0.14
        bsize, bfill = fit_size([body], cw - 2 * pad, avail, 12.5, 9.5)
        if bfill > 1:
            warn.append(("W", "cards-overflow", f"卡片文字溢出：{head[:14]}"))
        add_text(slide, ctx, xx + pad, top, cw - 2 * pad, avail, body, size=bsize,
                 color=t.ink, line_spacing=1.26)
        fill = bfill if fill is None else max(fill, bfill)
    return {"title": spec.get("title", ""), "fill": fill}


def s_process(slide, ctx, spec, idx, warn):
    """流程页：3–5 步横向推进，每步带验收标准，比时间轴更适合"我要怎么做"。"""
    t = ctx.theme
    steps = spec.get("steps") or []
    n = len(steps)
    if not 3 <= n <= 5:
        warn.append(("W", "process-count", f"流程 {n} 步，建议 3–5 步"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    gap = 0.22
    w = (CW - 2 * M - gap * (n - 1)) / n
    y = BODY_Y + 0.22
    h = BODY_H - 0.4
    fill = None
    for i, st in enumerate(steps):
        xx = M + i * (w + gap)
        add_rect(slide, ctx, xx, y, w, h, t.band)
        add_rect(slide, ctx, xx, y, w, 0.05, t.accent2)
        add_rect(slide, ctx, xx + 0.26, y + 0.3, 0.52, 0.52, t.accent,
                 shape=MSO_SHAPE.OVAL)
        add_text(slide, ctx, xx + 0.26, y + 0.36, 0.52, 0.42, str(i + 1),
                 size=15, bold=True, color=hexc("#FFFFFF"), align=PP_ALIGN.CENTER)
        head = str(st.get("title", ""))
        hsize, _ = fit_size([head], w - 0.52, 0.62, 15.5, 12)
        add_text(slide, ctx, xx + 0.26, y + 0.98, w - 0.52, 0.66, head,
                 size=hsize, bold=True, color=t.ink, line_spacing=1.15)
        detail = str(st.get("detail", ""))
        avail = h - 1.78
        dsize, dfill = fit_size([detail], w - 0.52, avail, 12.5, 10)
        add_text(slide, ctx, xx + 0.26, y + 1.66, w - 0.52, avail, detail,
                 size=dsize, color=t.muted, line_spacing=1.3)
        fill = dfill if fill is None else max(fill, dfill)
        if i < n - 1:
            ax = xx + w + 0.03
            add_rect(slide, ctx, ax, y + h / 2 - 0.13, 0.16, 0.26, t.accent2,
                     shape=MSO_SHAPE.ISOSCELES_TRIANGLE).rotation = 90
    return {"title": spec.get("title", ""), "fill": fill}


def s_matrix(slide, ctx, spec, idx, warn):
    """2×2 矩阵：把一个判断拆成两个轴，适合"岗位要什么 vs 我有什么"。"""
    t = ctx.theme
    quads = spec.get("quadrants") or []
    if len(quads) != 4:
        warn.append(("W", "matrix-count", f"四象限需要 4 块，当前 {len(quads)} 块"))
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    ax, ay = 1.0, 0.36
    top = BODY_Y + ay
    gw = CW - 2 * M - ax - 0.1
    gh = BODY_H - ay - 0.42
    gap = 0.22
    cw = (gw - gap) / 2
    ch = (gh - gap) / 2
    fill = None
    for i, q in enumerate(quads[:4]):
        row, col = divmod(i, 2)
        xx = M + ax + col * (cw + gap)
        yy = top + row * (ch + gap)
        add_rect(slide, ctx, xx, yy, cw, ch, t.band)
        add_rect(slide, ctx, xx, yy, 0.05, ch, t.accent2)
        head = str(q.get("title", ""))
        add_text(slide, ctx, xx + 0.24, yy + 0.18, cw - 0.48, 0.5, head,
                 size=15, bold=True, color=t.accent, line_spacing=1.1)
        body = str(q.get("text", ""))
        avail = ch - 0.82
        bsize, bfill = fit_size([body], cw - 0.48, avail, 12.5, 9.5)
        add_text(slide, ctx, xx + 0.24, yy + 0.72, cw - 0.48, avail, body,
                 size=bsize, color=t.ink, line_spacing=1.26)
        fill = bfill if fill is None else max(fill, bfill)
    xl = spec.get("x_label")
    if xl:
        add_text(slide, ctx, M + ax, top + gh + 0.12, gw, 0.3, str(xl),
                 size=10.5, color=t.muted, align=PP_ALIGN.CENTER)
    yl = spec.get("y_label")
    if yl:
        add_text(slide, ctx, M, BODY_Y - 0.02, gw + ax, 0.3, str(yl),
                 size=10.5, color=t.muted)
    return {"title": spec.get("title", ""), "fill": fill}


def s_twocol(slide, ctx, spec, idx, warn):
    """左文右文 / 左文右图：阅读型 PPT 最常用的一页两块。"""
    t = ctx.theme
    chrome(slide, ctx, spec.get("kicker"), spec.get("title", ""), idx)
    left = spec.get("left") or {}
    right = spec.get("right") or {}
    img = spec.get("image")
    gap = 0.4
    lw = (CW - 2 * M - gap) / 2
    y = BODY_Y + 0.12
    h = BODY_H - 0.28
    fill = None

    def panel(x, w, data):
        nonlocal fill
        if not data:
            return
        add_rect(slide, ctx, x, y, w, h, t.band)
        head = str(data.get("title", ""))
        add_text(slide, ctx, x + 0.3, y + 0.24, w - 0.6, 0.55, head, size=16,
                 bold=True, color=t.accent, line_spacing=1.12)
        bullets = data.get("bullets") or []
        texts = [b if isinstance(b, str) else str(b.get("text", "")) for b in bullets]
        if not texts:
            body = str(data.get("body", ""))
            texts = [body] if body else []
        avail = h - 1.06
        size = 13.5
        while size > 10 and sum(est_block_height([x_], w - 0.72, size) + 0.09
                                for x_ in texts) > avail:
            size -= 1
        yy = y + 0.92
        for i, txt in enumerate(texts):
            add_rect(slide, ctx, x + 0.3, yy + 0.11, 0.22, 0.045, t.accent2)
            add_text(slide, ctx, x + 0.66, yy, w - 0.96,
                     est_block_height([txt], w - 0.96, size) + 0.08, txt, size=size,
                     color=t.ink, line_spacing=1.26)
            yy += est_block_height([txt], w - 0.96, size) + 0.18
            if yy > y + h - 0.4:
                warn.append(("W", "twocol-overflow", f"两栏页文字溢出：{spec.get('title')}"))
                break
        used = (yy - (y + 0.92)) / max(0.01, avail)
        fill = used if fill is None else max(fill, used)

    panel(M, lw, left)
    if img:
        panel(M + lw + gap, lw, None)
        try:
            pic = slide.shapes.add_picture(img, ctx.x(M + lw + gap + 0.2),
                                           ctx.y(y + 0.2), width=ctx.w(lw - 0.4))
            if pic.height > ctx.h(h - 0.4):
                ratio = (h - 0.4) * 914400 * ctx.sy / pic.height
                pic.height = int(pic.height * ratio)
                pic.width = int(pic.width * ratio)
            pic.left = ctx.x(M + lw + gap + (lw - pic.width / 914400 / ctx.sx) / 2)
        except Exception:
            warn.append(("W", "image-missing", f"图片缺失：{img}"))
    else:
        panel(M + lw + gap, lw, right)
    return {"title": spec.get("title", ""), "fill": fill}


def s_evidence(slide, ctx, spec, idx, warn):
    """证据页：编号—证据—证明什么—出处，四列对齐，评委可以逐条核。"""
    items = spec.get("items") or []
    if not items:
        warn.append(("E", "evidence-empty", "evidence 页没有任何证据"))
        return {"title": spec.get("title", ""), "fill": 0.0}
    if len(items) > 7:
        warn.append(("W", "evidence-count", f"证据 {len(items)} 条，建议 ≤7（单页装不下）"))
    rows = []
    for i, it in enumerate(items, 1):
        rows.append([f"{i:02d}", str(it.get("name", "")),
                     str(it.get("prove", "")), str(it.get("source", ""))])
    spec2 = {"kicker": spec.get("kicker"), "title": spec.get("title", ""),
             "headers": ["#", "证据", "证明什么", "出处"],
             "rows": rows, "note": spec.get("note")}
    return s_table(slide, ctx, spec2, idx, warn)


def overlap_check(boxes, i) -> list[tuple[str, str, str]]:
    out = []
    for a in range(len(boxes)):
        x1, y1, w1, h1, t1 = boxes[a]
        if not t1.strip():
            continue
        for b in range(a + 1, len(boxes)):
            x2, y2, w2, h2, t2 = boxes[b]
            if not t2.strip():
                continue
            ix = max(0.0, min(x1 + w1, x2 + w2) - max(x1, x2))
            iy = max(0.0, min(y1 + h1, y2 + h2) - max(y1, y2))
            small = min(w1 * h1, w2 * h2)
            if small > 0 and (ix * iy) / small > 0.22:
                out.append(("W", "overlap",
                            "第 %d 页疑似压字：『%s』×『%s』" % (i, t1[:12], t2[:12])))
    return out


BUILDERS = {
    "cover": s_cover, "agenda": s_agenda, "section": s_section,
    "bullets": s_bullets, "detail": s_detail, "cards": s_cards, "kpi": s_kpi,
    "timeline": s_timeline, "process": s_process, "compare": s_compare,
    "matrix": s_matrix, "table": s_table, "evidence": s_evidence,
    "twocol": s_twocol, "image": s_image, "quote": s_quote, "closing": s_closing,
}


# ---------------------------------------------------------------- 校验与装配

def validate(outline: dict) -> list[tuple[str, str, str]]:
    errs = []
    if not isinstance(outline.get("slides"), list) or not outline["slides"]:
        errs.append(("E", "no-slides", "outline.json 里没有 slides 数组"))
        return errs
    for i, s in enumerate(outline["slides"], 1):
        if "type" not in s:
            errs.append(("E", "no-type", f"第 {i} 页缺少 type"))
            continue
        if s["type"] not in BUILDERS:
            errs.append(("E", "bad-type", f"第 {i} 页 type={s['type']} 不支持，可选：{', '.join(BUILDERS)}"))
    return errs


SKIP_KEYS = {"notes", "id", "path", "image", "dim"}


def spec_chars(spec) -> int:
    """一页的可见文字量（不含备注/路径/维度标注）。"""
    def walk(v):
        if isinstance(v, str):
            return len("".join(v.split()))
        if isinstance(v, (int, float)):
            return 0
        if isinstance(v, dict):
            return sum(walk(x) for k, x in v.items() if k not in SKIP_KEYS)
        if isinstance(v, list):
            return sum(walk(x) for x in v)
        return 0
    return walk(spec)


def pick_layout(prs, kind: str):
    names = {"cover": ("标题", "title", "封面"), "section": ("节", "section"),
             "bullets": ("内容", "content", "两栏"), "closing": ("标题", "title")}
    pool = names.get(kind, ("内容", "content"))
    best, best_score = None, -1
    for lay in prs.slide_layouts:
        nm = (lay.name or "").lower()
        score = sum(1 for k in pool if k.lower() in nm)
        if score > best_score:
            best, best_score = lay, score
    if best is None:
        best = prs.slide_layouts[min(1, len(prs.slide_layouts) - 1)]
    return best


def build(outline: dict, out_path: str, template: str | None = None, save: bool = True):
    theme = Theme(outline.get("theme_spec") or {})
    if isinstance(outline.get("theme"), str) and not outline.get("theme_spec"):
        theme = Theme({"name": outline["theme"], **DEFAULT_THEMES.get(outline["theme"], {})})
    slides = outline["slides"]
    if template:
        prs = Presentation(template)
        # 模板里若自带示例页，先清空，只保留母版/版式/主题
        stale = len(prs.slides)
        if stale:
            sld_ids = prs.slides._sldIdLst
            for sid in list(sld_ids):
                sld_ids.remove(sid)
            print(f"! 模板含 {stale} 页示例内容，已清空（母版与版式保留）", file=sys.stderr)
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(CW), Inches(CH)
    grp = outline.get("group") or "Z1"
    label = outline.get("group_label") or GROUPS.get(grp, GROUPS["Z1"])[0]
    ctx = Ctx(prs, theme, label, len(slides))
    ctx.total = len(slides)

    rows, warnings = [], []
    for i, spec in enumerate(slides, 1):
        warn: list[tuple[str, str, str]] = []
        kind = spec["type"]
        layout = pick_layout(prs, kind) if template else prs.slide_layouts[6]
        slide = prs.slides.add_slide(layout)
        if template:
            # 只借模板的母版/背景/主题色，不用它的占位符（我们的版式与它不对应，留着会重复出字）
            for ph in list(slide.placeholders):
                ph._element.getparent().remove(ph._element)
        ctx.boxes = []
        info = BUILDERS[kind](slide, ctx, spec, i, warn)
        warn.extend(overlap_check(ctx.boxes, i))
        if spec.get("notes"):
            slide.notes_slide.notes_text_frame.text = str(spec["notes"])
        title_len = len(info.get("title") or "")
        if title_len > 26:
            warn.append(("W", "title-long", f"标题 {title_len} 字，建议 ≤26：{info['title']}"))
        fv = info.get("fill")
        rows.append({"index": i, "type": kind, "title": info.get("title", ""),
                     "fill": round(float(fv), 2) if isinstance(fv, (int, float)) else None,
                     "dim": spec.get("dim", ""), "chars": spec_chars(spec),
                     "notes_chars": len("".join(str(spec.get("notes", "")).split())),
                     "warnings": warn})
        warnings.extend(warn)
    if save:
        prs.save(out_path)
    return rows, warnings


def audit(outline: dict, rows: list, template: str | None):
    """页数、维度覆盖、缺口、占比等结构性检查。"""
    issues = []
    group = outline.get("group") or "Z1"
    label, dims, (lo, hi) = GROUPS.get(group, GROUPS["Z1"])
    n = len(rows)
    if not lo <= n <= hi:
        issues.append(("W", "page-count", f"{label} 建议 {lo}–{hi} 页，当前 {n} 页"))
    covered = {r["dim"] for r in rows if r.get("dim")}
    missing = [d for d in dims if d not in covered]
    if missing:
        issues.append(("W", "dim-missing", f"未标注以下评分维度：{'、'.join(missing)}"))
    CONTENT = {"bullets", "detail", "cards", "process", "matrix", "table",
               "evidence", "twocol", "timeline", "compare"}
    thin = [r["index"] for r in rows if r["type"] in CONTENT and r.get("chars", 0) < 90]
    dense = [r["index"] for r in rows if r["type"] in CONTENT and r.get("chars", 0) > 520]
    if thin:
        issues.append(("W", "page-thin",
                       f"第 {'、'.join(map(str, thin))} 页文字量偏少（<90 字）：阅读型 PPT 要求每页能独立读懂"))
    if dense:
        issues.append(("W", "page-dense",
                       f"第 {'、'.join(map(str, dense))} 页文字量偏大（>520 字）：考虑拆页或删细节"))
    # 备注口播量：只作提示，不影响版面判定
    SPEED = 269  # 字/分
    notes_total = sum(r.get("notes_chars", 0) for r in rows)
    target = 1750 if group in ("Z1", "Z2") else 1500
    minutes = notes_total / SPEED
    if notes_total < target * 0.75:
        issues.append(("W", "notes-short",
                       f"备注合计 {notes_total} 字 ≈ {int(minutes)}′{int((minutes % 1) * 60):02d}″，"
                       f"距 {'7′' if group in ('Z1', 'Z2') else '6′'} 目标还差约 {max(0, target - notes_total)} 字——"
                       "逐字稿阶段补，不影响版面"))
    empty = [r["index"] for r in rows if r.get("notes_chars", 0) == 0]
    if empty:
        issues.append(("W", "notes-missing",
                       f"第 {'、'.join(map(str, empty))} 页备注为空：备注栏放口播要点，现场提示用"))
    types = [r["type"] for r in rows]
    for need in ("cover", "closing"):
        if need not in types:
            issues.append(("E", f"missing-{need}", f"缺少 {need} 页"))
    if types.count("bullets") > n * 0.6:
        issues.append(("W", "too-many-bullets", "纯要点页占比过高，建议换成 kpi/timeline/compare/table"))
    for g in outline.get("gaps") or []:
        issues.append(("G", "material-gap", str(g)))
    return issues


def write_report(path, outline, rows, warnings, issues, out_path):
    lines = [f"# PPT 生成审计：{os.path.basename(out_path or '（未生成）')}", ""]
    lines.append(f"- 主题：`{outline.get('theme', 'academic')}`　组别：`{outline.get('group', 'Z1')}`　页数：{len(rows)}")
    lines.append(f"- 模板：{'自带模板' if outline.get('_template') else '内置版式'}")
    lines.append("")
    lines.append("| # | 类型 | 标题 | 文字紧度 | 字数 | 维度 | 告警 |")
    lines.append("| --: | :-- | :-- | --: | --: | :-- | :-- |")
    for r in rows:
        seen = []
        for item in r["warnings"]:
            if item[2] not in [x[2] for x in seen]:
                seen.append(item)
        r["warnings"] = seen
        w = "；".join(f"[{lv}] {msg}" for lv, _, msg in r["warnings"]) or "—"
        fill_txt = f"{r['fill']:.0%}" if isinstance(r["fill"], (int, float)) else "—"
        lines.append(f"| {r['index']} | {r['type']} | {r['title'][:30]} | {fill_txt} | {r.get('chars', 0)} | {r.get('dim') or '—'} | {w} |")
    lines.append("")
    lines.append("## 结构检查")
    if issues:
        for lv, code, msg in issues:
            lines.append(f"- [{lv}] **{code}**：{msg}")
    else:
        lines.append("- 无问题")
    fills = [r["fill"] for r in rows if isinstance(r["fill"], (int, float))]
    fill_max = max(fills) if fills else 0
    lines.append("")
    notes_total = sum(r.get("notes_chars", 0) for r in rows)
    mins = notes_total / 269
    lines.append(f"备注口播：合计 {notes_total} 字 ≈ {int(mins)}′{int((mins % 1) * 60):02d}″"
                 f"（269 字/分；6′ 需 1500–1650 字，7′ 需 1750–1900 字）")
    lines.append(f"文字紧度峰值：{fill_max:.0%}（>100% 表示缩到下限仍溢出，必须删字）")
    lines.append("注：紧度 = 该页最紧文字块占用其可用空间的比例；kpi/table/image 页不参与计算，显示 —。")
    open(path, "w", encoding="utf-8").write("\n".join(lines) + "\n")


PREVIEW_FALLBACKS = ["PingFang SC", "Heiti SC", "Noto Sans CJK SC",
                     "Source Han Sans SC", "SimHei", "Arial Unicode MS"]


def local_fonts() -> set[str] | None:
    """本机可用字体名集合；拿不到 fc-list 时返回 None（表示无法判断）。"""
    if not shutil.which("fc-list"):
        return None
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True,
                             text=True, timeout=20).stdout
    except Exception:
        return None
    fams: set[str] = set()
    for line in out.splitlines():
        for part in line.split(","):
            if part.strip():
                fams.add(part.strip())
    return fams


def preview_font_for(font: str) -> str | None:
    """PDF 预览用字体：本机没有原字体时，挑一个装得上的，避免预览件被替换成宋体/楷体。"""
    fams = local_fonts()
    if fams is None or font in fams:
        return None
    for cand in PREVIEW_FALLBACKS:
        if cand in fams:
            return cand
    return None


def convert_pdf(pptx_path: str) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                    "--outdir", os.path.dirname(os.path.abspath(pptx_path)), pptx_path],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.exists(os.path.splitext(pptx_path)[0] + ".pdf")


def main():
    ap = argparse.ArgumentParser(description="outline.json → 职规赛展示 PPTX")
    ap.add_argument("outline")
    ap.add_argument("-o", "--out", default=None, help="输出 pptx 路径")
    ap.add_argument("--template", default=None, help="用户自带 .pptx 模板")
    ap.add_argument("--theme", default=None, choices=sorted(DEFAULT_THEMES))
    ap.add_argument("--font", default=None, help="中文字体名，默认微软雅黑")
    ap.add_argument("--report", default=None, help="审计报告 markdown 路径")
    ap.add_argument("--pdf", action="store_true", help="尝试用 LibreOffice 导出 PDF")
    ap.add_argument("--audit-only", action="store_true")
    args = ap.parse_args()

    with open(args.outline, encoding="utf-8") as f:
        outline = json.load(f)

    errs = validate(outline)
    fatal = [e for e in errs if e[0] == "E"]
    if fatal:
        for lv, code, msg in fatal:
            print(f"[{lv}] {code}: {msg}", file=sys.stderr)
        sys.exit(2)

    out = args.out or os.path.splitext(args.outline)[0] + ".pptx"
    if args.theme:
        outline["theme"] = args.theme
    if args.font:
        outline.setdefault("theme_spec", {})["font"] = args.font
    outline["_template"] = bool(args.template)
    rows, warnings = build(outline, out, args.template, save=not args.audit_only)
    issues = audit(outline, rows, args.template)
    report = args.report or os.path.splitext(out)[0] + "-audit.md"
    write_report(report, outline, rows, warnings, issues, out if not args.audit_only else None)

    hard = [w for w in warnings if w[0] == "E"] + [x for x in issues if x[0] == "E"]
    soft = [w for w in warnings if w[0] != "E"] + [x for x in issues if x[0] != "E"]
    if args.audit_only:
        print(f"✓ 审计完成（{len(rows)} 页，未写 pptx）")
    else:
        print(f"✓ 生成 {out}（{len(rows)} 页）")
    print(f"✓ 审计报告 {report}")
    if soft:
        print(f"! {len(soft)} 条提示（见报告）")
    for lv, code, msg in hard:
        print(f"✗ [{code}] {msg}", file=sys.stderr)

    if args.pdf and not args.audit_only:
        if not (shutil.which("soffice") or shutil.which("libreoffice")):
            print("! 未找到 LibreOffice，跳过 PDF", file=sys.stderr)
        else:
            cur_font = (outline.get("theme_spec") or {}).get("font") or DEFAULT_FONT
            alt = preview_font_for(cur_font)
            if alt:
                print(f"! 本机未装 {cur_font}，PDF 预览改用 {alt}（只影响预览件，pptx 字体不变）",
                      file=sys.stderr)
                tmp = os.path.splitext(out)[0] + ".preview.pptx"
                o2 = json.loads(json.dumps(outline))
                o2.setdefault("theme_spec", {})["font"] = alt
                build(o2, tmp, args.template)
                if convert_pdf(tmp):
                    os.replace(os.path.splitext(tmp)[0] + ".pdf",
                               os.path.splitext(out)[0] + ".pdf")
                    print(f"✓ 导出 PDF：{os.path.splitext(out)[0]}.pdf（预览字体 {alt}）")
                if os.path.exists(tmp):
                    os.remove(tmp)
            else:
                convert_pdf(out)
                print(f"✓ 导出 PDF：{os.path.splitext(out)[0]}.pdf")
    if hard:
        sys.exit(1)


if __name__ == "__main__":
    main()
