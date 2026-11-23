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
    "Z1": ("成长赛道·高教组", ["职业目标", "学习实践行动", "优化改进"], (15, 20)),
    "Z2": ("成长赛道·职教组", ["职业目标", "学习实践行动", "优化改进"], (15, 20)),
    "Z3": ("就业赛道·本科组", ["职业目标", "岗位胜任力", "发展潜力"], (15, 20)),
    "Z4": ("就业赛道·研究生组", ["职业目标", "岗位胜任力", "发展潜力"], (15, 20)),
    "Z5": ("就业赛道·职教组", ["职业目标", "岗位胜任力", "发展潜力"], (15, 20)),
}

# 备注口播的目标字数区间（字/分 = 269）。低于下限或高于上限都要报。
NOTES_RANGE = {"Z1": (1750, 1900), "Z2": (1750, 1900),
               "Z3": (1500, 1650), "Z4": (1500, 1650), "Z5": (1500, 1650)}
SPEED_CPM = 269

# 每页型的分级字数上限（与 references/slide-types.md §1 表格逐项对应）。
# 超限只报 W：版面会自适应缩字号，但阅读型 PPT 的可读性靠的是「字数先守住」。
CAPS: dict[str, dict[str, int]] = {
    "cover": {"title": 20},
    "section": {"title": 24},
    "closing": {"title": 22},
    "quote": {"text": 40},
    "image": {"caption": 30},
    "agenda": {"item_title": 40, "item_note": 40},
    "bullets": {"item_text": 48, "item_note": 40},
    "detail": {"lead": 60, "block_body": 90},
    "cards": {"card_body": 70},
    "kpi": {"value": 8, "label": 14},
    "timeline": {"what": 36, "detail": 18},
    "process": {"detail": 70},
    "compare": {"item": 28},
    "matrix": {"quad_text": 70},
    "evidence": {"name": 30, "prove": 30},
    "twocol": {"item": 30},
}

# 不会出现在版面上的键：字数统计时必须排除（列黑名单而非白名单，
# 未知字段一律计入，宁可保守也不要漏算而让 page-thin 假通过）
NON_VISIBLE_KEYS = {"notes", "id", "path", "image", "dim", "_template", "em", "no"}


# 内容页（参与「薄/满」与底部空白判定的页型）
CONTENT_TYPES = {"bullets", "detail", "cards", "process", "matrix", "table",
                 "evidence", "twocol", "timeline", "compare"}



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
    s = str(s).strip().lstrip("#")
    if len(s) != 6 or any(c not in "0123456789abcdefABCDEF" for c in s):
        raise ValueError(f"颜色必须是 6 位十六进制（如 #1F4E79），收到 {s!r}")
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

    # 4:3 模板的缩放系数是 0.75，10.5pt 的注释会被压到 7.9pt（评委看不清）。
    # 因此设 9.5pt 绝对下限：宁可略密，也不要把小字缩到读不出来。
    MIN_RENDER_PT = 9.5

    def pt(self, v: float) -> Pt:
        return Pt(max(self.MIN_RENDER_PT, v * self.fs))


def add_text(slide, ctx: Ctx, x, y, w, h, text, *, size=18, bold=False,
             color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.18, space_after=4, font=None, italic=False,
             char_spacing=None):
    # 登记的是「文字实际占的高度」而不是容器高度：容器常被撑到整栏，
    # 用容器量版面跨度会把「只有一行字的三卡页」误判成内容饱满（F6 的根因）。
    lines = text.split("\n") if isinstance(text, str) else list(text)
    ink_h = min(h, est_block_height(lines, w, size))
    ink_y = y + (h - ink_h) / 2 if anchor == MSO_ANCHOR.MIDDLE else y
    ctx.boxes.append((x, ink_y, w, ink_h, text if isinstance(text, str) else " ".join(text)))
    box = slide.shapes.add_textbox(ctx.x(x), ctx.y(y), ctx.w(w), ctx.h(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
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
    # 大数字用水印色而不是 band：band 在白底上对比度只有约 1.05:1，等于看不见（实测 F7）
    add_text(slide, ctx, M - 0.08, 1.28, 3.2, 1.55, str(num), size=78, bold=True,
             color=t.muted if t.muted != t.bg else hexc("#B9C3D0"))
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
    # 每张卡的注释按实际字数定高（原来固定 0.36in，长口径会溢出且无告警）
    note_h = 0.0
    for it in items:
        note = it.get("note")
        if note:
            sz, _ = fit_size([str(note)], inner, 0.9, 10.5, 9)
            note_h = max(note_h, est_block_height([str(note)], inner, sz) + 0.06)
    card_h = 0.28 + value_h + 0.20 + label_h + (note_h + 0.10 if note_h else 0) + 0.20
    # 上限抬到 3.40：4 张卡时原来 clamp 到 2.95 会留下整屏空白（实测 F6）
    card_h = max(2.05, min(max(2.95, note_h + value_h + label_h + 0.9), 3.40))
    # 注释按卡片实际可用高度复算：2 行以上或超出卡高就报 kpi-note-overflow（文档承诺的信号）
    note_sz = 10.5
    for it in items:
        note = it.get("note")
        if not note:
            continue
        avail_note = card_h - (0.28 + value_h + 0.20 + label_h + 0.08 + 0.20)
        nlines = est_lines(str(note), inner, note_sz)
        if nlines > 2 or est_block_height([str(note)], inner, note_sz) > avail_note:
            warn.append(("W", "kpi-note-overflow",
                         f"数字卡注释 {nlines} 行 / 超出卡高：{str(note)[:18]}"
                         "（压到 2 行内，或把长口径移到 source 页脚）"))

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
            add_text(slide, ctx, xx + 0.3, label_y + label_h + 0.08, inner,
                     max(note_h, 0.36), str(note), size=10.5, color=t.muted)
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
        size, fill = fit_size([f"· {x}" for x in items], pw - 0.56, ph - 0.85, 15.5, 10.5)
        # 竖直居中：两栏各 3–4 条的页面原来「文字贴顶、下半屏空」，看着像空壳
        add_text(slide, ctx, xx + 0.28, top + 0.78, pw - 0.56, ph - 0.95,
                 "\n".join(f"· {x}" for x in items), size=size, color=t.ink,
                 line_spacing=1.25, space_after=7, anchor=MSO_ANCHOR.MIDDLE)
        if fill > 1:
            warn.append(("W", "compare-overflow", f"对比栏溢出：{p.get('head')}"))
    if heads:
        warn.append(("W", "compare-items", "；".join(heads) + "，建议每栏 ≤4 条"))
    cx = M + pw + 0.75
    add_rect(slide, ctx, cx - 0.26, top + ph / 2 - 0.26, 0.52, 0.52, t.accent2,
             shape=MSO_SHAPE.OVAL)
    add_text(slide, ctx, cx - 0.26, top + ph / 2 - 0.16, 0.52, 0.32, "VS",
             size=12, bold=True, color=hexc("#FFFFFF"), align=PP_ALIGN.CENTER)
    src = spec.get("source") or spec.get("note")
    if src:
        # 来源必须上版面：口径是评委追问的落点，也是本技能「数字要能查」的底线。
        # note 里常已自带「来源：」，避免叠加成「数据来源：来源：…」
        txt = str(src).strip()
        for prefix in ("数据来源：", "数据来源:", "来源：", "来源:"):
            if txt.startswith(prefix):
                txt = txt[len(prefix):].strip()
                break
        add_text(slide, ctx, M, FOOT_Y - 0.42, CW - 2 * M, 0.3, f"数据来源：{txt}",
                 size=10, color=t.muted)
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
    top_y = BODY_Y + 0.12
    shape = slide.shapes.add_table(nrow, ncol, ctx.x(M), ctx.y(top_y),
                                   ctx.w(width), ctx.h(row_h * nrow))
    # 表格/图片是图形不是文本框，也要登记占地范围，否则「版面跨度」会误判成空页
    ctx.boxes.append((M, top_y, width, row_h * nrow, "[table]"))
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
        ctx.boxes.append((x, y, dw, dh, "[picture]"))
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
                 size=bsize, color=t.ink, line_spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
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
        # 卡片高度由页型定，正文垂直居中，避免「文字贴顶 + 下方整片空」被当成空壳页
        add_text(slide, ctx, xx + pad, top, cw - 2 * pad, avail, body, size=bsize,
                 color=t.ink, line_spacing=1.26, anchor=MSO_ANCHOR.MIDDLE)
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
                 size=dsize, color=t.muted, line_spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
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
                 size=bsize, color=t.ink, line_spacing=1.26, anchor=MSO_ANCHOR.MIDDLE)
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
        size = 15.5          # 从小字号起步会让 3 条要点只占三分之一，白白留空
        while size > 10 and sum(est_block_height([x_], w - 0.72, size) + 0.09
                                for x_ in texts) > avail:
            size -= 1
        block = sum(est_block_height([x_], w - 0.72, size) + 0.18 for x_ in texts)
        yy = y + 0.92 + max(0.0, (avail - block) / 2)   # 竖直居中，留白对称
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

    src = spec.get("source") or spec.get("note")
    if src:
        add_text(slide, ctx, M, FOOT_Y - 0.42, CW - 2 * M, 0.3, f"数据来源：{src}",
                 size=10, color=t.muted)
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

def _is_str_list(v) -> bool:
    """数组里每一项都是字符串（或可转成字符串的数字）。"""
    return isinstance(v, list) and all(isinstance(x, (str, int, float)) for x in v)


class WHOLE:
    """标记：该规则校验整个数组值，而不是逐个元素。"""
    def __init__(self, fn):
        self.fn = fn


def _check_list_field(errs, i, spec, key, *, required=True, item_check=None, code=""):
    if key not in spec or spec.get(key) is None:
        if required:
            errs.append(("E", code or f"field-missing",
                         f"第 {i} 页（{spec.get('type')}）缺少 {key}"))
        return None
    val = spec[key]
    if not isinstance(val, list):
        errs.append(("E", code or "field-type",
                     f"第 {i} 页（{spec.get('type')}）的 {key} 应为数组，"
                     f"实际是 {type(val).__name__}：{str(val)[:40]!r}"))
        return None
    if not val:
        errs.append(("E", code or "field-missing",
                     f"第 {i} 页（{spec.get('type')}）的 {key} 是空数组"))
        return None
    if isinstance(item_check, WHOLE):
        msg = item_check.fn(val)
        if msg:
            errs.append(("E", code or "field-type", f"第 {i} 页 {key} {msg}"))
        return val
    if item_check:
        for j, item in enumerate(val, 1):
            msg = item_check(item)
            if msg:
                errs.append(("E", code or "field-type",
                             f"第 {i} 页 {key}[{j}] {msg}"))
    return val


def _str_item(**fields):
    """生成「元素必须是 str 或带指定字符串字段的 dict」的校验器。"""
    def check(item):
        if isinstance(item, (str, int, float)):
            return ""
        if not isinstance(item, dict):
            return f"应为字符串或对象，实际是 {type(item).__name__}"
        for f in fields:
            if f not in item:
                return f"缺少字段 {f}"
            if not isinstance(item[f], (str, int, float)):
                return f"的 {f} 应为字符串，实际是 {type(item[f]).__name__}"
        return ""
    return check


def _panel(v):
    """compare/twocol 的左右栏：接受 {head|title, items|bullets} 或直接给数组。"""
    if isinstance(v, list):
        return "" if all(isinstance(x, (str, int, float)) for x in v) else \
            "数组元素应为字符串"
    if not isinstance(v, dict):
        return f"应为对象（含 items/bullets）或字符串数组，实际是 {type(v).__name__}"
    for key in ("items", "bullets"):
        if key in v:
            if not isinstance(v[key], list):
                return f"的 {key} 应为数组，实际是 {type(v[key]).__name__}"
            if not all(isinstance(x, (str, int, float)) for x in v[key]):
                return f"的 {key} 数组元素应为字符串"
    if "items" not in v and "bullets" not in v and "body" not in v \
            and "image" not in v:
        return "缺少 items/bullets/body"
    return ""


SPEC_RULES = {
    "cover": [("title", "str")],
    "section": [("title", "str")],
    "closing": [("title", "str")],
    "quote": [("text", "str")],
    "agenda": [("items", _str_item(title="title"))],
    "bullets": [("bullets", _str_item(text="text"))],
    "detail": [("blocks", _str_item(title="title", body="body"))],
    "cards": [("cards", _str_item(title="title", body="body"))],
    "kpi": [("items", _str_item(value="value", label="label"))],
    "timeline": [("items", _str_item(when="when", what="what"))],
    "process": [("steps", _str_item(title="title"))],
    "matrix": [("quadrants", _str_item(title="title", text="text"))],
    # WHOLE 标记：校验整个数组本身，而不是逐个元素
    "table": [("headers", WHOLE(lambda v: "" if _is_str_list(v) else "应为字符串数组")),
              ("rows", WHOLE(lambda v: "" if isinstance(v, list)
                             and all(_is_str_list(r) for r in v)
                             else "应为「字符串数组」的数组"))],
    "evidence": [("items", _str_item(name="name", prove="prove"))],
}

VALID_THEMES = set(DEFAULT_THEMES)


def validate(outline: dict) -> list[tuple[str, str, str]]:
    errs: list[tuple[str, str, str]] = []
    if not isinstance(outline, dict):
        return [("E", "no-outline", "outline 顶层应为 JSON 对象")]
    if not isinstance(outline.get("slides"), list) or not outline["slides"]:
        errs.append(("E", "no-slides", "outline.json 里没有 slides 数组"))
        return errs

    grp = outline.get("group")
    if grp is None:
        errs.append(("W", "group-missing",
                     f"未写 group，按 {sorted(GROUPS)[0]} 处理，页数与维度清单可能不符"))
    elif grp not in GROUPS:
        errs.append(("E", "group-invalid",
                     f"group={grp!r} 不是 Z1–Z5（可选：{'、'.join(sorted(GROUPS))}）；"
                     "写错组别会让页数区间与评分维度全部失效"))
    if grp in GROUPS and outline.get("group_label"):
        want = GROUPS[grp][0]
        if str(outline["group_label"]).strip() != want:
            errs.append(("W", "label-mismatch",
                         f"group_label={outline['group_label']!r} 与 group={grp} 不符"
                         f"（应为 {want!r}）；页脚会印错赛道"))

    theme = outline.get("theme")
    spec = outline.get("theme_spec")
    if spec is not None and not isinstance(spec, dict):
        errs.append(("E", "theme-spec-invalid", "theme_spec 应为对象"))
    elif isinstance(spec, dict):
        for k in ("ink", "accent", "accent2", "muted", "band", "line", "bg"):
            if k in spec:
                try:
                    hexc(spec[k])
                except ValueError as e:
                    errs.append(("E", "theme-spec-invalid", f"theme_spec.{k}：{e}"))
    elif theme is not None and theme not in VALID_THEMES:
        errs.append(("E", "theme-invalid",
                     f"theme={theme!r} 不存在（可选：{'、'.join(sorted(VALID_THEMES))}）；"
                     "写错会被静默替换成 academic"))

    for i, s in enumerate(outline["slides"], 1):
        if not isinstance(s, dict):
            errs.append(("E", "no-type", f"第 {i} 页应为对象，实际是 {type(s).__name__}"))
            continue
        if "type" not in s:
            errs.append(("E", "no-type", f"第 {i} 页缺少 type"))
            continue
        kind = s["type"]
        if kind not in BUILDERS:
            errs.append(("E", "bad-type",
                         f"第 {i} 页 type={kind} 不支持，可选：{', '.join(BUILDERS)}"))
            continue
        for key, rule in SPEC_RULES.get(kind, []):
            if isinstance(rule, str):          # 必填字符串字段
                if key not in s or s[key] is None:
                    errs.append(("E", "field-missing",
                                 f"第 {i} 页（{kind}）缺少 {key}"))
                elif not isinstance(s[key], (str, int, float)):
                    errs.append(("E", "field-type",
                                 f"第 {i} 页（{kind}）的 {key} 应为字符串，"
                                 f"实际是 {type(s[key]).__name__}"))
                elif not str(s[key]).strip():
                    errs.append(("E", "field-empty",
                                 f"第 {i} 页（{kind}）的 {key} 是空字符串"))
            else:
                _check_list_field(errs, i, s, key, item_check=rule)
        if kind == "compare":
            for side in ("left", "right"):
                if side not in s:
                    errs.append(("E", "field-missing", f"第 {i} 页（compare）缺少 {side}"))
                    continue
                msg = _panel(s[side])
                if msg:
                    errs.append(("E", "field-type", f"第 {i} 页 compare.{side} {msg}"))
        if kind == "twocol":
            if not s.get("left") and not s.get("image"):
                errs.append(("E", "field-missing",
                             f"第 {i} 页（twocol）至少要给 left 或 image"))
            if s.get("left"):
                msg = _panel(s["left"])
                if msg:
                    errs.append(("E", "field-type", f"第 {i} 页 twocol.left {msg}"))
            if s.get("right"):
                msg = _panel(s["right"])
                if msg:
                    errs.append(("E", "field-type", f"第 {i} 页 twocol.right {msg}"))
        if kind == "image" and "path" in s and s["path"] is not None \
                and not isinstance(s["path"], str):
            errs.append(("E", "field-type",
                         f"第 {i} 页（image）的 path 应为字符串，"
                         f"实际是 {type(s['path']).__name__}"))
        if kind == "matrix" and isinstance(s.get("quadrants"), list) \
                and len(s["quadrants"]) != 4:
            errs.append(("W", "matrix-count",
                         f"第 {i} 页四象限需要 4 块，当前 {len(s['quadrants'])} 块"))
        dim = s.get("dim")
        if dim is not None:
            allowed = GROUPS.get(grp, GROUPS["Z1"])[1]
            if str(dim).strip() not in allowed:
                errs.append(("W", "dim-invalid",
                             f"第 {i} 页 dim={dim!r} 不在 {grp or 'Z1'} 的维度清单"
                             f"（{'、'.join(allowed)}）"))
    return errs


def spec_chars(spec) -> int:
    """一页真正会上版面的文字量（不含备注/路径/维度标注/未渲染字段）。"""
    def walk(v):
        if isinstance(v, str):
            return len("".join(v.split()))
        if isinstance(v, (int, float)):
            return 0
        if isinstance(v, dict):
            return sum(walk(x) for k, x in v.items() if k not in NON_VISIBLE_KEYS)
        if isinstance(v, list):
            return sum(walk(x) for x in v)
        return 0
    return walk({k: v for k, v in spec.items() if k not in NON_VISIBLE_KEYS})


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


def clear_template_slides(prs) -> int:
    """清空模板自带的示例页，只保留母版/版式/主题。

    关键点：只从 sldIdLst 摘掉条目是不够的——slide part 和它的 rel 仍在包里，
    python-pptx 新建页时会复用同名 part（slide1.xml 出现两份），
    结果是 zip 内重名、LibreOffice 拒载、PowerPoint 弹修复。必须 drop_rel。
    """
    stale = len(prs.slides)
    sld_ids = prs.slides._sldIdLst
    for sid in list(sld_ids):
        rId = sid.get(qn("r:id"))
        if rId:
            prs.part.drop_rel(rId)
        sld_ids.remove(sid)
    return stale


BODY_TOP_LIMIT = 1.90      # 版心起点（标题装饰线以下）：上方的 kicker/标题/色条不算正文
BODY_MIN_SPAN = 2.50       # 版心高 4.55in，正文竖直跨度低于 55% 视作「没撑起来」


def content_span(boxes) -> float | None:
    """正文在竖直方向的实际跨度（英寸）；只看版心内的文字框，排除页脚/来源行。"""
    body = [b for b in boxes
            if b[1] >= BODY_TOP_LIMIT and b[1] + b[3] <= FOOT_Y - 0.30 and b[4].strip()]
    if not body:
        return None
    top = min(b[1] for b in body)
    bottom = max(b[1] + b[3] for b in body)
    return round(bottom - top, 2)


def _cap(code: str, i: int, what: str, got: int, limit: int) -> tuple[str, str, str]:
    return ("W", f"cap-{code}",
            f"第 {i} 页 {what} {got} 字，超过 {limit} 字上限（阅读型靠字数先守，不要靠缩字号）")


def cap_warnings(kind: str, spec: dict, i: int) -> list[tuple[str, str, str]]:
    """按 CAPS 表逐项查字数上限（文档 §1 表格的代码实现）。"""
    out: list[tuple[str, str, str]] = []
    caps = CAPS.get(kind, {})

    def n(v) -> int:
        return len("".join(str(v).split())) if v is not None else 0

    def one(key, what, got, limit):
        if got > limit:
            out.append(_cap(key, i, what, got, limit))

    if kind in ("cover", "section", "closing", "quote"):
        one("title" if kind != "quote" else "text",
            {"cover": "标题", "section": "标题", "closing": "标题", "quote": "引语"}[kind],
            n(spec.get("title" if kind != "quote" else "text")), caps.get("title", caps.get("text", 10**6)))
    if kind == "image" and "caption" in caps:
        one("caption", "图片说明", n(spec.get("caption")), caps["caption"])
    if kind == "agenda" and "item_title" in caps:
        for j, it in enumerate(spec.get("items") or [], 1):
            if isinstance(it, dict):
                one("item-title", f"目录第 {j} 项标题", n(it.get("title")), caps["item_title"])
                one("item-note", f"目录第 {j} 项注释", n(it.get("note")), caps["item_note"])
    if kind == "bullets" and "item_text" in caps:
        for j, it in enumerate(spec.get("bullets") or [], 1):
            text = it.get("text") if isinstance(it, dict) else it
            one("bullet", f"第 {j} 条要点", n(text), caps["item_text"])
            if isinstance(it, dict):
                one("bullet-note", f"第 {j} 条小字", n(it.get("note")), caps["item_note"])
    if kind == "detail":
        one("lead", "导语", n(spec.get("lead")), caps.get("lead", 10**6))
        for j, b in enumerate(spec.get("blocks") or [], 1):
            if isinstance(b, dict):
                one("block-body", f"第 {j} 块正文", n(b.get("body")), caps.get("block_body", 10**6))
    if kind == "cards":
        for j, c in enumerate(spec.get("cards") or [], 1):
            if isinstance(c, dict):
                one("card-body", f"第 {j} 张卡正文", n(c.get("body")), caps.get("card_body", 10**6))
    if kind == "kpi":
        for j, it in enumerate(spec.get("items") or [], 1):
            if isinstance(it, dict):
                one("kpi-value", f"第 {j} 张卡数值", n(it.get("value")), caps.get("value", 10**6))
                one("kpi-label", f"第 {j} 张卡标签", n(it.get("label")), caps.get("label", 10**6))
    if kind == "timeline":
        for j, it in enumerate(spec.get("items") or [], 1):
            if isinstance(it, dict):
                one("timeline-what", f"第 {j} 个节点标题", n(it.get("what")), caps.get("what", 10**6))
                one("timeline-detail", f"第 {j} 个节点小字", n(it.get("detail")), caps.get("detail", 10**6))
    if kind == "process":
        for j, st in enumerate(spec.get("steps") or [], 1):
            if isinstance(st, dict):
                one("process-detail", f"第 {j} 步说明", n(st.get("detail")), caps.get("detail", 10**6))
    if kind in ("compare", "twocol"):
        key = "items" if kind == "compare" else "bullets"
        for side in ("left", "right"):
            v = spec.get(side)
            items = v.get(key) if isinstance(v, dict) else v
            for j, it in enumerate(items or [], 1):
                text = it.get("text") if isinstance(it, dict) else it
                one("item", f"{side} 第 {j} 条", n(text), caps.get("item", 10**6))
    if kind == "matrix":
        for j, q in enumerate(spec.get("quadrants") or [], 1):
            if isinstance(q, dict):
                one("quad", f"第 {j} 格正文", n(q.get("text")), caps.get("quad_text", 10**6))
    if kind == "evidence":
        for j, it in enumerate(spec.get("items") or [], 1):
            if isinstance(it, dict):
                one("evidence-name", f"第 {j} 条证据名", n(it.get("name")), caps.get("name", 10**6))
                one("evidence-prove", f"第 {j} 条证明什么", n(it.get("prove")), caps.get("prove", 10**6))
    return out


def build(outline: dict, out_path: str, template: str | None = None, save: bool = True):
    theme = Theme(outline.get("theme_spec") or {})
    if isinstance(outline.get("theme"), str) and not outline.get("theme_spec"):
        theme = Theme({"name": outline["theme"], **DEFAULT_THEMES.get(outline["theme"], {})})
    slides = outline["slides"]
    if template:
        prs = Presentation(template)
        stale = clear_template_slides(prs)
        if stale:
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
        warn.extend(cap_warnings(kind, spec, i))
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
                     "content_span": content_span(ctx.boxes),
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
    CONTENT = CONTENT_TYPES
    thin = [r["index"] for r in rows if r["type"] in CONTENT and r.get("chars", 0) < 90]
    dense = [r["index"] for r in rows if r["type"] in CONTENT and r.get("chars", 0) > 520]
    if thin:
        issues.append(("W", "page-thin",
                       f"第 {'、'.join(map(str, thin))} 页文字量偏少（<90 字）：阅读型 PPT 要求每页能独立读懂"))
    if dense:
        issues.append(("W", "page-dense",
                       f"第 {'、'.join(map(str, dense))} 页文字量偏大（>520 字）：考虑拆页或删细节"))
    # 版面底部空白：正文最低点远高于版心下沿，就是「空壳页」（不只看字数，看实际排到哪）
    empties = [r["index"] for r in rows
               if r["type"] in CONTENT
               and isinstance(r.get("content_span"), (int, float))
               and r["content_span"] < BODY_MIN_SPAN]
    if empties:
        issues.append(("W", "page-empty-bottom",
                       f"第 {'、'.join(map(str, empties))} 页正文没撑起版面"
                       f"（竖直跨度 <{BODY_MIN_SPAN}in，版心高 4.55in）："
                       "加一条带口径的小字、换成 detail/cards，或把同类小页合并"))
    # 备注口播量：区间两端都查（原来只查下限，「3000 字 ≈ 11′」这种超时写稿不会被拦）
    notes_total = sum(r.get("notes_chars", 0) for r in rows)
    lo_notes, hi_notes = NOTES_RANGE.get(group, NOTES_RANGE["Z1"])
    seconds = notes_total / SPEED_CPM * 60
    clock = f"{int(seconds // 60)}′{int(seconds % 60):02d}″"
    if notes_total < lo_notes:
        issues.append(("W", "notes-short",
                       f"备注合计 {notes_total} 字 ≈ {clock}，未达 {lo_notes} 字下限——"
                       "逐字稿阶段按页补，不影响版面"))
    elif notes_total > hi_notes:
        issues.append(("W", "notes-long",
                       f"备注合计 {notes_total} 字 ≈ {clock}，超出 {hi_notes} 字上限 "
                       f"{notes_total - hi_notes} 字：PPT 备注是口播底稿不是逐字稿，"
                       "按页压到每页 90–110 字；直接截断会切掉关键判断"))
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


def load_outline(path: str) -> dict:
    """读 outline.json：容忍 BOM（Windows 记事本/Excel 导出常见），错误信息要能照着修。"""
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        sys.exit(
            f"[E] json-invalid: {path} 第 {e.lineno} 行第 {e.colno} 列解析失败：{e.msg}\n"
            "     常见原因：结尾多逗号、用了中文引号、文件带 BOM（本脚本已容忍 BOM）、"
            "或字段里混了未转义的换行。"
        )
    except OSError as e:
        sys.exit(f"[E] io-error: 读不到 {path}：{e}")


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
    ap.add_argument("--strict", action="store_true",
                    help="机器门禁模式：任何 E、页数越界、文字紧度>100%% 都退出 2")
    args = ap.parse_args()

    outline = load_outline(args.outline)

    errs = validate(outline)
    fatal = [e for e in errs if e[0] == "E"]
    if fatal:
        for lv, code, msg in fatal:
            print(f"[{lv}] {code}: {msg}", file=sys.stderr)
        sys.exit(2)
    for lv, code, msg in [e for e in errs if e[0] != "E"]:
        print(f"[{lv}] {code}: {msg}", file=sys.stderr)

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
    if args.strict:
        gate = [("W", c, m) for _, c, m in soft
                if c in ("page-count", "notes-long", "notes-missing", "notes-short")
                or c.startswith("cap-")]
        gate += [("W", "text-overflow",
                  f"第 {r['index']} 页文字紧度 {r['fill']:.0%}：缩到下限仍溢出，必须删字")
                 for r in rows if isinstance(r.get("fill"), (int, float)) and r["fill"] > 1]
        hard = hard + gate
        for lv, code, msg in gate:
            print(f"✗ [strict] [{code}] {msg}", file=sys.stderr)

    if args.audit_only:
        print(f"✓ 审计完成（{len(rows)} 页，未写 pptx）")
    else:
        print(f"✓ 生成 {out}（{len(rows)} 页）")
    print(f"✓ 审计报告 {report}")
    if soft and not args.strict:
        print(f"! {len(soft)} 条提示（见报告）")
    for lv, code, msg in [h for h in hard if h[0] == "E"]:
        print(f"✗ [{code}] {msg}", file=sys.stderr)

    pdf_ok: bool | None = None
    if args.pdf and not args.audit_only:
        if not (shutil.which("soffice") or shutil.which("libreoffice")):
            print("! 未找到 LibreOffice，跳过 PDF（装好后重跑 --pdf，或直接在 PowerPoint 里导出）",
                  file=sys.stderr)
            pdf_ok = False
        else:
            cur_font = (outline.get("theme_spec") or {}).get("font") or DEFAULT_FONT
            alt = preview_font_for(cur_font)
            if alt:
                print(f"! 本机未装 {cur_font}，PDF 预览改用 {alt}"
                      "（只影响预览件，pptx 字体不变；预览件的行高与字重仅供参考）",
                      file=sys.stderr)
                tmp = os.path.splitext(out)[0] + ".preview.pptx"
                o2 = json.loads(json.dumps(outline))
                o2.setdefault("theme_spec", {})["font"] = alt
                build(o2, tmp, args.template)
                pdf_ok = convert_pdf(tmp)
                if pdf_ok:
                    os.replace(os.path.splitext(tmp)[0] + ".pdf",
                               os.path.splitext(out)[0] + ".pdf")
                    print(f"✓ 导出 PDF：{os.path.splitext(out)[0]}.pdf（预览字体 {alt}）")
                for p in (tmp, os.path.splitext(tmp)[0] + ".pdf"):
                    if os.path.exists(p):
                        os.remove(p)
            else:
                pdf_ok = convert_pdf(out)
                if pdf_ok:
                    print(f"✓ 导出 PDF：{os.path.splitext(out)[0]}.pdf")
            if not pdf_ok:
                msg = ("PDF 未生成：soffice 打不开这份 pptx（常见于模板包结构损坏，"
                       "或本机 LibreOffice 版本过旧）——pptx 本身可能已损坏，请先按 "
                       "§11 把原句报给用户，不要手工改 pptx 绕过 JSON")
                print(f"✗ [pdf-failed] {msg}", file=sys.stderr)
                hard = hard + [("E", "pdf-failed", msg)]
    if hard:
        sys.exit(2 if args.strict or any(h[0] == "E" for h in hard) else 1)
    if pdf_ok is False:
        sys.exit(1)


if __name__ == "__main__":
    main()
