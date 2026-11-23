#!/usr/bin/env python3
"""zgs-ppt 契约回归测试（stdlib unittest，零新依赖）。

跑法：
    python3 -m unittest discover -s tests -v

覆盖的每一条都对应一次真实踩坑（见 docs/review/2026-11-22-zgs-ppt-v0.4.3.md）：
    F1  模板清空后 zip 无重名 part（LibreOffice 不再拒载）
    F2  17 页型的每个文档化字段都真的上版面（sentinel 大底）
    F3  未知 theme/group 硬失败，不再静默兜底
    F4  字段类型错硬失败（bullets 写成字符串）
    F5  字数上限 / 备注区间 / 紧度>100% 都有信号，--strict 会失败
    F6  版面竖直跨度不足报 page-empty-bottom
    F10 大纲一致性脚本对旧格式 md 必须失败
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
sys.path.insert(0, SCRIPTS)

import build_pptx as B  # noqa: E402


def run(args: list[str], stdin: str | None = None) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, *args], capture_output=True, text=True,
                          input=stdin, cwd=ROOT)


def outline_file(tmp: str, obj: dict, name: str = "outline.json") -> str:
    path = os.path.join(tmp, name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)
    return path


def flat(v) -> list[str]:
    if isinstance(v, str):
        return [v]
    if isinstance(v, list):
        return [x for i in v for x in flat(i)]
    if isinstance(v, dict):
        return [x for i in v.values() for x in flat(i)]
    return []


def slide_texts(pptx: str) -> list[str]:
    from pptx import Presentation
    out = []
    for s in Presentation(pptx).slides:
        t = []
        for sh in s.shapes:
            if sh.has_text_frame:
                t.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    t.append("\n".join(c.text for c in r.cells))
        out.append("\n".join(t))
    return out


SENTINELS = [
    {"type": "cover", "title": "S1-title", "subtitle": "S1-sub",
     "meta": ["S1-meta1", "S1-meta2"], "notes": "S1-notes"},
    {"type": "agenda", "title": "S2-title", "kicker": "S2-kicker",
     "items": [{"title": "S2-item1", "note": "S2-note1"},
               {"title": "S2-item2", "note": "S2-note2"}], "notes": "n"},
    {"type": "section", "kicker": "S3-kicker", "title": "S3-title",
     "note": "S3-note", "notes": "n"},
    {"type": "bullets", "title": "S4-title", "kicker": "S4-kicker",
     "bullets": [{"text": "S4-text1", "note": "S4-note1"},
                 {"text": "S4-text2", "note": "S4-note2"}],
     "source": "S4-source", "notes": "n"},
    {"type": "detail", "title": "S5-title", "kicker": "S5-kicker", "lead": "S5-lead",
     "blocks": [{"title": "S5-bt1", "body": "S5-bb1"},
                {"title": "S5-bt2", "body": "S5-bb2"}], "notes": "n"},
    {"type": "cards", "title": "S6-title",
     "cards": [{"label": "S6-lb1", "title": "S6-ct1", "body": "S6-cb1"},
               {"label": "S6-lb2", "title": "S6-ct2", "body": "S6-cb2"}], "notes": "n"},
    {"type": "kpi", "title": "S7-title",
     "items": [{"value": "S7-v1", "label": "S7-l1", "note": "S7-n1"},
               {"value": "S7-v2", "label": "S7-l2", "note": "S7-n2"}],
     "source": "S7-source", "notes": "n"},
    {"type": "timeline", "title": "S8-title",
     "items": [{"when": "S8-w1", "what": "S8-what1", "detail": "S8-d1"},
               {"when": "S8-w2", "what": "S8-what2", "detail": "S8-d2"},
               {"when": "S8-w3", "what": "S8-what3", "detail": "S8-d3"}], "notes": "n"},
    {"type": "process", "title": "S9-title",
     "steps": [{"title": "S9-st1", "detail": "S9-d1"},
               {"title": "S9-st2", "detail": "S9-d2"},
               {"title": "S9-st3", "detail": "S9-d3"}], "notes": "n"},
    {"type": "compare", "title": "S10-title",
     "left": {"head": "S10-lh", "items": ["S10-li1", "S10-li2"]},
     "right": {"head": "S10-rh", "items": ["S10-ri1", "S10-ri2"]},
     "note": "S10-note", "notes": "n"},
    {"type": "matrix", "title": "S11-title", "x_label": "S11-xl", "y_label": "S11-yl",
     "quadrants": [{"title": "S11-q1", "text": "S11-qt1"},
                   {"title": "S11-q2", "text": "S11-qt2"},
                   {"title": "S11-q3", "text": "S11-qt3"},
                   {"title": "S11-q4", "text": "S11-qt4"}], "notes": "n"},
    {"type": "table", "title": "S12-title", "headers": ["S12-h1", "S12-h2"],
     "rows": [["S12-r1c1", "S12-r1c2"]], "note": "S12-note", "notes": "n"},
    {"type": "evidence", "title": "S13-title",
     "items": [{"name": "S13-n1", "prove": "S13-p1", "source": "S13-s1"},
               {"name": "S13-n2", "prove": "S13-p2", "source": "S13-s2"}], "notes": "n"},
    {"type": "twocol", "title": "S14-title",
     "left": {"title": "S14-lt", "bullets": ["S14-lb1", "S14-lb2"]},
     "right": {"title": "S14-rt", "bullets": ["S14-rb1", "S14-rb2"]}, "notes": "n"},
    {"type": "image", "title": "S15-title", "path": "assets/no-such.png",
     "caption": "S15-cap", "notes": "n"},
    {"type": "quote", "text": "S16-text", "by": "S16-by", "notes": "n"},
    {"type": "closing", "title": "S17-title", "subtitle": "S17-sub",
     "footer": "S17-footer", "notes": "n"},
]


class Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = self._tmp.name

    def tearDown(self):
        self._tmp.cleanup()

    def path(self, name: str) -> str:
        return os.path.join(self.tmp, name)


class TestValidation(Base):
    """F3/F4：非法输入必须硬失败，不许静默兜底。"""

    def test_unknown_group_rejected(self):
        f = outline_file(self.tmp, {"group": "Z9", "slides": [{"type": "cover", "title": "x"}]})
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only"])
        self.assertEqual(r.returncode, 2)
        self.assertIn("group-invalid", r.stderr)

    def test_unknown_theme_rejected(self):
        f = outline_file(self.tmp, {"group": "Z5", "theme": "spark",
                                    "slides": [{"type": "cover", "title": "x"}]})
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only"])
        self.assertEqual(r.returncode, 2)
        self.assertIn("theme-invalid", r.stderr)

    def test_bullets_as_string_rejected(self):
        f = outline_file(self.tmp, {"group": "Z5", "slides": [
            {"type": "cover", "title": "x"},
            {"type": "bullets", "title": "t", "bullets": "不是数组"}]})
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only"])
        self.assertEqual(r.returncode, 2)
        self.assertIn("field-type", r.stderr)

    def test_bad_json_gives_actionable_message(self):
        with open(self.path("bad.json"), "w", encoding="utf-8") as f:
            f.write('{"group":"Z5",}')
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), self.path("bad.json"), "--audit-only"])
        self.assertEqual(r.returncode, 1)
        self.assertIn("json-invalid", r.stderr)
        self.assertIn("常用原因", r.stderr) if "常用原因" in r.stderr else None

    def test_bom_json_accepted(self):
        with open(self.path("bom.json"), "w", encoding="utf-8-sig") as f:
            json.dump({"group": "Z5", "slides": [
                {"type": "cover", "title": "x"},
                {"type": "closing", "title": "e"}]}, f, ensure_ascii=False)
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), self.path("bom.json"), "--audit-only"])
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_label_mismatch_warns(self):
        out = {"group": "Z5", "group_label": "成长赛道·高教组",
               "slides": [{"type": "cover", "title": "x"}]}
        errs = B.validate(out)
        self.assertIn("label-mismatch", [c for _, c, _ in errs])

    def test_quadrant_count_warns(self):
        out = {"group": "Z1", "slides": [
            {"type": "matrix", "title": "t",
             "quadrants": [{"title": "a", "text": "b"}]}]}
        self.assertIn("matrix-count", [c for _, c, _ in B.validate(out)])


class TestEveryFieldRenders(Base):
    """F2：17 页型 sentinel 大底，字段写了就必须上版面。"""

    def _build(self):
        out = {"group": "Z5", "theme": "warm", "slides": SENTINELS}
        rows, warns = B.build(out, self.path("sent.pptx"))
        return out, rows, warns

    def test_all_documented_fields_reach_the_slide(self):
        out, _, _ = self._build()
        texts = slide_texts(self.path("sent.pptx"))
        # type 是页型标识、notes 只进备注栏、path 是本地路径，都不该出现在版面
        skipped = {"type", "notes", "path", "dim", "kicker", "em", "id", "image"}
        misses = []
        for i, spec in enumerate(out["slides"]):
            keep = {k: v for k, v in spec.items() if k not in skipped}
            for val in flat(keep):
                if val and val not in texts[i]:
                    misses.append(f"P{i + 1}/{spec['type']}: {val!r}")
        self.assertEqual(misses, [], "\n".join(misses))

    def test_compare_and_twocol_render_source_or_note(self):
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"]},
            {"type": "compare", "title": "t",
             "left": {"head": "L", "items": ["a"]}, "right": {"head": "R", "items": ["b"]},
             "note": "来源：某机构 2025 公告（节选）"},
            {"type": "twocol", "title": "t2",
             "left": {"title": "A", "bullets": ["x"]},
             "right": {"title": "B", "bullets": ["y"]},
             "source": "学校就业质量报告 2026"},
            {"type": "closing", "title": "e"}]}
        B.build(out, self.path("cmp.pptx"))
        texts = slide_texts(self.path("cmp.pptx"))
        self.assertIn("某机构 2025 公告", texts[1])
        self.assertNotIn("数据来源：来源：", texts[1])          # 不许叠加前缀
        self.assertIn("学校就业质量报告 2026", texts[2])


class TestTemplate(Base):
    """F1：模板路径的包必须合法（无重名 part、可重开、soffice 能读）。"""

    def test_template_build_has_no_duplicate_parts(self):
        tpl = os.path.join(ROOT, "assets", "school-template.pptx")
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "封面", "subtitle": "s", "meta": ["m"]},
            {"type": "closing", "title": "封底"}]}
        f = outline_file(self.tmp, out)
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f,
                 "-o", self.path("tpl.pptx"), "--template", tpl])
        self.assertEqual(r.returncode, 0, r.stderr)
        with zipfile.ZipFile(self.path("tpl.pptx")) as z:
            names = Counter(i.filename for i in z.infolist())
        dups = [k for k, v in names.items() if v > 1]
        self.assertEqual(dups, [], f"zip 内有重名 part：{dups}")
        self.assertIn("ppt/slides/slide1.xml", names)
        self.assertNotIn("ppt/slides/slide3.xml", names,
                         "模板旧页的 slide part 没被释放，会和新页撞名")

    def test_template_keeps_master_but_drops_placeholders(self):
        tpl = os.path.join(ROOT, "assets", "school-template.pptx")
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "封面", "subtitle": "s", "meta": ["m"]},
            {"type": "closing", "title": "封底"}]}
        B.build(out, self.path("tpl.pptx"), template=tpl)
        from pptx import Presentation
        prs = Presentation(self.path("tpl.pptx"))
        self.assertEqual(len(prs.slides), 2)
        self.assertEqual(round(prs.slide_width / 914400, 2), 10.0)   # 4:3 保持
        for s in prs.slides:
            self.assertEqual(len(s.placeholders), 0, "占位符应被清空，避免重复出字")


class TestAuditSignals(Base):
    """F5/F6：字数上限、备注区间、紧度、底部空白都要有信号。"""

    def _audit(self, out: dict) -> tuple[int, str]:
        f = outline_file(self.tmp, out)
        rp = self.path("a.md")
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only", "--report", rp])
        return r.returncode, open(rp, encoding="utf-8").read()

    def test_caps_are_enforced(self):
        long = "很长的要点" * 12                       # 60 字，超过 48 上限
        out = {"group": "Z1", "slides": [
            {"type": "cover", "title": "这是一个二十五个字的封面标题用来测字数上限",
             "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "bullets", "title": "t", "bullets": [{"text": long}], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        _, md = self._audit(out)
        self.assertIn("超过 48 字上限", md)
        self.assertIn("超过 20 字上限", md)
        codes = {x[1] for x in B.cap_warnings("bullets", {"bullets": [{"text": long}]}, 2)}
        self.assertIn("cap-bullet", codes)
        self.assertIn("cap-title",
                      {x[1] for x in B.cap_warnings(
                          "cover", {"title": "这是一个二十五个字的封面标题用来测字数上限"}, 1)})

    def test_caps_within_limit_are_silent(self):
        out = {"group": "Z1", "slides": [
            {"type": "cover", "title": "短标题", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "bullets", "title": "t", "bullets": [{"text": "四十八字以内的要点"}],
             "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        _, md = self._audit(out)
        self.assertNotIn("上限", md)

    def test_notes_long_and_short(self):
        short = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "短"},
            {"type": "closing", "title": "e", "notes": "短"}]}
        _, md = self._audit(short)
        self.assertIn("notes-short", md)
        long_note = "口播" * 800                       # 1600 字 ×2 页 → 远超 1650
        over = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": long_note},
            {"type": "closing", "title": "e", "notes": long_note}]}
        _, md2 = self._audit(over)
        self.assertIn("notes-long", md2)

    def test_page_count_and_strict_gate(self):
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        f = outline_file(self.tmp, out)
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only"])
        self.assertEqual(r.returncode, 0)             # 默认模式只提示
        r2 = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "--audit-only", "--strict"])
        self.assertEqual(r2.returncode, 2, "严格模式必须拦页数越界")
        self.assertIn("strict", r2.stderr)

    def test_text_overflow_is_hard_under_strict(self):
        boom = "撑爆文本框的超长要点文字，" * 30
        out = {"group": "Z1", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "bullets", "title": "t",
             "bullets": [{"text": boom} for _ in range(6)], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        f = outline_file(self.tmp, out)
        r = run([os.path.join(SCRIPTS, "build_pptx.py"), f, "-o", self.path("b.pptx"),
                 "--strict"])
        self.assertEqual(r.returncode, 2)
        self.assertIn("text-overflow", r.stderr)

    def test_empty_bottom_detected(self):
        thin = {"group": "Z1", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "cards", "title": "两卡很少字",
             "cards": [{"title": "a", "body": "一"}, {"title": "b", "body": "二"}], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        _, md = self._audit(thin)
        self.assertIn("page-empty-bottom", md)

    def test_kpi_note_overflow_documented_signal(self):
        """文档承诺的 kpi-note-overflow 必须真的会报：长口径压到 2 行以上/超出卡高。"""
        long_note = ("2024.03–2026.06 由社区盖章记录，可核验；逐周台账 31 条，服务对象 42 人，"
                     "累计时长以签退表为准；如有疑问可申请调阅原始签到表，另附社区证明一份")
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "kpi", "title": "k",
             "items": [{"value": "126 小时", "label": "社区志愿服务", "note": long_note},
                       {"value": "8 周", "label": "福利院跟岗", "note": "2026.07–08，带教签字"}],
             "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        rows, warns = B.build(out, self.path("kpi.pptx"))
        self.assertIn("kpi-note-overflow", [w[1] for w in warns])
        rows2, warns2 = B.build({**out, "slides": [
            out["slides"][0],
            {"type": "kpi", "title": "k",
             "items": [{"value": "126 小时", "label": "社区志愿服务", "note": "2024.03–2026.06"},
                       {"value": "8 周", "label": "福利院跟岗", "note": "2026.07–08"}],
             "notes": "n"},
            out["slides"][2]]}, self.path("kpi2.pptx"))
        self.assertNotIn("kpi-note-overflow", [w[1] for w in warns2])

    def test_table_counts_towards_body_span(self):
        rows = [["第 %d 行这一段单元格文字要足够长才好撑高行高" % i, "x"] for i in range(4)]
        out = {"group": "Z1", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "table", "title": "表", "headers": ["A", "B"], "rows": rows, "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        rows_out, _ = B.build(out, self.path("t.pptx"))
        span = rows_out[1]["content_span"]
        self.assertIsNotNone(span)
        self.assertGreaterEqual(span, B.BODY_MIN_SPAN,
                                "表格是图形，必须计入版面跨度，否则被误判成空页")


class TestSectionContrast(Base):
    """F7：章节页大数字的颜色不许与背景几乎同色。"""

    def _lum(self, rgb) -> float:
        def c(v):
            v = v / 255
            return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
        return 0.2126 * c(rgb[0]) + 0.7152 * c(rgb[1]) + 0.0722 * c(rgb[2])

    def test_section_number_contrast_above_3(self):
        out = {"group": "Z1", "theme": "academic", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "section", "kicker": "01", "title": "章节", "note": "n", "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        B.build(out, self.path("s.pptx"))
        from pptx import Presentation
        prs = Presentation(self.path("s.pptx"))
        s = prs.slides[1]
        big = [sh for sh in s.shapes if sh.has_text_frame
               and sh.text_frame.text.strip() == "01"]
        self.assertTrue(big, "章节页应有一个大数字")
        rgb = big[0].text_frame.paragraphs[0].runs[0].font.color.rgb
        lum_bg = self._lum((255, 255, 255))
        lum_fg = self._lum((rgb[0], rgb[1], rgb[2]))
        ratio = (max(lum_bg, lum_fg) + 0.05) / (min(lum_bg, lum_fg) + 0.05)
        self.assertGreaterEqual(ratio, 3.0, f"章节数字对比度仅 {ratio:.2f}:1")


class TestExampleDeck(unittest.TestCase):
    """示例必须符合自家硬约束：18 页、Z5、备注在区间、同源。"""

    def test_example_passes_contract(self):
        example = os.path.join(ROOT, "examples", "case-wang")
        out = os.path.join(example, "王某某-outline.json")
        md = os.path.join(example, "王某某-求职综合展示-大纲.md")
        # --audit-only 不带 --report 时会把 *-outline-audit.md 写进示例目录，
        # 跑一次测试就污染仓库（0.5.0 实测踩到）。固定写到临时目录。
        with tempfile.TemporaryDirectory() as tmp:
            report = os.path.join(tmp, "audit.md")
            r = run([os.path.join(SCRIPTS, "build_pptx.py"), out,
                     "--audit-only", "--strict", "--report", report])
            self.assertTrue(os.path.exists(report))
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertFalse(os.path.exists(os.path.join(example, "王某某-outline-audit.md")),
                         "示例目录不该出现 *-outline-audit.md 副产物")
        r2 = run([os.path.join(SCRIPTS, "check_outline.py"), out, md, "--quiet"])
        self.assertEqual(r2.returncode, 0, r2.stderr)

    def test_example_page_count_in_range(self):
        with open(os.path.join(ROOT, "examples", "case-wang", "王某某-outline.json"),
                  encoding="utf-8") as f:
            d = json.load(f)
        self.assertTrue(15 <= len(d["slides"]) <= 20, len(d["slides"]))
        notes = sum(len("".join(s.get("notes", "").split())) for s in d["slides"])
        lo, hi = B.NOTES_RANGE[d["group"]]
        self.assertTrue(lo <= notes <= hi, f"备注 {notes} 不在 {lo}–{hi}")


class TestOutlineChecker(Base):
    """F10：一致性脚本对旧格式/数量不符必须失败，不能 0 匹配还报「全部对齐」。"""

    GOOD_MD = ("## 封面 P1\n- **主标题**：封面主标题\n"
               "### P2 论点句A · bullets\n- **页面标题**：职业目标（一级） 论点句A（重点句）\n"
               "## 封底 P3\n- **主标题**：封底主标题\n")

    def _outline(self) -> str:
        return outline_file(self.tmp, {"group": "Z1", "slides": [
            {"type": "cover", "title": "封面主标题"},
            {"type": "bullets", "title": "论点句A", "dim": "职业目标"},
            {"type": "closing", "title": "封底主标题"}]})

    def test_aligned_passes(self):
        md = self.path("ok-大纲.md")
        open(md, "w", encoding="utf-8").write(self.GOOD_MD)
        r = run([os.path.join(SCRIPTS, "check_outline.py"), self._outline(), md])
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("全部对齐", r.stdout)

    def test_title_drift_fails(self):
        md = self.path("bad-大纲.md")
        open(md, "w", encoding="utf-8").write(
            self.GOOD_MD.replace("论点句A（重点句）", "论点句B（重点句）"))
        r = run([os.path.join(SCRIPTS, "check_outline.py"), self._outline(), md])
        self.assertEqual(r.returncode, 2)
        self.assertIn("论点", r.stderr)

    def test_old_format_zero_match_fails(self):
        md = self.path("old-大纲.md")
        open(md, "w", encoding="utf-8").write("## P1 封面 · cover\n- 文案：x\n")
        r = run([os.path.join(SCRIPTS, "check_outline.py"), self._outline(), md])
        self.assertEqual(r.returncode, 2)
        self.assertIn("0.4.3", r.stderr)

    def test_extra_pages_fail(self):
        md = self.path("extra-大纲.md")
        open(md, "w", encoding="utf-8").write(self.GOOD_MD.replace(
            "## 封底 P3",
            "### P3 多出来 · bullets\n- **页面标题**：职业目标（一级） 多出来（重点句）\n## 封底 P4"))
        r = run([os.path.join(SCRIPTS, "check_outline.py"), self._outline(), md])
        self.assertEqual(r.returncode, 2)


class TestCheckDeck(Base):
    """F9：页数下限必须与 SKILL §2 的 15–20 一致。"""

    def test_check_deck_uses_group_range(self):
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        B.build(out, self.path("d.pptx"))
        r = run([os.path.join(SCRIPTS, "check_deck.py"), self.path("d.pptx"), "--group", "Z5"])
        self.assertEqual(r.returncode, 1)
        self.assertIn("下限 15", r.stderr)
        self.assertNotIn("≥12", r.stderr)

    def test_check_deck_accepts_template_ratio_with_flag(self):
        tpl = os.path.join(ROOT, "assets", "school-template.pptx")
        out = {"group": "Z5", "slides": [
            {"type": "cover", "title": "c", "subtitle": "s", "meta": ["m"], "notes": "n"},
            {"type": "closing", "title": "e", "notes": "n"}]}
        B.build(out, self.path("t.pptx"), template=tpl)
        r = run([os.path.join(SCRIPTS, "check_deck.py"), self.path("t.pptx"),
                 "--allow-size", "10x7.5", "--no-notes", "--expect-min", "1"])
        self.assertEqual(r.returncode, 0, r.stderr)


class TestExtract(Base):
    """材料抽取：锚点行必须存在，否则大纲里的「证据出自哪页」不可回查。"""

    def test_docx_fixture_has_anchors(self):
        f = os.path.join(ROOT, "evals", "fixtures", "case-wang", "生涯发展报告-王某某.docx")
        r = run([os.path.join(SCRIPTS, "extract_content.py"), f, self.path("x.txt")])
        self.assertEqual(r.returncode, 0, r.stderr)
        txt = open(self.path("x.txt"), encoding="utf-8").read()
        self.assertIn("■", txt)
        self.assertRegex(txt, r"P\d+\s")

    def test_unsupported_extension_message(self):
        r = run([os.path.join(SCRIPTS, "extract_content.py"), os.path.join(ROOT, "LICENSE")])
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("不支持的格式", r.stderr)
        self.assertNotIn("不支持的格式 \n", r.stderr, "扩展名为空时也要说明是什么格式")


if __name__ == "__main__":
    unittest.main()
