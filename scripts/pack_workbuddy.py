#!/usr/bin/env python3
"""zgs-ppt / pack_workbuddy.py —— 按 WorkBuddy 开放平台规范打包技能 zip。

平台规则（2026-11 实测，写死在这里免得再翻笔记）：
  1. zip ≤ 3MB；只支持两级目录（`a/b/c/file` 会报「目录层级超限」）；
  2. SKILL.md 必须在 zip 根目录，不要套 `zgs-ppt/` 外层目录；
  3. frontmatter 要 snake_case：name / display_name / display_name_en /
     description_zh / description_en / examples_zh / examples_en / author；
  4. 剔除 .git / .github / .gitignore / LICENSE 与全部 .pptx/.pdf/.docx/.png；
  5. 压平后要同步改 README、install.sh、evals.json 里的 `examples/case-wang/` 路径；
  6. 必须用 Python zipfile 打包——`zip` 命令不置 UTF-8 标志，中文文件名会乱码。

用法：
    python3 scripts/pack_workbuddy.py                    # → dist/zgs-ppt-flat.zip
    python3 scripts/pack_workbuddy.py --out /tmp/x.zip
    python3 scripts/pack_workbuddy.py --check-only       # 只校验 dist 里现有包

版本号取自 SKILL.md 的 frontmatter，所以「改 SKILL.md → 跑本脚本 → 上传」是完整流程。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(ROOT, "dist", "zgs-ppt-flat.zip")

BIN_EXT = {".pptx", ".pdf", ".docx", ".png", ".jpg", ".zip"}
SKIP_TOP = {".git", ".github", ".gitignore", ".DS_Store", "LICENSE",
            "tests", "docs", "dist"}
# 仓库三级路径 → 平台两级路径
FLATTEN = {"examples/case-wang": "examples", "evals/fixtures/case-wang": "evals"}
PREFIX = {"evals": "case-wang-"}
REQUIRED_FM = ["name", "display_name", "display_name_en", "description_zh",
               "description_en", "examples_zh", "examples_en", "author"]


def fm_version(fm: str) -> str | None:
    """从 frontmatter 文本里取 version（逐行解析，不用正则尤其实在更稳）。"""
    for line in fm.splitlines():
        if line.startswith("version:"):
            return line.split(":", 1)[1].strip().strip("'\"")
    return None


def read_version() -> str:
    with open(os.path.join(ROOT, "SKILL.md"), encoding="utf-8") as f:
        ver = fm_version(f.read().split("---")[1])
    if not ver:
        sys.exit("[E] SKILL.md frontmatter 里没有 version")
    return ver


def collect() -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    for root, dirs, files in os.walk(ROOT):
        if root == ROOT:
            dirs[:] = [d for d in dirs if d not in SKIP_TOP]
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        rel_root = os.path.relpath(root, ROOT).replace(os.sep, "/")
        if rel_root == ".":
            rel_root = ""
        for f in files:
            if f == ".DS_Store" or os.path.splitext(f)[1].lower() in BIN_EXT:
                continue
            if f.endswith("-outline-audit.md"):     # --audit-only 的副产物
                continue
            src_rel = f"{rel_root}/{f}" if rel_root else f
            if src_rel.split("/")[0] in SKIP_TOP:
                continue
            dest_rel = src_rel
            for long_pre, short_pre in FLATTEN.items():
                if src_rel.startswith(long_pre + "/"):
                    rest = src_rel[len(long_pre) + 1:]
                    dest_rel = f"{short_pre}/{PREFIX.get(short_pre, '')}{rest}"
                    break
            items.append((src_rel, dest_rel))
    return sorted(items)


def rewrite(path: str) -> None:
    """包内路径引用跟着压平改；包里没有的东西不留死链。"""
    if not path.endswith((".md", ".sh", ".json")):
        return
    text = open(path, encoding="utf-8").read()
    text = text.replace("examples/case-wang/", "examples/")
    text = text.replace("evals/fixtures/case-wang/", "evals/case-wang-")
    text = text.replace("`assets/school-template.pptx`", "仓库内的 `assets/school-template.pptx`")
    text = text.replace("| `tests/` | 契约回归（17 页型字段、上限、模板包结构） | 改脚本前后 |\n", "")
    text = text.replace("`tests/test_contract.py` 的字段清单", "仓库 `tests/test_contract.py` 的字段清单")
    if path.endswith("README.md"):
        text = re.sub(r"├── assets/school-template\.pptx.*\n", "", text)
        text = re.sub(r"├── tests/.*\n", "", text)
        text = re.sub(r"├── docs/review/.*\n", "", text)
        text = re.sub(r"```bash\npython3 -m unittest discover -s tests -v.*?```",
                      "", text, flags=re.S)
        text = text.replace("python3 -m unittest discover -s tests -v"
                            "                        # 契约层（含模板包结构）\n", "")
        # 示例目录压平后，README 的结构行要跟着改（用宽松匹配，别依赖具体空格数）
        text = re.sub(
            r"├── examples/[^\n]*\n",
            "├── examples/                   18 页示例：outline.json + 大纲 + 审计"
            "（pptx/pdf 由脚本生成）\n", text)
        # 仓库 README 里的「发布产物」段在包内是自指（包里没有 dist/），去掉
        text = re.sub(r"## 发布产物\n\n[\s\S]*?\n## 依赖", "## 依赖", text)
        if "本发布包不含" not in text:
            text = text.replace(
                "## 已知限制\n",
                "## 已知限制\n\n- 本发布包不含 `.pptx/.pdf/.docx` 与模板夹具（平台规范排除二进制）；"
                "需要模板与契约测试请在 GitHub 仓库里跑 `tests/`。\n")
    elif path.endswith("evals.json"):
        text = re.sub(r'"files":\s*\[\s*"assets/school-template\.pptx"\s*\]', '"files": []', text)
        text = text.replace("模板在 assets/school-template.pptx。", "模板由我提供（学校统一模板，稍后发你）。")
        text = text.replace("--template assets/school-template.pptx", "--template <学校模板>.pptx")
    open(path, "w", encoding="utf-8").write(text)


def add_platform_frontmatter(path: str) -> None:
    """补平台要的 snake_case 字段（内容沿用已有的中文/英文描述）。"""
    text = open(path, encoding="utf-8").read()
    fm, body = text.split("---", 2)[1], text.split("---", 2)[2]
    have = {m.group(1) for m in re.finditer(r"^([A-Za-z0-9_]+):", fm, re.M)}
    lines = []
    if "display_name" not in have:
        dn = re.search(r"^displayName:\s*(.+)$", fm, re.M)
        if dn:
            lines.append(f"display_name: {dn.group(1).strip()}")
    if "display_name_en" not in have:
        lines.append("display_name_en: Career-Planning Competition Deck Builder")
    if "description_zh" not in have:
        lines.append("description_zh: 把选手的生涯发展报告/求职简历与素材，做成 15–20 页阅读型职规赛展示 PPT："
                     "先出逐页大纲待确认，自检通过再渲染可编辑 .pptx，并附文字溢出与信息量审计。")
    if "description_en" not in have:
        lines.append("description_en: Turns a contestant's report, resume and materials into a 15-20 page "
                     "reading-type PPTX for China's National College Student Career Planning Competition "
                     "- outline-first, audited before render.")
    if "examples_zh" not in have:
        lines += ["examples_zh:",
                  '- "帮我用这份生涯发展报告做职规赛 PPT，就业赛道职教组"',
                  '- "我是成长赛道高教组，目标岗位芯片测试工程师，先出个 PPT 大纲"',
                  '- "这套大纲可以，套我们学校模板生成，要 PDF"']
    if "examples_en" not in have:
        lines += ["examples_en:",
                  '- "Build a career-planning competition deck from this report (Employment track, vocational group)"',
                  '- "Growth track, higher-ed group, target role chip test engineer - give me a page-by-page outline first"',
                  '- "Outline approved - apply our school template and export a PDF"']
    if "author" not in have:
        lines.append('author: "John"')
    new_fm = fm.rstrip("\n") + ("\n" + "\n".join(lines) + "\n" if lines else "\n")
    open(path, "w", encoding="utf-8").write("---" + new_fm + "---" + body)


def validate(zip_path: str, verbose: bool = True) -> list[str]:
    """校验平台硬规则；返回问题清单（空 = 通过）。"""
    size = os.path.getsize(zip_path)
    problems = []
    if size > 3 * 1024 * 1024:
        problems.append(f"体积 {size/1024/1024:.1f}MB 超过 3MB")
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        depth = max(n.count("/") for n in names)
        if depth > 1:
            problems.append(f"最大目录层级 {depth}（平台限 2 级，即文件最多在二级目录）")
        if "SKILL.md" not in names:
            problems.append("SKILL.md 不在 zip 根目录")
        bins = [n for n in names if os.path.splitext(n)[1].lower() in BIN_EXT]
        if bins:
            problems.append(f"混入二进制：{bins[:3]}")
        fm = z.read("SKILL.md").decode("utf-8").split("---")[1]
        missing = [k for k in REQUIRED_FM if not re.search(rf"^{k}:", fm, re.M)]
        if missing:
            problems.append(f"frontmatter 缺平台字段：{missing}")
        if not all(i.flag_bits & 0x800 for i in z.infolist()
                   if any(ord(c) > 127 for c in i.filename)):
            problems.append("中文文件名未置 UTF-8 标志（需用 Python zipfile 打包）")
        if z.testzip() is not None:
            problems.append("zip CRC 校验失败")
        ver = fm_version(fm)
        if ver is None:
            problems.append("SKILL.md frontmatter 里没有可解析的 version")
        if verbose:
            print(f"  {len(names)} 文件 · {size/1024:.0f} KB · 层级 {depth} · 版本 {ver}")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser(description="按 WorkBuddy 平台规范打包")
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--check-only", action="store_true", help="只校验 --out 指向的现有 zip")
    args = ap.parse_args()

    if args.check_only:
        problems = validate(args.out)
        for p in problems:
            print(f"✗ {p}", file=sys.stderr)
        if problems:
            return 1
        print("✓ 通过平台规范校验")
        return 0

    version = read_version()
    items = collect()
    stage = tempfile.mkdtemp(prefix="zgs-wb-")
    try:
        for src_rel, dest_rel in items:
            dest = os.path.join(stage, dest_rel)
            os.makedirs(os.path.dirname(dest) or stage, exist_ok=True)
            shutil.copy2(os.path.join(ROOT, src_rel), dest)
        for src_rel, dest_rel in items:
            rewrite(os.path.join(stage, dest_rel))
        add_platform_frontmatter(os.path.join(stage, "SKILL.md"))

        deep = [d for _, d in items if d.count("/") > 1]
        if deep:
            sys.exit(f"[E] 目录层级超限（平台只支持两级）：{deep[:5]}")

        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        if os.path.exists(args.out):
            os.remove(args.out)
        # Python zipfile：自动为中文名置 0x800（zip 命令不会）
        with zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED) as z:
            for _, dest_rel in items:
                info = zipfile.ZipInfo(dest_rel)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o644 << 16
                z.writestr(info, open(os.path.join(stage, dest_rel), "rb").read())
    finally:
        shutil.rmtree(stage, ignore_errors=True)

    problems = validate(args.out)
    if problems:
        for p in problems:
            print(f"✗ {p}", file=sys.stderr)
        return 1
    print(f"✓ 已生成 {args.out}（版本 {version}）—— 上传到 WorkBuddy 开放平台")
    return 0


if __name__ == "__main__":
    sys.exit(main())
