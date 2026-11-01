# CHANGELOG

## 0.2.0 (2026-11-01)

**内容优化（按用户反馈）**
- 页数改为**所有组别 ≥12 页**：Z1/Z2 12–16，Z3–Z5 12–15；审计同步（`page-count`）。
- 新增 `references/awarded-cases.md`：上一届（第三届）获奖数据、7 个可查案例、7 条结构规律、按组别落页做法，全部带来源链接。
- 改为**阅读型 PPT**：新增 7 种页型（agenda / detail / cards / process / matrix / evidence / twocol），`bullets` 支持每条小字注释、最多 6 条，`timeline` 支持节点说明，`compare` 支持简写列表；新增每页字数审计（<90 字 "page-thin"、>520 字 "page-dense"）。
- 文档同步：`page-skeletons.md` 三套 13 页骨架；`slide-types.md` 扩到 17 种页型并补字数上限。

**修复**
- 4 卡布局正文区过窄导致溢出 → 卡片内边距与字号重算，新增 `cards-overflow` 告警。
- `matrix` 纵轴标签换行挤压象限 → 标签移到网格上方，不再压字。


## 0.1.0 (2026-11-01)

首个版本。

**新增**
- `SKILL.md`：取材 → 定组别页数 → 写大纲（JSON + md）→ 人工卡点 → 渲染 → 估算审计 + 渲染复核 → 交付。
- `references/page-skeletons.md`：Z1/Z2 成长、Z3/Z4 就业、Z5 职教三套页面骨架，含七段式映射与页数/时长换算。
- `references/slide-types.md`：`outline.json` 契约（10 种页型）与文案上限。
- `references/material-intake.md`：素材卡 18 字段、材料抽取映射、缺口替代方案、隐私底线。
- `scripts/build_pptx.py`：13.333×7.5 画布、3 套主题、10 种页型、自动缩字号、压字与页数/维度/缺口审计、模板套版、PDF 导出。
- `scripts/verify_pptx.py`：转 PDF 后用真实文字坐标查越界、压字、过小字号。
- `scripts/extract_content.py`：docx/pptx/pdf/md/txt 抽取，带页码锚点。
- `evals/`：3 个测试提示 + 一份样例材料（Z5 养老服务）。

**已知限制**
- 内置版式是结构优先的通用版面，不做精细视觉设计；要美化请接 `pptx` / `office-pptx`。
- PDF 预览的字体替换取决于本机字体环境，仅用于查版式。
- 估算与渲染复核都不能替代人眼审美判断。

**修复记录（首轮自测发现）**
- KPI 数字卡在标签折行时会被注释压住 → 卡高与标签行数联动计算。
- KPI 页的 `source` 字段被静默丢弃 → 补页脚数据来源。
- 封面 meta 过长出现孤字 → 字数自适应。
- 估算过于乐观（CJK 字宽按 1.03 em）→ 改为 1.10 em，并新增 `verify_pptx.py` 做渲染层复核。
