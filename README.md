# zgs-ppt

全国大学生职业规划大赛（第四届）**生涯发展展示 / 求职综合展示 PPT** 生成器。属于 DSH/Claude 技能，安装即用。

成品是**阅读型**版面：页数按组别 12–16 页，每页不看人讲也能读懂；写大纲前会先对照上一届金奖项目的结构。

## 它解决什么

职规赛材料三件套里，报告有 `zgs-review`、问答有 `zgs-qa`，唯独"现场放的这份 PPT"没人管。这个技能补的就是这一块：
把散材料收成素材卡 → 按组别套七段式页面骨架 → **先出逐页大纲让人确认** → 再渲成可编辑 `.pptx` → 附审计报告（溢出、压字、页数、维度覆盖、素材缺口）。

## 安装

```bash
git clone git@github.com:John-20260101/zgs-ppt.git ~/Desktop/Program/deepseek/zgs-ppt
cd ~/Desktop/Program/deepseek/zgs-ppt
./install.sh          # 软链到 ~/.agents/skills、~/.codex/skills、~/.claude/skills
```

- 新开一个会话后生效（技能目录在会话启动时扫描）。
- 不想用软链：`./install.sh --copy`；卸载：`./install.sh --uninstall`。
- 自检：

```bash
python3 scripts/build_pptx.py examples/case-wang/王某某-outline.json --audit-only
```

## 用法

对 agent 说人话就行：

- "帮我用这份生涯发展报告做职规赛 PPT，就业赛道职教组"
- "我是成长赛道高教组，目标岗位芯片测试工程师，先出个 PPT 大纲"
- "这套大纲可以，套我们学校模板生成，要 PDF"

## 结构

```
zgs-ppt/
├── SKILL.md                     主流程（取材→大纲→卡点→渲染→审计→交付）
├── references/
│   ├── page-skeletons.md        三套 13 页骨架（成长/就业/职教）+ 页数预算 + 七段式映射
│   ├── awarded-cases.md         上一届（第三届）获奖案例、结构规律、按组别落页做法
│   ├── slide-types.md           outline.json 契约：10 种页型、字段、文案上限
│   └── material-intake.md       素材卡 18 字段、材料抽取映射、缺口替代方案
├── scripts/
│   ├── build_pptx.py            outline.json → pptx（+估算审计，可导出 PDF）
│   ├── verify_pptx.py           渲染复核：转 PDF 查越界/压字/过小字号
│   └── extract_content.py       材料抽取（docx/pptx/pdf/md/txt，带页码锚点）
├── assets/themes.json           3 套主题色（academic / tech / warm）
└── evals/                       3 个测试提示 + 一份样例材料
```

## 依赖

- Python 3.9+，`python-pptx`、`python-docx`、`Pillow`
- 可选：`pdftotext`（读 PDF 材料）、`soffice`（导出 PDF）

```bash
pip3 install python-pptx python-docx Pillow
```

## 已知限制

- PDF 预览的字体替换取决于本机字体环境，只用于查版式，最终外观请在 PowerPoint 确认。
- 内置版式是"结构优先"的通用版面（17 种页型），不做花哨动画/图表；要精细设计请接 `pptx` / `office-pptx` 技能继续美化。
- 阅读型 ≠ 塞满：审计会查每页字数（<90 字报"信息量偏少"、>520 字报"建议拆页"）。
- 不生成图片素材：缺图会画占位框并在审计报告里列出。
