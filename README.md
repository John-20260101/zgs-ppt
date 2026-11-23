# zgs-ppt

全国大学生职业规划大赛（第四届）**生涯发展展示 / 求职综合展示 PPT** 生成器。属于 DSH/Claude 技能，安装即用。

成品是**阅读型**版面：页数按组别 15–20 页，每页不看人讲也能读懂；写大纲前会先对照上一届金奖项目的结构。

四层检查，每层都有退出码，不靠模型自觉：**估算审计**（字数上限/紧度/底部空白/备注区间）→ **渲染复核**（越界/压字/过小字号/字体替换）→ **文件级**（页数/尺寸/备注）→ **同源核对**（outline.json 与大纲 md 逐页对齐）。

## 它解决什么

职规赛材料三件套里，报告有 `zgs-review`、问答有 `zgs-qa`，唯独"现场放的这份 PPT"没人管。这个技能补的就是这一块：
把散材料收成素材卡 → 按组别套七段式页面骨架 → **先出逐页大纲让人确认** → **自检通过后才渲染**可编辑 `.pptx` → 附审计报告（溢出、压字、页数、维度覆盖、素材缺口）。

内容以 `outline.json` 为唯一事实源，大纲 md 是它的视图；改内容只改 JSON，再重跑渲染——不直接改 pptx，否则下次重跑就丢。

## SkillHub

已发布到 SkillHub：**职规赛路演PPT制作**（`@user_3b69e11c/zgs-ppt`）

```bash
skillhub install zgs-ppt --namespace user_3b69e11c
```

或在网页上复制安装提示词：`请根据 https://skillhub.cn/install/skillhub.md，安装 @user_3b69e11c/zgs-ppt。`

页面：https://skillhub.cn/skills/user_3b69e11c/zgs-ppt

> 发布包不含 `.pptx` / `.pdf` / `.docx`（平台只收文本类文件），`LICENSE`、`.gitignore` 也在发布时排除；
> 示例的 outline.json 与大纲/审计文档保留，装好后可直接 `python3 scripts/build_pptx.py examples/case-wang/王某某-outline.json` 复现 18 页成稿。

## 安装

### 方式一：DSH 插件（推荐）

```bash
dsh plugin --profile default add github:John-20260101/zgs-ppt
```

包内声明了 `dsh.bundle`（`cordis.patch.yml` + `lib/index.js`），装好后本包自带的 `SKILL.md` 直接出现在会话技能目录里。改 `SKILL.md` 不需要重装：provider 监听文件，改完热刷新。

### 方式二：装成技能目录（Claude Code / Codex / 任何扫描 skills 的客户端）

```bash
git clone git@github.com:John-20260101/zgs-ppt.git ~/Desktop/Program/deepseek/zgs-ppt
cd ~/Desktop/Program/deepseek/zgs-ppt
./install.sh          # 软链到 ~/.agents/skills、~/.codex/skills、~/.claude/skills
```

- 新开一个会话后生效（技能目录在会话启动时扫描）。
- 不想用软链：`./install.sh --copy`；卸载：`./install.sh --uninstall`。

### 自检

```bash
node scripts/smoke_plugin.mjs                                   # 插件层
python3 -m unittest discover -s tests -v                        # 契约层（含模板包结构）
python3 scripts/build_pptx.py examples/case-wang/王某某-outline.json --audit-only --strict
python3 scripts/check_outline.py \
  examples/case-wang/王某某-outline.json \
  examples/case-wang/王某某-求职综合展示-大纲.md                  # 同源核对
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
├── lib/index.js                 DSH 技能提供者（把本包 SKILL.md 挂进 ctx.skills）
├── cordis.patch.yml             bundle patch（dsh plugin add 读它）
├── package.json                 声明 dsh.bundle；npm/市场据此安装
├── references/
│   ├── page-skeletons.md        三套 15–20 页骨架（成长/就业/职教）+ 页数预算 + 七段式映射
│   ├── awarded-cases.md         上一届（第三届）获奖案例、结构规律、按组别落页做法
│   ├── slide-types.md           outline.json 契约：17 种页型、字段、文案上限（与脚本 CAPS 同一份）
│   └── material-intake.md       素材卡 18 字段、材料抽取映射、缺口替代方案
├── scripts/
│   ├── build_pptx.py            outline.json → pptx（+估算审计，可导出 PDF）
│   ├── verify_pptx.py           渲染复核：转 PDF 查越界/压字/过小字号
│   ├── check_deck.py            文件级检查：页数（按组别区间）/尺寸/备注
│   ├── check_outline.py         JSON 与大纲 md 同源核对（格式不符即退出 2）
│   ├── smoke_plugin.mjs         插件层自检（manifest、frontmatter、list/get）
│   └── extract_content.py       材料抽取（docx/pptx/pdf/md/txt，带页码锚点）
├── assets/themes.json           3 套主题色（academic / tech / warm）
├── assets/school-template.pptx  4:3 模板夹具（测「清空示例页 + 等比缩放」）
├── examples/case-wang/          18 页成稿（pptx/pdf/大纲/审计/outline.json，数据为演示虚构）
├── tests/                       契约回归（stdlib unittest，零新依赖）
├── docs/review/                 生成链路实测审查报告与证据图
├── evals/                       3 个测试提示 + 样例材料
└── install.sh                   安装到各 CLI 技能目录（软链/复制/卸载）
```

## 发布产物

`dist/zgs-ppt-flat.zip` 是按 WorkBuddy 开放平台规范打好的发布包（两级目录、SKILL.md 用 snake_case frontmatter、无二进制、101KB），
更新技能时直接上传这个文件；重新生成一条命令：

```bash
python3 scripts/pack_workbuddy.py            # → dist/zgs-ppt-flat.zip（版本取自 SKILL.md）
python3 scripts/pack_workbuddy.py --check-only   # 只校验现有包是否符合平台规范
```

脚本内写死了平台六条规则（两级目录、SKILL.md 在根、snake_case frontmatter、剔除二进制、
压平后改引用、Python zipfile 保 UTF-8），跑完自检不过就退出 1。

## 依赖

- Python 3.9+，`python-pptx`、`python-docx`、`Pillow`
- 可选：`pdftotext`（读 PDF 材料）、`soffice`（导出 PDF）

```bash
pip3 install python-pptx python-docx Pillow
```

## 已知限制

- PDF 预览的字体替换取决于本机字体环境，只用于查版式，最终外观请在 PowerPoint 确认。
- 内置版式是"结构优先"的通用版面（17 种页型），不做花哨动画/图表；要精细设计请接 `pptx` / `office-pptx` 技能继续美化。
- 阅读型 ≠ 塞满：审计会查每页字数（<90 字报"信息量偏少"、>520 字报"建议拆页"），也会查**版面竖直跨度**（<2.5in 报 `page-empty-bottom`）。
- 字数上限是硬约定：cover 20 / bullets 单条 48 / compare 单条 28 / timeline what 36……超限报 `cap-*`，`--strict` 下直接失败。
- 机器门禁：`--strict` 把 E 级、页数越界、备注出区间、超上限、紧度 >100% 全部按失败处理（退出 2）。
- 不生成图片素材：缺图会画占位框并在审计报告里列出。
