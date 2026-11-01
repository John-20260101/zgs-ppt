# outline.json 契约：17 种页型、字段与字数上限

`scripts/build_pptx.py` 只认这一份 JSON。先写 JSON 再渲染，好处是**版面问题在写的时候就被字数上限挡住**，而不是等生成完再改。

## 0 顶层结构

```json
{
  "group": "Z1",                       // 必填：Z1–Z5，决定页数区间与维度清单
  "group_label": "成长赛道·高教组",     // 可选：页脚显示
  "theme": "academic",                 // academic | tech | warm
  "title": "整份 PPT 的名字",           // 可选，仅用于报告
  "gaps": ["缺企业实习证明"],           // 可选：素材缺口，进审计报告
  "slides": [ ... ]
}
```

每页 = 一个对象。`type` 必填；`dim` 写这一页服务哪个评分维度（审计据此查覆盖）；`notes` 写进 PPT 备注栏（口播要点）。

**阅读型硬约定**：除 cover / section / closing / quote 外，每页应能独立读懂——有论点标题 + 2–6 条可读文字或数字。审计会对内容页查字数：< 90 字报「信息量偏少」，> 520 字报「建议拆页」。

## 1 页型清单

| type | 用途 | 必填字段 | 上限 |
| :--- | :--- | :--- | :--- |
| `cover` | 封面 | title, subtitle, meta[] | title ≤ 20 字 |
| `agenda` | 目录 | title, items[{title,note}] | 2–6 项，note ≤ 40 字 |
| `section` | 章节过渡 / 大论点 | kicker, title, note | title ≤ 24 字 |
| `bullets` | 3–6 条要点（可带小字注释） | title, bullets[{text,note,em}] | ≤ 6 条，单条 ≤ 48 字，note ≤ 40 字 |
| `detail` | 导语 + 2–4 个并列小块 | title, lead, blocks[{title,body}] | lead ≤ 60 字，每块 body ≤ 90 字 |
| `cards` | 2–4 张信息卡 | title, cards[{label,title,body}] | 每卡 body ≤ 70 字 |
| `kpi` | 2–4 个数字卡 | title, items[{value,label,note}] | value ≤ 8 字，label ≤ 14 字 |
| `timeline` | 3–5 个时间节点 | title, items[{when,what,detail}] | what ≤ 36 字，detail ≤ 18 字 |
| `process` | 3–5 步行动路径 | title, steps[{title,detail}] | detail ≤ 70 字 |
| `compare` | 左右对照 | title, left{head,items[]}, right{...} | 每栏 ≤ 4 条，单条 ≤ 28 字 |
| `matrix` | 2×2 四象限 | title, x_label, y_label, quadrants[4] | 每格 text ≤ 70 字 |
| `table` | 对照表 | title, headers[], rows[][] | ≤ 6 列 × 6 行 |
| `evidence` | 证据清单（编号—证据—证明什么—出处） | title, items[{name,prove,source}] | 4–7 条 |
| `twocol` | 左右两块（文/文 或 文/图） | title, left{title,bullets[]}, right{...} | 每栏 ≤ 4 条，单条 ≤ 30 字 |
| `image` | 一张图 + 说明 | title, path, caption | caption ≤ 30 字 |
| `quote` | 一句话大字 | text, by | text ≤ 40 字 |
| `closing` | 签名收尾 | title, subtitle, footer | title ≤ 22 字 |

## 2 逐页型示例

### cover
```json
{"type": "cover",
 "title": "让每一度电都算得清",
 "subtitle": "从一次电池热失控，到我确定要做的那个岗位",
 "meta": ["成长赛道·高教组", "XX大学·新能源材料与器件", "李某某", "目标岗位：电池测试工程师"],
 "dim": "职业目标"}
```

### agenda
```json
{"type": "agenda", "kicker": "AGENDA", "title": "这份材料分四段读",
 "items": [
   {"title": "为什么是电池测试", "note": "一次热失控，让我看见数据漏测的代价"},
   {"title": "岗位到底要什么", "note": "拆 5 条招聘公告，分清会做和能签字"},
   {"title": "我做到了哪一步", "note": "3 组可核验证据"},
   {"title": "还差什么、怎么补", "note": "12 个月行动表 + 备选"}],
 "dim": "职业目标"}
```

### bullets（带小字注释）
```json
{"type": "bullets", "title": "我原本没打算做电池", "kicker": "起点",
 "bullets": [
   {"text": "大一志愿是计算机，被调剂到材料", "note": "2024 级第 3 志愿录取"},
   {"text": "直到那支电芯在测试柜里热失控", "note": "2025.07，编号 A-17", "em": true}],
 "source": "实验室异常事件复盘报告（2025.07）", "dim": "职业目标"}
```
`note` 是小字灰色说明，用来放**口径、时间、出处**——阅读型 PPT 靠它把话说实。

### detail
```json
{"type": "detail", "title": "那次事故教会我三件事", "kicker": "转折",
 "lead": "热失控之后我重做 3 遍复盘，结论不是设备不行，而是方案漏了关键工况。",
 "blocks": [
   {"title": "测什么比怎么测更重要", "body": "原方案只做常温 1C 循环，漏了 45℃ 高温搁置。补测第 6 天复现异常。"},
   {"title": "报告要能被复核", "body": "师兄按我的记录重跑，偏差 2.3%，说明记录方式可用。"}],
 "dim": "职业目标"}
```

### cards
```json
{"type": "cards", "title": "这个岗位一天在做什么", "kicker": "岗位认知",
 "cards": [
   {"label": "01 方案", "title": "写测试大纲", "body": "按 GB/T 31467 拆工况，先定什么算不合格。"},
   {"label": "02 执行", "title": "跑台架与记录", "body": "通道排期、样品编号、异常复测，数据要能追到哪支电芯。"}],
 "dim": "岗位胜任力"}
```

### evidence
```json
{"type": "evidence", "title": "四份材料，撑得起追问",
 "items": [
   {"name": "循环测试台账 37 条", "prove": "会独立排期与判读", "source": "实验室 LIMS 导出"},
   {"name": "企业跟岗鉴定表", "prove": "见过产线级节奏", "source": "带教工程师签字"}],
 "note": "缺 CNAS 实习证明，见 gaps", "dim": "岗位胜任力"}
```

### matrix
```json
{"type": "matrix", "title": "我的能力落在四个格子里",
 "x_label": "岗位使用频率 →", "y_label": "我的掌握程度 ↑",
 "quadrants": [
   {"title": "高频·已掌握", "text": "循环测试执行、数据记录，可以直接上手。"},
   {"title": "高频·待补齐", "text": "异常定位、报告撰写，只写过 2 份周报。"},
   {"title": "低频·已掌握", "text": "设备点检、安全规范。"},
   {"title": "低频·待补齐", "text": "标准体系与客户对接。"}],
 "dim": "优化改进"}
```
顺序固定：左上、右上、左下、右下。

### process
```json
{"type": "process", "title": "12 个月，分四步补",
 "steps": [
   {"title": "3 个月：补标准", "detail": "精读 GB/T 31467，验收：能口述 5 个必测项。"},
   {"title": "6 个月：进企业", "detail": "跟满 1 个项目，验收：独立出 1 份测试周报。"}],
 "dim": "优化改进"}
```

### timeline
```json
{"type": "timeline", "title": "三年，从旁观到上手",
 "items": [
   {"when": "2024.09", "what": "入学，被调剂到材料", "detail": "第一次知道电池是个专业方向"},
   {"when": "2025.07", "what": "实验室热失控", "detail": "主笔复盘"}],
 "dim": "学习实践行动"}
```

### compare
```json
{"type": "compare", "title": "招聘公告 4 条，我够 2 条",
 "left": {"head": "公告硬要求", "items": ["材料相关专业", "熟 GB/T 31467", "台架实操", "可倒班"]},
 "right": {"head": "我的现状（2026.09）", "items": ["专业对口", "只读过 2 份标准", "37 支独立排期", "可倒班"]},
 "note": "来源：某电池企业 2026 校招", "dim": "优化改进"}
```

### twocol
```json
{"type": "twocol", "title": "如果测试岗没去成",
 "left": {"title": "备选一：电芯工艺", "bullets": ["同属电池链，测试经验可迁移", "差距：工艺调优为零"]},
 "right": {"title": "备选二：储能系统测试", "bullets": ["更偏集成", "差距：电力电子基础弱"]},
 "dim": "发展潜力"}
```
要配图时，把 `right` 换成 `"image": "路径.png"`。

## 3 版式常量（改代码时对照）

- 画布 13.333×7.5in；页边距 0.72；标题基线 0.62（高度 0.95）；正文起 1.92；页脚 6.92。
- 标题字号 30→21 自适应，>26 字告警；正文最小 13pt（detail/cards/process 内部可到 9.5–10.5）。
- 页脚左侧写组别，右侧写 `03 / 13` 页码。

## 4 常见改造

| 审计提示 | 怎么改 |
| :--- | :--- |
| `page-thin`（<90 字） | 加 `note` 补口径、加一条证据、或把该页换成 `detail`/`cards` |
| `page-dense`（>520 字） | 拆成两页：论点页 + 证据页 |
| `bullets-tiny` | 删到 4–5 条，或把长条目拆成 `cards` |
| `compare-items` | 每栏删到 4 条，把第 5 条移到 `evidence` |
| `cards-overflow` | 每卡 body 压到 60 字以内 |
| `matrix-count` | 必须正好 4 块 |
| `page-count` | Z1/Z2 12–16、Z3–Z5 12–15 |
