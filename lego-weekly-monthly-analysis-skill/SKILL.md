---
name: lego-weekly-monthly-analysis
description: "苏中周月维度分析看板生成与维护技能。触发词：更新周月分析、周月看板、周分析、月分析、周月维度分析、机会点、模块结论、成交率、转换率、客流、长尾款、区域对比布局。当用户需要生成、修改或重排乐高苏中区域7家门店的周/月维度分析看板HTML时使用此技能。涵盖12脚本流水线（HTML生成+JSON提取+JSON注入+JS校验）的完整流程、每个模块的结论+机会点生成逻辑，以及成交率模块（月_conversion+周_weekly_conversion）与区域对比布局重排（长尾款+成交率半宽并排第三排）的可复用工作流。详细规则参见 references/weekly_monthly_rules.md（第十八章）。"
agent_created: true
---

# 苏中周月维度分析看板

## 概述

每周自动生成苏中区域7家门店的周维度+月维度交互分析HTML，包含同比、环比、产品结构、导购排名、样品销售等维度。

## 核心工作流

### 手动触发周月分析

```
1. 运行 auto_run_weekly.py：
   <你的python> <项目目录>/auto_run_weekly.py
2. 部署 outputs_analysis/ 目录到云端（cloudstudio-deploy）
3. present_files 展示云端链接
```

### 12脚本流水线

```
Step 1: HTML模板生成（7个脚本，必须先运行）
  gen_weekly_html.py → gen_weekly_p2.py → gen_weekly_p3.py → gen_weekly_p4.py
  gen_monthly_html.py → gen_monthly_p2.py → gen_monthly_p3.py

Step 2: JSON数据提取（4个脚本）
  init_weekly_json.py → extract_weekly_guide.py → extract_weekly_enhanced.py → extract_monthly.py

Step 3: JSON注入（1个脚本）
  inject_json.py → 替换HTML中 const DATA = {...}
```

**关键**：HTML生成必须在JSON提取之前，因为 inject_json.py 依赖HTML中的 `const DATA = {...}` 占位符。

### 修改看板逻辑

- 页面结构/KPI卡片/图表 → 修改对应 `gen_*.py` 脚本
- 数据计算逻辑 → 修改对应 `extract_*.py` 或 `init_weekly_json.py`
- 修改后运行 `auto_run_weekly.py` 全流程重新生成

## 关键规则速查

### 数据清洗（同日看板）

去重key、排除赠品/样品/零金额、排除4类人员、门店映射 — 与日看板完全一致。详见 `references/weekly_monthly_rules.md` 第四节。

### 新店 null 保护（淮安万象城）

- 所有同比字段设为 None
- JS中所有 `yoy_pct.toFixed()` 必须加null检查：
  ```javascript
  (day.yoy_pct != null ? (day.yoy_pct >= 0 ? '+' : '') + day.yoy_pct.toFixed(1) + '%' : 'N/A')
  ```
- 区域同比用6店口径（排除淮安万象城）
- 区域环比用7店口径（淮安万象城有上月数据）

### 导购指标特殊处理

- 指标表中只有李婷/黄小莉，需手动创建李婷1/黄小莉1条目（复制指标）
- 删除李婷/黄小莉本体（他们的销售数据已排除）
- 姓名映射：扬万支援销售→李婷1，泰州万象城-支援2→黄小莉1，长期兼职-前缀去掉

### JSON依赖链

```
init_weekly_json.py
  ← 读取 dashboard_latest.json（区域数据，来自日看板输出）
  → 输出 weekly_analysis.json（基础结构）

extract_weekly_guide.py
  ← 读取 weekly_analysis.json → 更新导购+门店周同比

extract_weekly_enhanced.py
  ← 读取 weekly_analysis.json → 更新环比+系列+产品+样品+日指标

extract_monthly.py
  → 输出 monthly_analysis.json（月度全部数据）

inject_json.py
  ← 读取 weekly/monthly_analysis.json → 替换HTML中的const DATA
```

### JS语法校验（部署前必须通过）

```javascript
// Node.js: new Function() 检查
const stubs = 'const Chart=function(){};const document={...};';
new Function(stubs + jsContent);  // 有语法错误会抛异常
```
- 常见错误：`const` 变量重复声明
- 校验失败 → 不部署，页面会空白

### 指标同期Excel读取

- 行范围：rows 13-20（门店+区域去年同期数据）
- 指标同期表为**当月单列**结构：7月表 col2-32=7/1-7/31，8月表 col2-32=8/1-8/31，**没有跨月列**
- **同一月内**周同比：`WK_COL_START = WK_START.day + 1`，值 = `sum(col(WK_COL_START) : col(WK_COL_END))`（去年同期该周累计）
- 列索引安全：`min(j, target_df.shape[1])`
- ⚠️ **跨月周（WK起点月份 ≠ 报告月）绝不能用 `WK_COL_START = 2` 取去年同周**（会丢7月底列）→ 必须按下方「跨月周取数铁律」实算

### ⚠️ 跨月周取数铁律（最高优先级）

**触发**：自然周 `WK_START.month ≠ REPORT_MONTH`（如 `WK31=2026-07-27` 起，报告月8月 → 周跨 7/27–8/2）。此时周销售额/周目标/周完成率/周同比都必须按完整周窗口实算，**不能按"当月"切**。

**根因**：指标同期表按月单列（见上），任何"按 month 取列"的裸逻辑在跨月周会丢/错 7月底数据，造成两类 bug：
1. **周同比假高（如 +140%）**：分子 = 整周7天（`wk_amt`），分母（去年同周）却只取到当月1号起的列（8/1–8/2 两天）→ 分母严重不足
2. **7月底日目标错位**：`extract_weekly_enhanced.py` daily 段原 `t = s_targets[day - 1]` 按"8月表"索引，`7/27`（day=27）误取 `s_targets[26]` = **8/27 的目标**

**修法（标准动作）**：
1. `extract_weekly_enhanced.py` 2.5段：同时读 `7月指标同期.xlsx` 与 `8月指标同期.xlsx`，合并成按 `(月,日)` 索引的 dict：`store_daily_target[(m,d)]`、`store_daily_yoy[(m,d)]`、`region_daily_target[(m,d)]`、`region_daily_yoy[(m,d)]`（7月底查7月表、8月初查8月表）
2. `extract_weekly_enhanced.py` daily 构建段（门店 ~538行 / 区域 ~567行）：`t = s_targets[day - 1]` → `t = store_daily_target[(ym, day)]`；`y = s_yoys[day - 1]` → `y = store_daily_yoy[(ym, day)]`
3. `extract_weekly_guide.py` 周同比段（~246行）：删除 `if WK_START.month == REPORT_MONTH` 的列索引 hack。改为遍历实际周窗口每天，按 `(年-1,月,日)` 从对应月份指标同期表取去年同期行求和（跨月自动 7月+8月两段相加）

**验证（重跑前必做）**：独立脚本读两表 + 当前 `weekly_analysis.json`，算：
- 区域周同比 = sum(7天 `wk_amt`) / sum(去年同周7天 yoy) − 1
- 区域周完成率 = sum(7天 `wk_amt`) / sum(7天 `daily.target`)
预期：修复后周同比应接近月同比量级（如 −0.4%），而非 +140%。

**指标表路径**：`/Users/a123/Desktop/。/7月指标同期.xlsx`、`/Users/a123/Desktop/。/8月/8月指标同期.xlsx`

### 长尾款SKU读取

- `extract_weekly_guide.py` 中有已知bug：使用 `iloc[:,0]` 读取SKU（那是序号列）
- 正确应使用 `lt_sku_df['乐高编号']`
- init_weekly_json 和 extract_weekly_enhanced 使用正确列名

## KPI卡片说明

### 周门店详情页（8个卡片）

周累计 | 周环比 | 周同比 | 周连带率 | 周客单价 | WOS周累计 | 长尾款周累计 | 即时零售周累计

> 笔数已改为连带率（2026-07-19变更）

### 月门店详情页（8个卡片）

月累计 | 月环比 | 月同比 | 月连带率 | 月客单价 | WOS月累计 | 长尾款月累计 | 即时零售月累计

## 区域对比页图表说明

### 门店累计销售&完成率对比图（周/月通用）

- **图表类型**：双轴（柱=累计金额 + 线=完成率%）
- **排序**：按完成率从高到低
- **x轴标签**：门店名 + 完成率%（如"扬州万象汇 52.1%"）
- **左轴(y)**：累计金额（¥），柱状图，门店色块
- **右轴(y1)**：完成率（%），红色折线，`grid:{drawOnChartArea:false}`
- **tooltip**：柱→销售+指标+完成率；线→完成率
- **实现文件**：`gen_weekly_p4.py`（周）+ `gen_monthly_p3.py`（月）
- **修改注意**：gen_weekly_p4.py用unicode转义存储中文，Edit工具无法匹配，需用Python脚本替换

## 配色规则

与日看板一致的模块色。门店Tab高亮色：

| 门店 | 颜色 |
|------|------|
| 扬州万象汇 | #7c3aed |
| 扬州京华城 | #0ea5e9 |
| 扬州江都金鹰 | #059669 |
| 泰州万象城 | #dc2626 |
| 宿迁宝龙 | #ea580c |
| 淮安新亚 | #0891b2 |
| 淮安万象城 | #9333ea |
| 区域对比 | #1e293b |

## 模块结论与机会点

每个图表/表格模块下方自动生成两个区域：

### CSS样式

结论区（紫色左边框）：
```css
.conclusion{border-left:3px solid var(--purple);background:#f8fafc}
.conc-title{font-weight:700;color:#1e293b}       /* 📋 结论 */
.conc-body{color:#475569}
```

机会点区（琥珀色背景）：
```css
.conc-opportunity{background:#fef3c7;border-top:1px solid #f59e0b;color:#92400e}
.conc-label{font-weight:700;color:#b45309}       /* 💡 需关注的机会点 */
```

### `conc()` 函数签名

```javascript
function conc(conclusion, opportunity) {
  if (!opportunity) return 单区域（只有结论，向后兼容）
  return 双区域（结论 + 机会点）
}
```

- `conclusion`：必填，模块结论文本
- `opportunity`：可选，机会点文本；不传则只显示结论区

### `conc*()` 函数族

9个函数，每个对应一个图表/数据模块（定义在 `gen_weekly_p2.py` / `gen_monthly_html.py`）：

| 函数 | 模块 | 机会点逻辑 |
|------|------|-----------|
| `concKPI` | KPI卡片汇总 | 门店所有KPI（完成率/同比/环比/客单价/连带率/WOS占比/长尾款占比）逐项**vs区域**，低于的全部列出具体差距 |
| `concSeries` | 产品系列环形图（TOP10） | TOP10逐项标注占比 / TOP3集中度 / CITY·TECHNIC·NINJAGO是否在TOP3 / MINECRAFT是否在TOP5 / BOTANICALS植物系列占比（指标4.1%，低于指标标注机会点） |
| `concRadar` | 能力雷达图 | 5维度逐项**vs区域均值**，标注具体差距（¥/比率/占比%） |
| `concLongTail` | 长尾款分析 | 长尾款占比 vs 区域均值 + 金额 vs 区域门店均值（<70%标注） |
| `concWow` / `concMom` | 周环比 / 月环比 | <-8%下滑 / <-3%小幅下滑 / <3%持平 / ≥3%增长 |
| `concTop` | TOP10产品排名 | TOP1占比>25%提醒；当月新品是否进入TOP10；区域TOP10产品门店覆盖检查 |
| `concGuideSales` | 导购销售排名 | **门店人均+区域人均**双对比，低于70%列出导购+金额 |
| `concGuideJdKdj` | 导购连带+客单价 | **vs区域均值**，显示¥差值和比率差值 |
| `concGuideLt` | 导购长尾款分析 | **门店人均**对比 + 零长尾款销售导购标注 |

### 机会点生成规则（数据驱动，均值对比优先）

**核心原则**：优先与区域/门店均值对比，显示具体差距数值；绝对阈值仅作为辅助参考。

| 维度 | 对比基准 | 阈值 | 建议文本 |
|------|---------|------|---------|
| KPI汇总 | 区域整体 | <区域值 | "低于区域整体的KPI：完成率XX%(区域XX%,差Xpct)；客单价¥XX(区域¥XX,差¥XX)等，需逐项分析差距原因" |
| 完成率 | 区域完成率 | <区域-5pct | "完成率低于区域均值X个百分点，需检视目标拆解和执行节奏" |
| 完成率 | 绝对值 | <40% | "完成率偏低（X%），建议加强周末高峰冲量" |
| 完成率 | 绝对值 | 40-54% | "完成率处于中游（X%），关注后续销售节奏" |
| 完成率 | 绝对值 | ≥55% | "完成率表现良好（X%），保持势头" |
| 雷达维度 | 区域均值 | 低于区域 | 逐维度标注差距：客单价差¥XX / 连带率差X.XX / 占比差X.X% |
| 环比 | 绝对值 | <-8% | "环比下滑X%，幅度较大，需排查原因" |
| 环比 | 绝对值 | -8%~-3% | "环比小幅下滑X%，持续关注" |
| 环比 | 绝对值 | -3%~3% | "环比基本持平，处于稳定区间" |
| 环比 | 绝对值 | ≥3% | "环比增长X%，势头良好" |
| TOP1依赖 | 绝对值 | >25% | "TOP1占比X%，过度依赖单一产品" |
| 长尾款 | 区域门店均值 | <均值×0.7 | "长尾款金额低于区域门店均值（¥XX），需重点跟进长尾款推荐" |
| 导购销售 | 门店人均70% | <门店均值×0.7 | "XX(¥XX)低于门店均值¥XX，需重点跟进辅导" |
| 导购销售 | 区域人均70% | <区域均值×0.7 | "XX低于区域导购均值¥XX，建议安排导师带教" |
| 导购连带 | 区域均值 | <区域-0.1 | "连带率低于区域均值X.XX，建议强化搭配推荐话术培训" |
| 导购客单价 | 区域均值85% | <区域×0.85 | "客单价低于区域¥XX，建议引导客户关注高价套装" |
| 导购长尾 | 门店人均 | 零销售/占比<3% | "XX本周零长尾款销售，需重点跟进79款SKU推荐" |

### 区域对比页 inline conc() 机会点逻辑

区域对比页的 inline `conc()` 同样采用均值对比：

| 模块 | 对比方式 | 机会点输出 |
|------|---------|-----------|
| 门店完成率对比 | 计算区域均值，列出低于均值的门店 | "低于区域均值（X%）的门店：XX(差Xpct)、XX(差Xpct)…" |
| 客单价&连带率对比 | vs区域均值 | "XX客单价低于区域¥XX，XX连带率低于区域X.XX" |
| 周环比对比 | vs区域环比均值 | "XX环比低于区域均值X%，需复盘客流转化" |
| 导购排名 | 计算区域导购人均，列出低于70%的 | "低于区域均值¥XX的导购：XX(¥XX)、XX(¥XX)…，需重点跟进" |

### 实现位置

- **`conc*()` 函数定义**：`gen_weekly_p2.py`（周）+ `gen_monthly_html.py`（月）
- **inline `conc()` 调用**：`gen_weekly_p3.py` / `gen_weekly_p4.py`（周门店/区域）+ `gen_monthly_p2.py` / `gen_monthly_p3.py`（月门店/区域）
- **CSS样式**：`gen_weekly_html.py`（周）+ `gen_monthly_html.py`（月）

### 新增机会点注意事项

1. 修改 `conc*()` → 同步修改周(`gen_weekly_p2.py`)和月(`gen_monthly_html.py`)两份文件
2. 修改 inline `conc()` → 同步修改门店页和区域对比页
3. 机会点必须是中文、可落地的业务建议，避免空洞模板话术
4. 数据阈值以当时实际业务情况为准，可按需调整

## 成交率模块 + 区域对比布局重排（可复用工作流）

> 本工作流封装自实际改造：在周月看板加入「成交率（笔数/客流）」模块，并把区域对比页的「长尾款 + 成交率」缩成半宽并排、挪到第三排，门店页也加成交率卡片。月/周两侧口径一致、排版一致。

### 触发场景

- 用户：「在周月看板加成交率（=成交笔数/客流量）」
- 用户：「区域对比页把成交率和长尾款缩小放一排、挪到第三排」
- 用户引用一份周/月客流 Excel：「对应也生成周/月成交率，排版与另一侧一致」
- 用户：「封装成Skill」

### 数据层（extract_weekly_enhanced.py 是周月共享入口）

- **月成交率 `_conversion`**：本月笔数 ÷ 月客流文件（`7月客流（截止19日）.xlsx`），由 `extract_monthly.py` 算、结构同
- **周成交率 `_weekly_conversion`**：动态 **WK31 窗口（`WK28_BASE+21天` ~ `+27天` = 2026-07-27 ~ 2026-08-02，随 `max_date` 自动推进）** 笔数 ÷ `wk31周客流.xlsx` 客流，写在 `extract_weekly_enhanced.py` 的 4.9 块（在 `_conversion` 块之后、`# 5. 更新 weekly_analysis.json` 之前）。⚠️ 周成交率窗口必须与周销售/周同比窗口一致（均为 WK31），否则口径错位
- 两者都写入 `weekly_analysis.json`（周）/ `monthly_analysis.json`（月）
- **客流名→销售名映射** `FLOW_MAP`（扬州万象汇新店→扬州万象汇、淮安新亚广场→淮安新亚，其余同名）
- **统一结构**：`{store: {flow, cnt, rate}, '_region': {flow, cnt, rate}}`，rate = cnt/flow×100（flow=0 时 rate=0）

### 渲染层

**门店页成交率卡**
- 周：`wcvo=(DATA['_weekly_conversion']||{})`（`gen_weekly_p3.py`），标签 `ZK29周成交率`，sub `(cnt)笔 / (flow)客流`
- 月：`cvo=(DATA['_conversion']||{})`（`gen_monthly_p2.py`），标签 `月成交率`，sub `(cnt)笔/(flow)客流`
- 值 ≥ 区域 rate 标绿（class `green`），否则标红（class `red`）

**区域对比第三排（半宽并排）**
- 同个 `chart-grid` 内放两个 `chart-card`：长尾款 `cmpLtSample` + 成交率 `cmpConv`
- 周：`wcvo=(DATA['_weekly_conversion']||{})`；月：`cvo=(DATA['_conversion']||{})`
- `cmpConv` 图表：bar，label 周=`周成交率`/月=`月成交率`，tooltip 返回 `成交率: x% | 笔数n / 客流m`
- 汇总表追加「成交率」列（周=`周成交率`，月=`成交率`），≥区域标绿否则标红

### ⚠️ 关键坑 1：WK 动态替换误伤成交率标签（用 ZK29 占位符）

周看板 `gen_weekly_p3/p4.py` 末尾有 `.replace('WK29', f'WK{WK_NUM_INT}')` 把正文周号动态化。当 `max_date` 推进到 WK30 时，若正文写死 `"WK29周成交率"`，会被误改成 `"WK30周成交率"`——但**数据是按固定 WK29 窗口算的**，标签与数据错位会误导用户。

**修复（三处一致）**：
1. 成交率相关字面量用 `ZK29` 占位：区域对比 h3「各门店ZK29周成交率对比」、conc 文本「区域ZK29周成交率」、门店卡标签「ZK29周成交率」
2. 末尾 replace 链末尾追加 `.replace('ZK29','WK29')` 还原：
   ```python
   _html_out = HTML.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}').replace('ZK29','WK29')
   ```
3. 月看板（`gen_monthly_p2/p3.py`）不做 WK 替换，**无此坑**，直接写 `月成交率`

### ⚠️ 关键坑 2：gen_weekly_p4.py 中文是 \uXXXX 转义

`gen_weekly_p4.py` 用 raw string 存中文（如 `\u5404\u95e8\u5e97`），Edit 工具无法匹配中文。
- 改 **JS 正文** → 必须用 Python 脚本读取→字符串替换→写回
- 改 **Python 部分**（如末尾 replace 链、`WK_NUM_INT` 计算）→ 可直接 Edit

### ⚠️ 关键坑 3：区域对比「成交率」空值会整页崩空白

门店页成交率卡用 `cvoRate===null?'缺客流'` 做了空值保护，但**区域对比页**（gen_weekly_p4.py L92 / gen_monthly_p3.py L90）的 `conc()` 成交率文字段直接对 `convData` 里的 `x.v`（`rate`，可能为 `null`）调 `.toFixed(2)`，且 `avg=convData.reduce(...x.v...)`、`best/worst` 也假设非 null：

- 周页：该 `conc()` 在 `document.getElementById('content').innerHTML=html` **之前**执行 → 一抛错整块区域内容没写进 DOM → **区域对比 Tab 全空白**（用户报"没有数据"）。
- 月页：区域 `renderComparison` 内部有 try/catch 吞掉异常 → 同样空白但不报错。

**触发条件**：`_weekly_conversion` / `_conversion` 中门店 `flow=0`（周客流/月客流文件未加载，WK_FLOW_FILE 未更新）→ 所有 `rate=null`。

**修法（已落地）**：把区域成交率文字段改 null 安全——
- 列表：`convData.map(x=>x.s+(x.v===null?'N/A':x.v.toFixed(2)+'%'))`
- 机会点：`const valid=convData.filter(x=>x.v!==null); const avg=valid.length?valid.reduce(...)/valid.length:0; const below=valid.filter(x=>x.v<avg);` 且 best/worst 取自 `valid`（包在 `if(valid.length){...}` 内）
- 图表 tooltip 同改 `x.v===null?'缺客流':x.v.toFixed(2)+'%'`

**验证**：用 Node 桩（`document`/`Chart` 打桩，元素用普通对象而非 Proxy 以免 `._html` 被拦截）真实调用 `renderComparison()`，断言 content 长度 > 500 且含关键模块。Proxy 当 stub 时读 `._html` 会被 `get` 陷阱返回函数导致长度误判为 0，是个坑。

### 修复损坏渲染文件的锚点清理法

若重排脚本把 `gen_monthly_p3.py` / `gen_weekly_p4.py` 改坏（半宽块 `'<div...' +` 语法错、重复成交率 full 块、chart init 丢失），按以下步骤修：
1. 用**唯一注释锚点**切分文本（如 `// 门店WK29周成交率 & 长尾款 对比（半宽并排，第三排）`、孤立成交率块锚、`// 区域TOP10产品` 锚）
2. 删除重复块、重建 Row3（长尾款+成交率半宽并排）
3. **断言唯一性**后写回：`cmpLtSample` canvas 出现 1 次、`cmpConv` canvas 出现 1 次、`const cvo`/`const wcvo` 声明各 1 次
4. 跑 `auto_run_weekly.py` 的 node JS 校验，确保无语法错误再部署

### 验证清单

- [ ] 周/月成交率数据正确（区域值、各店值、笔数/客流对得上客流文件）
- [ ] 周看板显示「WK29周成交率」、无 `ZK29` 残留
- [ ] 区域对比第三排 = 长尾款 + 成交率 半宽并排
- [ ] 门店页有成交率卡片（周 ZK29 / 月 月成交率）
- [ ] node JS 校验通过，流水线重跑成功
- [ ] 部署 outputs_analysis，present_files 展示固定链接

## 参考资料

完整的JSON结构、计算公式、页面架构、脚本细节详见 `references/weekly_monthly_rules.md`。修改看板逻辑前务必加载该文件。详见该文件「十五、成交率模块与区域对比布局重排」。
