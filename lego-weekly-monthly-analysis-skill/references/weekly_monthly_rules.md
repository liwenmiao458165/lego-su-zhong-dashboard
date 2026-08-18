# 苏中周月维度分析看板 — 完整需求与逻辑知识库

> 文档版本：2026-07-19
> 涵盖脚本：12个脚本（4个周HTML生成 + 3个月HTML生成 + 4个JSON提取 + 1个JSON注入 + 1个编排器）
> 部署链接：https://320f470c4e3746d6840aca494a217a7b.app.codebuddy.work
> 页面标题：苏中周月维度分析

---

## 一、项目概述

### 1.1 目标
每周自动生成苏中区域7家门店的周维度+月维度交互分析HTML，包含同比、环比、产品结构、导购排名、样品销售等维度。

### 1.2 架构
```
auto_run_weekly.py（编排器）
  ├── 依次运行12个脚本
  ├── JS语法校验（Node.js new Function()）
  ├── 复制到 outputs_analysis/ 独立目录
  └── 部署到云端（独立链接）
```

### 1.3 脚本流水线（12步）
```
Step 1: HTML模板生成（4周 + 3月 = 7个）
  ├── gen_weekly_html.py    → 周分析HTML骨架+CSS+导航+工具函数
  ├── gen_weekly_p2.py      → 周分析辅助函数（格式化、区域计算等）
  ├── gen_weekly_p3.py      → 周门店详情页（renderStore函数）
  ├── gen_weekly_p4.py      → 周区域对比页（renderComparison+selectStore+init）
  ├── gen_monthly_html.py   → 月分析HTML骨架+CSS+导航+工具函数
  ├── gen_monthly_p2.py     → 月门店详情页（renderStore函数）
  ├── gen_monthly_p3.py     → 月区域对比页（renderComparison+selectStore+init）
Step 2: JSON数据提取（4个）
  ├── init_weekly_json.py      → 周JSON基础结构（daily数组+门店周数据）
  ├── extract_weekly_guide.py  → 导购周数据+门店周同比
  ├── extract_weekly_enhanced.py → 周环比+产品系列+TOP10+导购环比+样品+日指标
  ├── extract_monthly.py       → 月度全部数据
Step 3: JSON注入（1个）
  └── inject_json.py → 替换HTML中 `const DATA = {...}` 为实际数据
```

### 1.4 输出文件
| 文件 | 路径 | 用途 |
|------|------|------|
| `weekly_analysis.html` | `outputs/` → `outputs_analysis/` | 周维度交互分析 |
| `monthly_analysis.html` | `outputs/` → `outputs_analysis/` | 月维度交互分析 |
| `weekly_analysis.json` | `outputs/` | 周JSON数据 |
| `monthly_analysis.json` | `outputs/` | 月JSON数据 |
| `index.html` | `outputs_analysis/` | 入口页（周月选择） |

---

## 二、门店架构与特殊处理

### 2.1 门店列表（同日看板，7家固定顺序）
同日看板知识库的门店架构完全适用。

### 2.2 新店特殊处理（淮安万象城）
- **同比**：所有同比字段设为 `None`，前端显示"N/A"
- **同比口径排除**：区域合计的同比用6店口径（排除淮安万象城）
- **环比**：淮安万象城有上月数据，MoM不排除新店
- **null保护**：JS中所有 `yoy_pct.toFixed()` 需加null检查
  ```javascript
  (day.yoy_pct != null ? (day.yoy_pct >= 0 ? '+' : '') + day.yoy_pct.toFixed(1) + '%' : 'N/A')
  ```

---

## 三、数据源与文件位置

### 3.1 动态文件查找
周月脚本不依赖 auto_run.py 的正则替换，自行glob查找文件：
```python
# 历史CSV
csvs = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
# 今日xlsx
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'))
# 指标同期
TARGET_XLSX = sorted(glob.glob('/Users/a123/Desktop/。/*月指标同期.xlsx'), key=os.path.getmtime)[-1]
# 导购指标
GUIDE_XLSX = sorted(glob.glob('企微缓存/苏中区域*月员工指标分解.xlsx'), key=os.path.getmtime)[-1]
# 长尾款SKU
LT_SKU_XLSX = sorted(glob.glob('企微缓存/长尾款sku明细.xlsx'), key=os.path.getmtime)[-1]
```

### 3.2 数据保留范围
- `init_weekly_json.py`：保留近3个月数据
- `extract_weekly_guide.py`：保留近4周数据（`WK_START - 3weeks`）
- `extract_weekly_enhanced.py`：保留近4周数据
- `extract_monthly.py`：保留近3个月数据

---

## 四、数据清洗规则

### 4.1 核心规则（同日看板）
去重key、排除赠品/样品/零金额、排除4类人员、门店映射 — 与日看板完全一致。

### 4.2 样品数据分离
- 在排除规则前先保存到 `sample_df`
- 条件：`下单店仓名称` 含"样品" + 排除赠品 + 排除零金额
- 样品数据保留 `牌价额` 字段

### 4.3 日期范围过滤
- init_weekly_json：`_date_3m_ago = max_date - 3months`
- extract_weekly_enhanced：`_data_cutoff = WK_START - 3weeks`
- extract_monthly：`_date_3m_ago = max_date - 3months`
- 样品数据同步做相同的日期范围过滤

---

## 五、JSON数据结构

### 5.1 weekly_analysis.json

```json
{
  "扬州万象汇": {
    "daily": [
      {"date": "7/13", "weekday": "周一", "total": ..., "wos": ..., "jl": ..., "lt": ..., "lt_qty": ..., "tx_cnt": ..., "jd": ..., "kdj": ..., "target": ..., "rate": ..., "yoy_amt": ..., "yoy_pct": ...}
    ],
    "wk_amt": ..., "wk_wos": ..., "wk_jl": ..., "wk_lt": ..., "wk_total": ..., "month_amt": ..., "wk28_amt": ..., "wow": ..., "wk_yoy": ...
  },
  "_region": {
    "wk_amt": ..., "wk28_amt": ..., "wow": ..., "wk_yoy": ..., "wk_atv": ..., "daily": [...],
    // 其他区域KPI字段（来自dashboard_latest.json）
  },
  "_guides": {
    "扬州万象汇": [
      {"name": "...", "store": "...", "title": "...", "target": ..., "wk_sales": ..., "wk_qty": ..., "wk_cnt": ..., "wk_atv": ..., "wk_jd": ..., "wk_wos": ..., "wk_lt": ..., "wk_lt_qty": ..., "wk_rate": ..., "wk28_sales": ..., "wow": ...}
    ]
  },
  "_series": {"扬州万象汇": {"系列A": 金额, ...}, "_region": {...}},
  "_products": {"扬州万象汇": {"商品名": {"amt": ..., "qty": ...}, ...}, "_region": {...}},
  "_sample": {"扬州万象汇": {"wk_qty": ..., "wk_retail": ..., "wk_amt": ..., "month_qty": ..., "month_retail": ..., "month_amt": ...}},
  "_sample_region": {"wk_qty": ..., "wk_retail": ..., "wk_amt": ..., "month_qty": ..., "month_retail": ..., "month_amt": ...}
}
```

### 5.2 monthly_analysis.json

```json
{
  "扬州万象汇": {
    "month_amt": ..., "mom": ..., "mom_prev": ..., "yoy": ..., "yoy_prev": ..., "target": ..., "month_rate": ..., "month_cnt": ..., "month_qty": ..., "month_atv": ..., "month_jd": ..., "month_wos": ..., "month_jl": ..., "month_lt": ..., "month_lt_qty": ..., "daily": [...], "daily_prev": [...]
  },
  "_region": {
    "month_amt": ..., "mom_prev": ..., "yoy_prev": ..., "yoy": ..., "target": ..., "month_rate": ..., "month_cnt": ..., "month_qty": ..., "month_atv": ..., "month_jd": ..., "month_wos": ..., "month_jl": ..., "month_lt": ..., "daily": [...]
  },
  "_guides": {
    "扬州万象汇": [
      {"name": "...", "store": "...", "target": ..., "m_sales": ..., "m_qty": ..., "m_cnt": ..., "m_atv": ..., "m_jd": ..., "m_wos": ..., "m_lt": ..., "m_lt_qty": ..., "m_rate": ..., "m_prev": ..., "mom": ...}
    ]
  },
  "_series": {"扬州万象汇": {...}, "_region": {...}},
  "_products": {"扬州万象汇": {...}, "_region": {...}},
  "_sample": {"扬州万象汇": {"month_qty": ..., "month_retail": ..., "month_amt": ...}},
  "_sample_region": {"month_qty": ..., "month_retail": ..., "month_amt": ...},
  "_meta": {"report_date": "7月19日", "report_day": 19, "report_day_index": 20, "max_date": "2026-07-19", "store_order": [...]}
}
```

---

## 六、周分析计算逻辑

### 6.1 init_weekly_json.py — 基础结构

| 计算项 | 公式 |
|--------|------|
| WK周编号 | `WK_NUM_INT = 28 + (max_date - WK28_BASE).days // 7` |
| WK起始日 | `WK_START = WK28_BASE + Timedelta(weeks=weeks_since_wk28)` |
| 门店日销售 | `wk29_df[门店==store][日期==d]['结算金额'].sum()` |
| 日WOS | `day_data[支付方式含'WOS']['结算金额'].sum()` |
| 日即时零售 | `支付方式含'美团'/'淘宝'/'京东' OR 特殊销售类型含'饿了么'` |
| 日长尾款 | `day_data[乐高编号∈lt_skus]['结算金额'].sum()` |
| 日长尾款件数 | `day_lt[结算金额>0]['数量'].sum()` |
| 日笔数 | `day_pos['单号'].nunique()`（正单=结算金额>0） |
| 日连带率 | `正单数量 / 正单单数` |
| 日客单价 | `日金额 / 正单单数` |
| 周累计 | `wk29_df[门店==store]['结算金额'].sum()` |
| 月累计 | `all_df[门店==store][月份==REPORT_MONTH]['结算金额'].sum()` |
| 区域数据 | 从 `dashboard_latest.json` 复制（引用日看板的区域数据） |

### 6.2 extract_weekly_guide.py — 导购+同比

| 计算项 | 公式 |
|--------|------|
| 导购周销售 | `g_wk['结算金额'].sum()` |
| 导购周笔数 | `pos['单号'].nunique()` |
| 导购周件数 | `pos['数量'].sum()` |
| 导购客单价 | `wk_sales / wk_cnt` |
| 导购连带率 | `wk_qty / wk_cnt` |
| 导购WOS | `g_wk[支付方式含'WOS']['结算金额'].sum()` |
| 导购长尾款 | `g_wk[乐高编号∈lt_skus]['结算金额'].sum()` |
| 导购长尾款件数 | `g_lt[结算金额>0]['数量'].sum()` |
| 导购月完成率 | `wk_sales / 月指标 × 100` |
| 门店周同比 | `(wk_actual - wk_yoy_last) / wk_yoy_last × 100` |
| 区域周同比 | `(6店wk_actual - region_wk_yoy) / region_wk_yoy × 100`（排除淮安万象城） |

#### 同期数据读取（指标同期Excel）
- 行范围：rows 13-20（门店+区域）
- WK起始列：`WK_COL_START = WK_START.day + 1`（同一月内）
  - 跨月时 `WK_COL_START = 2`
- WK结束列：`WK_COL_END = REPORT_DAY_INDEX`
- 周同比值 = `sum(col(WK_COL_START) : col(WK_COL_END))`（去年同期该周的累计）
- 列索引安全：`min(j, target_df.shape[1])`

### 6.3 extract_weekly_enhanced.py — 增强数据

| 计算项 | 公式 |
|--------|------|
| 上周(WK28)数据 | `WK_PREV_START = WK_START - 7days`, `WK_PREV_END = WK_START - 1day` |
| 门店周环比 | `(WK29金额 - WK28金额) / WK28金额 × 100`（WK28=0时设None） |
| 区域周环比 | `(region_wk29 - region_wk28) / region_wk28 × 100` |
| 产品系列结构 | `wk29_df.groupby('产品大类')['结算金额'].sum()`（按金额降序） |
| 产品TOP10 | `wk29_pos.groupby('商品名称')['结算金额'].sum()` + `'数量'.sum()`（按金额降序） |
| 导购周环比 | `(WK29_sales - WK28_sales) / WK28_sales × 100` |
| 导购匹配 | 主名 + `长期兼职-`前缀 + 支援名映射 |
| 李婷1/黄小莉1指标 | 从指标表中复制李婷/黄小莉的指标，删除李婷/黄小莉本体 |
| 日指标 | `month_target / DAYS_IN_MONTH`（日均分配） |
| 日同比 | `(day.total - day.yoy_amt) / day.yoy_amt × 100`（yoy_amt来自同期Excel） |
| 日完成率 | `day.total / day.target × 100` |
| 样品周数据 | `sample_df[门店==store][WK_START <= 日期 <= max_date]` |
| 样品月数据 | `sample_df[门店==store][month_start <= 日期 <= max_date]` |
| 样品件数 | `pos['数量'].sum()`（仅正单） |
| 样品零售价 | `pos['牌价额'].sum()`（仅正单） |
| 样品结算金额 | `sample_df['结算金额'].sum()`（含退货） |
| 区域日销售数组 | 每日汇总7店，同比排除淮安万象城 |

#### 同期Excel读取（每日数据）
- 行范围：rows 13-19（门店）+ row 20（区域）
- 每日同期：`col(j) for j in range(2, ncols)`
- `store_daily_yoy[门店名] = [float(col_j) for j in range(2, ncols)]`

---

## 七、月分析计算逻辑（extract_monthly.py）

### 7.1 月度门店KPI

| 指标 | 公式 | 说明 |
|------|------|------|
| 月累计金额 | `cur_month_df[门店==store]['结算金额'].sum()` | |
| 月环比(MoM) | `(月累计 - 上月同期) / 上月同期 × 100` | 上月同期=上月1号到上月report_day号 |
| 月同比(YoY) | `(月累计 - 去年同期月累计) / 去年同期月累计 × 100` | 新店为None |
| 月指标完成率 | `月累计 / 月指标 × 100` | |
| 月笔数 | `pos['单号'].nunique()` | 正单 |
| 月件数 | `pos['数量'].sum()` | 正单 |
| 月客单价 | `月金额 / 月笔数` | |
| 月连带率 | `月件数 / 月笔数` | |
| 月WOS | `cur_month[支付方式含'WOS']['结算金额'].sum()` | |
| 月即时零售 | 同渠道口径（美团/饿了么/淘宝/京东） | |
| 月长尾款 | `cur_month[乐高编号∈lt_skus]['结算金额'].sum()` | |
| 月长尾款件数 | `lt[结算金额>0]['数量'].sum()` | |

### 7.2 上月同期区间
```python
_prev_month_start = _month_start - DateOffset(months=1)
jun_same_period = prev_month_df[prev_month_df['销售日期_dt'].dt.day <= report_day]
```
- 上月同时间区间：上月1号到上月report_day号

### 7.3 区域合计口径
- **月环比**：7店（含淮安万象城，有上月数据）
- **月同比**：6店（排除淮安万象城，无同期数据）

### 7.4 日销售趋势数组
```json
"daily": [
  {"day": 1, "amt": ..., "target": ..., "rate": ..., "yoy_amt": ..., "yoy_pct": ..., "lt": ..., "lt_qty": ..., "wos": ..., "jl": ..., "tx_cnt": ...}
]
```
- 日指标 = `月指标 / DAYS_IN_MONTH`
- 日同比 = `(日金额 - 去年同日) / 去年同日 × 100`（新店为None）
- 日完成率 = `日金额 / 日指标 × 100`
- 同时构建 `daily_prev` 数组（上月同日金额）

### 7.5 样品月数据
- 件数/零售价/结算金额
- 区域汇总

---

## 八、导购数据处理

### 8.1 导购姓名映射（同日看板）
```python
normalize_guide_name(name):
    # 去掉"长期兼职-"前缀
    # 扬万支援销售 → 李婷1
    # 泰州万象城-支援2 / 泰州万象城支援2 → 黄小莉1
```

### 8.2 周导购处理差异（extract_weekly_enhanced.py）
- **李婷1/黄小莉1指标创建**：从指标表中复制李婷/黄小莉的指标
- **删除李婷/黄小莉本体**：`guide_info.pop('李婷')` / `guide_info.pop('黄小莉')`
- **导购匹配函数**：`get_match_names(gname)` 返回所有可能的姓名变体
- **周环比计算**：同时计算WK28和WK29的数据，得出个人环比

### 8.3 月导购处理（extract_monthly.py）
- 同样的映射规则和指标创建
- 月环比：本月 vs 上月同时间区间
- 月完成率：月销售 / 月指标 × 100

---

## 九、周分析页面结构

### 9.1 页面架构
```
weekly_analysis.html（单文件SPA）
  ├── Header（标题：苏中区域 WK29 门店周维度交互分析）
  ├── Store Tabs（区域对比 + 7家门店）
  ├── Container（动态渲染）
  │   ├── 区域对比页（renderComparison）
  │   │   ├── 9个KPI卡片（周累计/环比/同比/连带率/客单价/WOS/长尾款/样品）
  │   │   ├── 周累计&完成率对比图（双轴：柱=金额+线=完成率%，按完成率排序）
  │   │   ├── 周同比对比图
  │   │   ├── 客单价&连带率对比图
  │   │   ├── WK28 vs WK29环比图
  │   │   ├── WOS&即时零售对比图
  │   │   ├── 区域产品系列结构图
  │   │   ├── 长尾款&样品周累计对比图
  │   │   ├── 区域TOP10产品排名图
  │   │   ├── 区域周维度汇总表
  │   │   └── 全区域导购周销售排名图
  │   └── 门店详情页（renderStore）
  │       ├── 8个KPI卡片（周累计/环比/同比/连带率/客单价/WOS/长尾款/即时零售）
  │       ├── KPI汇总机会点（concKPI）
  │       ├── 产品系列环形图（TOP10 + "其他"）
  │       ├── 能力雷达图（vs区域均值）
  │       ├── WK28 vs WK29环比图
  │       ├── TOP10产品排名图
  │       ├── WOS&即时零售周累计对比图（vs区域）
  │       ├── 导购周销售+WOS+环比图
  │       ├── 导购连带率+客单价图
  │       ├── 导购长尾款分析图
  │       └── 导购周维度明细表
  └── Footer
```

### 9.2 门店详情页KPI卡片（最新版）
| 卡片 | 标签 | 值 | 子文本 |
|------|------|-----|--------|
| 1 | WK29周累计 | 金额(模块色) | vs WK28 环比 |
| 2 | 周环比 | 百分比 | vs WK28 金额 |
| 3 | 周同比 | 百分比(或N/A) | vs去年同期 |
| 4 | 周连带率 | 数值(蓝) | 区域连带率 |
| 5 | 周客单价 | ¥金额(蓝) | 区域客单价 |
| 6 | WOS周累计 | 金额(橙) | 占比% |
| 7 | 长尾款周累计 | 金额(紫) | 占比% |
| 8 | 即时零售周累计 | 金额(蓝) | 占比% |

> ⚠️ 笔数已改为连带率（2026-07-19变更）
> ⚠️ 日均数据已移除（2026-07-19变更）

### 9.3 区域对比页KPI卡片（最新版）
| 卡片 | 标签 | 值 |
|------|------|-----|
| 1 | 区域WK29周累计 | 金额 |
| 2 | 区域周环比 | 百分比 |
| 3 | 区域周同比 | 百分比(或N/A) |
| 4 | 区域周连带率 | 数值(蓝) |
| 5 | 区域周客单价 | ¥金额(蓝) |
| 6 | 区域WOS周累计 | 金额(橙) |
| 7 | 区域长尾款周累计 | 金额(紫) + 件数 |
| 8 | 区域样品周累计 | 金额(紫) + 件数 |

---

## 十、月分析页面结构

### 10.1 页面架构
```
monthly_analysis.html（单文件SPA）
  ├── Header（标题：苏中区域 月维度交互分析）
  ├── Store Tabs（区域对比 + 7家门店）
  ├── Container
  │   ├── 区域对比页（renderComparison）
  │   │   ├── KPI卡片
  │   │   ├── 月累计&完成率对比图（双轴：柱=金额+线=完成率%，按完成率排序）
  │   │   ├── 月同比对比图
  │   │   ├── 客单价&连带率对比图
  │   │   ├── 月环比对比图
  │   │   ├── WOS&即时零售月累计对比图
  │   │   ├── 区域产品系列结构图
  │   │   ├── 长尾款&样品月累计对比图
  │   │   ├── 区域TOP10产品月销排名图
  │   │   ├── 区域月维度汇总表
  │   │   └── 全区域导购月销售排名图
  │   └── 门店详情页（renderStore）
  │       ├── KPI卡片
  │       ├── KPI汇总机会点（concKPI）
  │       ├── 产品系列环形图（TOP10 + "其他"）
  │       ├── WOS&即时零售月累计对比图（vs区域）
  │       ├── 导购月销售+WOS+环比图
  │       ├── 导购连带率+客单价图
  │       ├── 导购长尾款分析图
  │       └── 导购月维度明细表
  └── Footer
```

### 10.2 门店详情页KPI卡片（最新版）
| 卡片 | 标签 | 值 | 子文本 |
|------|------|-----|--------|
| 1 | 月累计 | 金额 | 完成率 |
| 2 | 月环比 | 百分比 | 上月同期金额 |
| 3 | 月同比 | 百分比(或N/A) | 去年同期金额 |
| 4 | 月连带率 | 数值(蓝) | 区域连带率 |
| 5 | 月客单价 | ¥金额(蓝) | 区域客单价 |
| 6 | WOS月累计 | 金额(橙) | 占比% |
| 7 | 长尾款月累计 | 金额(紫) | 件数+占比% |
| 8 | 即时零售月累计 | 金额(蓝) | 占比% |

> ⚠️ 笔数已改为连带率（2026-07-19变更）
> ⚠️ 日均数据/每日明细数据表已移除（2026-07-19变更）

---

## 十一、inject_json.py — JSON注入机制

### 11.1 工作原理
- 在HTML中查找 `const DATA = ` 标记
- 用花括号匹配算法定位JSON对象边界
- 处理嵌套对象、字符串转义
- 用 `json.dumps(data, ensure_ascii=False, separators=(',',':'))` 生成紧凑JSON
- 替换 `html[start:end]` 为新JSON字符串

### 11.2 注入目标
- `weekly_analysis.html` ← `weekly_analysis.json`
- `monthly_analysis.html` ← `monthly_analysis.json`

---

## 十二、JS语法校验

### 12.1 校验逻辑
```javascript
// Node.js: new Function() 检查
const js = '<script>标签内的JS内容（排除const DATA行）';
const stubs = 'const Chart=function(){};const document={getElementById:()=>{},querySelectorAll:()=>[],addEventListener:()=>{}};';
new Function(stubs + js);  // 如果有语法错误会抛异常
```

### 12.2 常见错误
- `const` 变量重复声明（如 `const sd` 在同一函数内声明两次）
- 捕获后必须修复才能部署，否则页面空白

### 12.3 校验时机
- 在所有12脚本运行完成后，部署前执行
- 校验失败 → 打印错误信息 → `sys.exit(1)` → 不部署

---

## 十三、auto_run_weekly.py — 编排器

### 13.1 运行流程
```
1. 检测数据文件（仅用于日志确认，脚本自行glob）
2. 依次运行12个脚本
3. JS语法校验（Node.js new Function()）
4. 复制到 outputs_analysis/ 独立目录
```

### 13.2 脚本运行顺序
```python
ANALYSIS_SCRIPTS = [
    # Step 1: 生成HTML模板（必须先于JSON提取）
    ('gen_weekly_html.py', '生成周分析HTML模板(p1)'),
    ('gen_weekly_p2.py', '生成周分析HTML(p2)'),
    ('gen_weekly_p3.py', '生成周分析HTML(p3-门店)'),
    ('gen_weekly_p4.py', '生成周分析HTML(p4-区域对比)'),
    ('gen_monthly_html.py', '生成月分析HTML模板(p1)'),
    ('gen_monthly_p2.py', '生成月分析HTML(p2-门店)'),
    ('gen_monthly_p3.py', '生成月分析HTML(p3-区域对比)'),
    # Step 2: 提取JSON数据
    ('init_weekly_json.py', '初始化周JSON'),
    ('extract_weekly_guide.py', '提取周导购数据'),
    ('extract_weekly_enhanced.py', '提取周增强数据'),
    ('extract_monthly.py', '提取月度数据'),
    # Step 3: 注入JSON到HTML模板
    ('inject_json.py', '注入JSON到HTML'),
]
```

### 13.3 输出检查
- 检查 `weekly_analysis.html` 和 `monthly_analysis.html` 是否存在
- 显示文件大小
- 复制到 `outputs_analysis/`

### 13.4 失败处理
- 任何脚本返回非0退出码 → 记录失败 → 继续运行其他脚本
- 全部完成后汇总失败脚本列表
- JS校验失败 → 直接退出，不部署

---

## 十四、配色与可视化规则

### 14.1 模块配色（同日看板）
| 模块 | 图标 | 颜色 | 色值 |
|------|------|------|------|
| 门店 | 🏪 | 紫色 | #7c3aed |
| 长尾款 | 📦 | 橙色 | #ea580c |
| 导购 | 👩 | 绿色 | #059669 |
| WOS | 🔗 | 蓝色 | #0284c7 |
| 即时零售 | ⚡ | 蓝色 | #0284c7 |
| 样品 | 🧪 | 紫色 | #9333ea |

### 14.2 门店颜色映射（Tab高亮）
| 门店 | 颜色 |
|------|------|
| 扬州万象汇 | #7c3aed (紫) |
| 扬州京华城 | #0ea5e9 (天蓝) |
| 扬州江都金鹰 | #059669 (绿) |
| 泰州万象城 | #dc2626 (红) |
| 宿迁宝龙 | #ea580c (橙) |
| 淮安新亚 | #0891b2 (青) |
| 淮安万象城 | #9333ea (深紫) |
| 区域对比 | #1e293b (深灰) |

### 14.3 数据颜色规则
- 同比/环比：负值红（跌），正值绿（涨），null→N/A
- 完成率数据条：≥50% 绿色，<50% 红色
- 值≤0标红

### 14.4 图表配色
- 系列环形图：使用预设色板 `SC`（前10个系列 + "其他"）
- 每日趋势：门店主色柱状图 + 绿色完成率线 + 红色同比虚线（日趋势图已移除，区域对比页保留）
- 雷达图：门店色 vs 区域灰

### 14.5 区域对比页双轴图表（周/月通用）

门店累计销售&完成率对比图（`cmpAmt`）使用双轴格式：

```javascript
// 数据准备：按完成率排序
const storeRates = STORE_ORDER.map(s => {
  const d = DATA[s];
  return { s, v: d.wk_amt, target: d.wk_target || d.target, rate: d.wk_rate };
}).sort((a, b) => b.rate - a.rate);

// 图表配置
charts.cmpAmt = new Chart(document.getElementById('cmpAmt'), {
  data: {
    labels: storeRates.map(x => x.s + '\n' + x.rate.toFixed(1) + '%'),
    datasets: [
      { type: 'bar', label: '周累计', data: storeRates.map(x => x.v),
        backgroundColor: storeRates.map(x => STORE_COLORS[x.s] + 'cc'),
        borderColor: storeRates.map(x => STORE_COLORS[x.s]),
        borderWidth: 1, borderRadius: 6, yAxisID: 'y' },
      { type: 'line', label: '完成率', data: storeRates.map(x => x.rate),
        borderColor: '#dc2626', backgroundColor: 'transparent',
        borderWidth: 2, pointRadius: 5, pointBackgroundColor: '#dc2626',
        tension: 0.3, yAxisID: 'y1' }
    ]
  },
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top' },
      tooltip: {
        callbacks: {
          label: ctx => {
            const x = storeRates[ctx.dataIndex];
            if (ctx.dataset.type === 'line')
              return '完成率: ' + x.rate.toFixed(1) + '%';
            return '销售: ' + fmtAmt(x.v) + ' | 指标: ' + fmtAmt(x.target)
                   + ' | 完成率: ' + x.rate.toFixed(1) + '%';
          }
        }
      }
    },
    scales: {
      y: { position: 'left', title: { display: true, text: '累计(¥)' },
           ticks: { callback: v => fmtAmt(v) } },
      y1: { position: 'right', title: { display: true, text: '完成率(%)' },
            grid: { drawOnChartArea: false },
            ticks: { callback: v => v + '%' } }
    }
  }
});
```

**关键设计点**：
- **双轴**：y（左轴）=金额柱状图，y1（右轴）=完成率折线
- **网格线**：y1设置`drawOnChartArea:false`避免双网格线重叠
- **排序**：按完成率从高到低，x轴标签含门店名+完成率%
- **tooltip**：柱hover显示金额+指标+完成率；线hover显示完成率
- **字段差异**：周用`wk_amt`/`wk_rate`/`wk_target`，月用`month_amt`/`month_rate`/`target`
- **实现位置**：周=`gen_weekly_p4.py`（⚠️unicode转义，需Python脚本替换），月=`gen_monthly_p3.py`

---

## 十五、自动化运行流程

### 15.1 手动触发
- 用户说"出看板" → 只运行 auto_run.py（主看板）
- 用户说"更新周月分析" → 运行 auto_run_weekly.py + 部署 outputs_analysis

### 15.2 定时自动化
| 时间 | 执行内容 | 说明 |
|------|---------|------|
| 13:00 | auto_run.py → 部署 | 午间主看板 |
| 18:00 | auto_run.py → 部署 | 晚间主看板 |
| 22:30 | auto_run.py + auto_run_weekly.py → 部署两个看板 | 最终版 |

### 15.3 周月分析更新规则
- **周月分析看板等我每晚22:30更新**，未更新就停留在前一天的周和月数据维度
- 周日/节假日不提前更新周月分析

### 15.4 部署
- 主看板：`outputs/` → CloudStudio（固定链接）
- 周月分析：`outputs_analysis/` → CloudStudio（独立固定链接）
- 入口页：`index.html` 提供周/月选择导航

---

## 十六、关键注意事项

### 16.1 JSON依赖链
```
init_weekly_json.py
  ← 读取 dashboard_latest.json（区域数据）
  → 输出 weekly_analysis.json（基础结构）

extract_weekly_guide.py
  ← 读取 weekly_analysis.json
  → 更新导购数据+门店周同比

extract_weekly_enhanced.py
  ← 读取 weekly_analysis.json
  → 更新环比+系列+产品+样品+日指标+日同比

inject_json.py
  ← 读取 weekly_analysis.json + monthly_analysis.json
  → 替换HTML中的const DATA
```

### 16.2 HTML生成必须在JSON提取之前
- `inject_json.py` 依赖HTML中 `const DATA = {...}` 占位符
- 所以 `gen_*.py` 必须先运行，创建带占位符的HTML
- 然后 `init/extract_*.py` 创建JSON数据
- 最后 `inject_json.py` 替换占位符

### 16.3 null保护（淮安万象城）
- 所有JS中 `yoy_pct` 相关的 `.toFixed()` 必须加null检查
- 模板中使用 `isNA()` 函数判断
- `isNA(v) = v === null || v === undefined`

### 16.4 长尾款SKU读取易错点
- `extract_weekly_guide.py` 中有一个bug：使用 `iloc[:,0]` 读取SKU（那是序号列）
- 正确应使用 `lt_sku_df['乐高编号']`
- ⚠️ init_weekly_json 和 extract_weekly_enhanced 使用正确列名

### 16.5 导购指标特殊处理
- 指标表中只有李婷/黄小莉，需手动创建李婷1/黄小莉1条目
- 李婷/黄小莉本体需删除（他们的销售数据已排除，指标分配给1版本）
- 导购指标列名：`姓名`、`员工个人指标`、`门店`

### 16.6 月环比口径差异
- 区域月环比：7店（淮安万象城有上月数据）
- 区域月同比：6店（淮安万象城无同期数据）
- 门店月环比：7店（淮安万象城环比=本月 vs 上月同期）
- 门店月同比：6店（淮安万象城为None/N/A）

### 16.7 样品数据口径
- 样品件数：仅正单（`结算金额>0`）的 `数量.sum()`
- 样品零售价：仅正单的 `牌价额.sum()`
- 样品结算金额：`结算金额.sum()`（含退货）
- 样品数据从主数据分离，独立统计

---

## 十七、模块结论与机会点

### 17.1 设计目标
每个图表/表格模块下方自动生成两个区域：
1. **📋 结论**：基于当前门店/区域数据生成的一行结论文本
2. **💡 需关注的机会点**：基于数据阈值判断，给出可落地的业务建议

### 17.2 CSS样式

```css
/* 结论区 */
.conclusion {
  font-size: 13px;
  color: #475569;
  margin-top: 10px;
  padding: 0;
  background: #f8fafc;
  border-radius: 6px;
  border-left: 3px solid var(--purple);
  line-height: 1.5;
  overflow: hidden;
}
.conc-title { padding: 8px 12px 4px 12px; font-weight: 700; color: #1e293b; }
.conc-body { padding: 0 12px 8px 12px; }

/* 机会点区 */
.conc-opportunity {
  padding: 6px 12px;
  background: #fef3c7;       /* 琥珀色背景 */
  border-top: 1px solid #f59e0b;
  font-size: 12px;
  color: #92400e;            /* 深琥珀文字 */
  line-height: 1.6;
}
.conc-opportunity .conc-label { font-weight: 700; color: #b45309; }
```

CSS 定义位置：
- 周分析：`gen_weekly_html.py`
- 月分析：`gen_monthly_html.py`

### 17.3 `conc()` 函数签名

```javascript
function conc(c, o) {
  // o 可选 — 不传时只渲染结论区（向后兼容）
  if (!o) {
    return '<div class="conclusion">' +
           '<div class="conc-title">📋 结论</div>' +
           '<div class="conc-body">' + c + '</div>' +
           '</div>';
  }
  // 有机会点 — 渲染双区域
  return '<div class="conclusion">' +
         '<div class="conc-title">📋 结论</div>' +
         '<div class="conc-body">' + c + '</div>' +
         '<div class="conc-opportunity">' +
         '<span class="conc-label">💡 需关注的机会点</span><br>' + o +
         '</div></div>';
}
```

### 17.4 `conc*()` 函数族

8个函数，每个对应一个图表模块。定义位置：
- 周分析：`gen_weekly_p2.py`
- 月分析：`gen_monthly_html.py`

| 函数名 | 对应模块 | 机会点逻辑 |
|--------|---------|-----------|
| `concKPI(s)` | KPI卡片汇总 | 门店所有KPI（完成率/同比/环比/客单价/连带率/WOS占比/长尾款占比）逐项vs区域，低于的全部列出具体差距 |
| `concSeries(sr, s)` | 产品系列环形图（TOP10） | TOP10逐项标注占比 / TOP3集中度 / CITY·TECHNIC·NINJAGO是否在TOP3 / MINECRAFT是否在TOP5 / BOTANICALS植物系列占比（指标4.1%，低于指标标注机会点） |
| `concRadar(s)` | 能力雷达图 | 5维度逐项vs区域均值，标注具体差距(¥/比率/占比%) |
| `concLongTail(d, s)` | 长尾款分析 | 长尾款占比 vs 区域均值 + 金额 vs 区域门店均值（<70%标注） |
| `concWow(d, s)` / `concMom(d, s)` | 周环比 / 月环比 | 环比幅度阈值 |
| `concTop(p, s)` | TOP10产品排名 | TOP1占比>25%提醒；当月新品是否进入TOP10；区域TOP10产品门店覆盖检查 |
| `concGuideSales(g, s)` | 导购销售排名 | 门店人均+区域人均双对比，低于70%列出导购+金额 |
| `concGuideJdKdj(g, s)` | 导购连带+客单价 | vs区域均值，显示¥差值和比率差值 |
| `concGuideLt(g, s)` | 导购长尾款分析 | 门店人均对比 + 零长尾款销售导购标注 |

### 17.5 机会点生成规则（数据驱动，均值对比优先）

**核心原则**：优先与区域/门店均值对比，显示具体差距数值；绝对阈值仅作为辅助参考。

| 维度 | 对比基准 | 阈值条件 | 生成的建议文本 |
|------|---------|---------|--------------|
| KPI汇总 | 区域整体 | <区域值 | "低于区域整体的KPI：完成率XX%(区域XX%,差Xpct)；客单价¥XX(区域¥XX,差¥XX)等，需逐项分析差距原因" |
| 完成率 | 区域完成率 | <区域-5pct | "完成率低于区域均值X个百分点，需检视目标拆解和执行节奏" |
| 完成率 | 绝对值 | <40% | "完成率偏低（X%），建议加强周末高峰冲量" |
| 完成率 | 绝对值 | 40%-54% | "完成率处于中游（X%），关注后续销售节奏" |
| 完成率 | 绝对值 | ≥55% | "完成率表现良好（X%），建议保持势头" |
| 长尾款 | 区域门店均值 | <均值×0.7 | "长尾款金额低于区域门店均值（¥XX），需重点跟进长尾款推荐" |
| 雷达维度 | 区域均值 | 低于区域 | 逐维度标注差距：客单价差¥XX / 连带率差X.XX / 占比差X.X% |
| 占比 vs 区域 | <区域-2% | "XX占比低于区域平均（XX%），需加强该渠道/品类的推广" |
| 占比 vs 区域 | >区域+3% | "XX占比高于区域平均（XX%），关注是否过度依赖单一渠道" |
| 占比 vs 区域 | 区域±2% | "XX占比与区域持平" |
| 环比 | 绝对值 | <-8% | "环比下滑X%，幅度较大，需排查原因并制定改善方案" |
| 环比 | 绝对值 | -8%~-3% | "环比小幅下滑X%，持续关注趋势" |
| 环比 | 绝对值 | -3%~3% | "环比基本持平，处于稳定区间" |
| 环比 | 绝对值 | ≥3% | "环比增长X%，势头良好" |
| TOP1依赖度 | 绝对值 | >25% | "TOP1产品占比X%，过度依赖单一产品，建议丰富推荐结构" |
| 导购销售 | 门店人均×0.7 | <门店均值×0.7 | "XX(¥XX)低于门店均值¥XX，需重点跟进销售技巧辅导" |
| 导购销售 | 区域人均×0.7 | <区域均值×0.7 | "XX低于区域导购均值¥XX，建议安排导师带教和排班优化" |
| 导购完成率 | 绝对值 | <30% | "X位导购完成率不足30%，需一对一辅导" |
| 导购连带 | 区域均值 | <区域-0.1 | "连带率低于区域均值X.XX，建议强化搭配推荐话术培训" |
| 导购客单价 | 区域均值85% | <区域×0.85 | "客单价低于区域¥XX，建议引导客户关注高价套装" |
| 导购长尾 | 门店人均 | 零销售/占比<3% | "XX本周零长尾款销售，需重点跟进79款SKU推荐" |

### 17.6 当月新品检测逻辑

从长尾款SKU文件的"上市时间"列筛选 >= 2026-05-01 的新品，存入 `_new_products` 字段（extract_weekly_enhanced.py + extract_monthly.py）。

`concTop` 中检查当月新品是否进入门店TOP10（双向模糊匹配 `tn.includes(np)||np.includes(tn)`）：
- 有→"当月新品「XX」进入TOP10，表现突出，建议加大陈列和推荐力度"
- 无→"当月新品无一款进入TOP10，需关注新品陈列位置和推荐话术"

### 17.7 inline `conc()` 调用

除了 `conc*()` 函数外，部分模块直接在 `renderStore()` / `renderComparison()` 函数内内联调用 `conc()`：

**周分析**：
- `gen_weekly_p3.py`：1处（WOS&即时零售模块）
- `gen_weekly_p4.py`：9处（各门店对比、同比、环比、客单价、WOS、系列、长尾款样品、TOP10、导购排名）

**月分析**：
- `gen_monthly_p2.py`：1处（WOS&即时零售模块）
- `gen_monthly_p3.py`：9处（同周分析结构）

### 17.8 IIFE动态计算（均值对比）

区域对比页的机会点需要遍历多门店数据动态计算，使用立即执行函数表达式（IIFE）。
核心逻辑改为与区域均值对比，显示具体差距数值：

```javascript
// 示例：门店完成率 vs 区域均值
conc('结论文本', function() {
  const sr = STORE_ORDER.map(s => ({
    s, amt: DATA[s].wk_amt, wt: DATA[s].daily.reduce((a,x)=>a+(x.target||0),0),
    rate: wt>0 ? DATA[s].wk_amt/wt*100 : 0
  })).sort((a,b) => b.rate - a.rate);
  const avgRate = sr.reduce((a,x) => a+x.rate, 0) / sr.length;
  const below = sr.filter(x => x.rate < avgRate);
  let o = '';
  if (below.length > 0)
    o += '低于区域均值（' + avgRate.toFixed(1) + '%）的门店：'
       + below.map(x => x.s + '(' + x.rate.toFixed(1) + '%,差' + (avgRate-x.rate).toFixed(1) + 'pct)').join('、')
       + '，需重点跟进目标达成。';
  return o || '各门店完成率较为均衡。';
}())
```

```javascript
// 示例：导购排名 vs 区域人均
conc('结论文本', function() {
  const all = [];
  STORE_ORDER.forEach(s => (DATA['_guides']||{})[s]?.forEach(g => all.push(g)));
  all.sort((a,b) => b.wk_sales - a.wk_sales);
  const regAvg = all.filter(g=>g.wk_sales>0).reduce((a,g)=>a+g.wk_sales,0)
                  / all.filter(g=>g.wk_sales>0).length;
  const below = all.filter(g => g.wk_sales > 0 && g.wk_sales < regAvg * 0.7);
  let o = '';
  if (below.length > 0)
    o += '⚠️ 低于区域均值¥' + Math.round(regAvg) + '的导购：'
       + below.map(g => g.name + '(' + g.store + ',¥' + Math.round(g.wk_sales) + ')').join('、')
       + '，需重点进行销售辅导和复盘。';
  return o || '导购团队整体平稳（区域人均¥' + Math.round(regAvg) + '）。';
}())
```

### 17.9 新增/修改机会点注意事项

1. **同步修改**：修改 `conc*()` 函数 → 同步修改 `gen_weekly_p2.py`（周）和 `gen_monthly_html.py`（月）
2. **同步修改**：修改 inline `conc()` → 同步修改门店页和区域对比页
3. **中文可落地**：机会点必须是中文、可落地的业务建议，避免空洞模板话术
4. **数据驱动**：机会点应基于实际数据阈值动态生成，不是固定文案
5. **向后兼容**：`o` 参数可选，不传时只显示结论区
6. **字段差异**：周分析用 `wk_*` 字段，月分析用 `month_*` 字段，机会点逻辑相同但字段名不同
7. **均值对比优先**：对比基准应优先用区域均值而非绝对阈值，标出具2-3条具体数字差距
8. **门店/区域双对比**：导购模块需同时计算门店人均和区域人均两个基准
9. **环比对比增强**：区域对比页的环比，加入与区域环比均值的比较

---

## 十八、成交率模块（月 `_conversion` + 周 `_weekly_conversion`）与区域对比布局重排

> 本工作流封装自实际改造：在周月看板加入「成交率（笔数/客流）」模块，区域对比页把「长尾款 + 成交率」缩成半宽并排、挪到第三排，门店页也加成交率卡片。月/周两侧口径一致、排版一致。

### 18.1 成交率口径

- **月成交率** = 月成交笔数 / 月客流量（客流来源：`7月客流（截止19日）.xlsx`）
- **周成交率** = 固定 **WK29 窗口（2026-07-13 ~ 2026-07-19）** 笔数 / WK29 客流（客流来源：`wk29客流.xlsx`）
- 笔数口径与日看板成交率一致（订单去重后的单号计数）
- 客流名→销售名映射 `FLOW_MAP`（扬州万象汇新店→扬州万象汇、淮安新亚广场→淮安新亚，其余同名）
- **统一 JSON 结构**：`{store: {flow, cnt, rate}, '_region': {flow, cnt, rate}}`，`rate = cnt/flow×100`（flow=0 时 rate=0）

### 18.2 数据层代码（extract_weekly_enhanced.py 4.9 块）

写在 4.8 月成交率块之后、`# 5. 更新 weekly_analysis.json` 之前：

```python
WK29_FLOW_FILE = '/Users/a123/Desktop/wk29客流.xlsx'
FLOW_MAP_WK = {
    '扬州万象汇新店': '扬州万象汇', '扬州京华城': '扬州京华城',
    '扬州江都金鹰': '扬州江都金鹰', '泰州万象城': '泰州万象城',
    '宿迁宝龙': '宿迁宝龙', '淮安新亚广场': '淮安新亚', '淮安万象城': '淮安万象城'
}
WK29_START = WK28_BASE + pd.Timedelta(days=7)   # 2026-07-13 (WK29 周一)
WK29_END   = WK28_BASE + pd.Timedelta(days=13)  # 2026-07-19 (WK29 周日)
wk29_flow_df = all_df[(all_df['销售日期_dt'] >= WK29_START) & (all_df['销售日期_dt'] <= WK29_END)]
weekly_conversion = {}
if os.path.exists(WK29_FLOW_FILE):
    wk_flow_raw = pd.read_excel(WK29_FLOW_FILE, sheet_name=0, header=0)
    wk_flow_raw['门店名'] = wk_flow_raw['门店名'].map(FLOW_MAP_WK)
    flow_map = dict(zip(wk_flow_raw['门店名'], wk_flow_raw['客流量']))
    for s in STORE_ORDER:
        fl = flow_map.get(s, 0)
        cnt = int(wk29_flow_df[wk29_flow_df['store'] == s].shape[0])
        rate = round(cnt / fl * 100, 2) if fl > 0 else 0
        weekly_conversion[s] = {'flow': int(fl), 'cnt': int(cnt), 'rate': rate}
rf = sum(weekly_conversion[s]['flow'] for s in STORE_ORDER)
rc = sum(weekly_conversion[s]['cnt'] for s in STORE_ORDER)
weekly_conversion['_region'] = {'flow': int(rf), 'cnt': int(rc), 'rate': round(rc / rf * 100, 2) if rf > 0 else 0}
print("   WK29门店周成交率: " + " | ".join(f"{s} {weekly_conversion[s]['rate']}%" for s in STORE_ORDER) + f" | 区域 {weekly_conversion['_region']['rate']}%")
```

并在写出 JSON 前挂载：

```python
weekly['_conversion'] = conversion            # 月成交率（4.8 块）
weekly['_weekly_conversion'] = weekly_conversion  # 周成交率（4.9 块）
```

> 月成交率 `_conversion` 由 `extract_monthly.py` 计算，结构完全相同，源文件为月客流 Excel。

### 18.3 渲染层 - 门店页成交率卡片

**周**（`gen_weekly_p3.py` `renderStore`）：

```javascript
const wcvo=(DATA['_weekly_conversion']||{});
const cvoRate=(wcvo[store]?wcvo[store].rate:0);
const cvoRegionRate=(wcvo._region?wcvo._region.rate:0);
// ...
'<div class="kpi-card"><div class="label">ZK29周成交率</div><div class="value '+(cvoRate>=cvoRegionRate?'green':'red')+'">'+cvoRate.toFixed(2)+'%</div><div class="sub">区域'+(cvoRegionRate?cvoRegionRate.toFixed(2):'N/A')+'% | '+(wcvo[store]?wcvo[store].cnt:'0')+'笔 / '+(wcvo[store]?wcvo[store].flow:'0')+'客流</div></div>'
```

末尾 replace 链：

```python
_js_out = JS_RENDER_STORE.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}').replace('ZK29','WK29')
```

**月**（`gen_monthly_p2.py` `renderStore`）：

```javascript
const cvo=(DATA['_conversion']||{});
const cvoRate=(cvo[store]?cvo[store].rate:0);
const cvoRegionRate=(cvo._region?cvo._region.rate:0);
// ...
'<div class="kpi-card"><div class="label">月成交率</div><div class="value '+(cvoRate>=cvoRegionRate?'green':'red')+'">'+cvoRate.toFixed(2)+'%</div><div class="sub">区域'+(cvoRegionRate?cvoRegionRate.toFixed(2):'N/A')+'% | '+(cvo[store]?cvo[store].cnt:'0')+'笔/'+ (cvo[store]?cvo[store].flow:'0') +'客流</div></div>'
```

> 月看板不做 WK 动态替换，无 ZK29 占位符需求。

### 18.4 渲染层 - 区域对比第三排（半宽并排）

**月**（`gen_monthly_p3.py` `renderComparison`），Row3 结构：

```javascript
// 门店月成交率 & 长尾款 对比（半宽并排，第三排）
const cvo=(DATA['_conversion']||{});
const convData=STORE_ORDER.map(s=>({s,v:cvo[s]?cvo[s].rate:0,flow:cvo[s]?cvo[s].flow:0,cnt:cvo[s]?cvo[s].cnt:0})).sort((a,b)=>b.v-a.v);
const rConv=cvo._region||{};
html+='<div class="chart-grid">'+
  '<div class="chart-card"><h3><span class="icon">📦</span> 各门店长尾款 & 样品销售月累计对比（标注占比）</h3><div class="chart-wrap"><canvas id="cmpLtSample"></canvas></div>'+
  conc('区域长尾款...', function(){...}()) +
  '<div class="chart-card"><h3><span class="icon">🎯</span> 各门店月成交率对比（成交笔数/客流量，本月累计）</h3><div class="chart-wrap"><canvas id="cmpConv"></canvas></div>'+
  conc('区域月成交率'+rConv.rate.toFixed(2)+'%（月成交'+rConv.cnt+'笔 ÷ 月客流'+rConv.flow+'人次）。各店：'+convData.map(x=>x.s+x.v.toFixed(2)+'%').join('、')+'。', function(){ /* 低于均值门店列出 + 最高最低差距 */ }()) +
  '</div></div>';
```

`cmpConv` 图表 init（月）：

```javascript
charts.cmpConv=new Chart(document.getElementById('cmpConv'),{
  type:'bar',
  data:{labels:convData.map(x=>x.s),datasets:[{label:'月成交率',data:convData.map(x=>x.v),backgroundColor:convData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:convData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6}]},
  options:{...tooltip:{callbacks:{label:ctx=>{const x=convData[ctx.dataIndex];return '月成交率: '+x.v.toFixed(2)+'% | 笔数'+x.cnt+' / 客流'+x.flow;}}}}
});
```

月汇总表追加「成交率」列：

```javascript
tHtml+='<thead><tr>...<th>成交率</th><th>WOS</th>...</tr></thead>';
// 门店行
'<td style="color:'+((cvo[s]&&cvo[s].rate>=rConv.rate)?'#16a34a':'#dc2626')+';font-weight:700">'+((cvo[s]&&cvo[s].rate)?cvo[s].rate.toFixed(2):'N/A')+'%</td>'
```

**周**（`gen_weekly_p4.py` `renderComparison`），Row3 结构与月一致，但用 `wcvo=(DATA['_weekly_conversion']||{})`，且 h3 / conc 文本 / 门店卡标签用 **ZK29** 占位（见 18.5）。周汇总表列名为「周成交率」。

### 18.5 ⚠️ 关键坑 1：WK 动态替换误伤成交率标签（ZK29 占位符）

`gen_weekly_p3/p4.py` 末尾有 `.replace('WK29', f'WK{WK_NUM_INT}')` 把正文周号动态化。当 `max_date` 推进到 WK30 时，若正文写死 `"WK29周成交率"`，会被误改成 `"WK30周成交率"`——但**数据是按固定 WK29 窗口算的**，标签与数据错位会误导用户。

**修复（三处一致）**：
1. 成交率相关字面量用 `ZK29` 占位：
   - 区域对比 h3：`各门店ZK29周成交率对比（ZK29笔数/客流量）`
   - conc 文本：`区域ZK29周成交率` + `ZK29成交` + `ZK29客流`
   - 门店卡标签：`ZK29周成交率`
2. 末尾 replace 链末尾追加 `.replace('ZK29','WK29')` 还原（必须放在最后，避免被前面的 WK29→当前周 替换再次命中）：
   ```python
   _html_out = HTML.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}').replace('ZK29','WK29')
   ```
3. 月看板（`gen_monthly_p2/p3.py`）不做 WK 替换，**无此坑**，直接写 `月成交率`。

> 验证：周看板应显示「WK29周成交率」，全文不得残留 `ZK29`。

### 18.6 ⚠️ 关键坑 2：gen_weekly_p4.py 中文是 \uXXXX 转义

`gen_weekly_p4.py` 用 raw string 存中文（如 `各门店` → `\u5404\u95e8\u5e97`），**Edit 工具无法匹配中文**。
- 改 **JS 正文**（如 h3 文案、cmpConv init）→ 必须用 **Python 脚本**读取→字符串替换→写回
- 改 **Python 部分**（如末尾 replace 链、`WK_NUM_INT` 计算）→ 可直接 Edit
- 月看板 `gen_monthly_p2/p3.py` 用真实中文，Edit 可正常匹配

### 18.7 修复损坏渲染文件的锚点清理法

若重排脚本把 `gen_monthly_p3.py` / `gen_weekly_p4.py` 改坏（典型症状：半宽块 `'<div class="chart-grid">' + <div...` 语法错误、重复成交率 full 块、孤立的 `cmpLtSample` 图表块被误删只剩 chart init），按以下步骤修复：

1. 用**唯一注释锚点**切分文本，例如：
   - GARBAGE = `// 门店月成交率 & 长尾款 对比（半宽并排，第三排）`
   - WOS = `// WOS & JL comparison + Series proportion`
   - DUP = `// 门店月成交率对比（成交率 = 成交笔数 / 客流量）`
   - TOP10 = `// 区域TOP10产品`
2. 删除重复块、重建 Row3（长尾款 `cmpLtSample` + 成交率 `cmpConv` 半宽并排）
3. **断言唯一性**后写回（防止再次损坏）：
   - `cmpLtSample` canvas 出现 **1** 次
   - `cmpConv` canvas 出现 **1** 次
   - `const cvo` / `const wcvo` 声明各 **1** 次
4. 跑 `auto_run_weekly.py`（含 node JS 校验），确保无语法错误再部署

> 不要用零散的 `sed`/手工拼接乱码段；按锚点精准清理最稳。

### 18.8 验证清单

- [ ] 周/月成交率数据正确（区域值、各店值、笔数/客流对得上客流文件）
- [ ] 周看板显示「WK29周成交率」、无 `ZK29` 残留
- [ ] 区域对比第三排 = 长尾款 + 成交率 半宽并排
- [ ] 门店页有成交率卡片（周 ZK29 / 月 月成交率）
- [ ] node JS 校验通过，流水线重跑成功
- [ ] 部署 outputs_analysis，present_files 展示固定链接
