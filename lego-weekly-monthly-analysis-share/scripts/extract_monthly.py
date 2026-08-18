#!/usr/bin/env python3
"""
月度交互分析数据提取脚本
- 月累计 / 月环比(上月同时间区间) / 月同比(去年同月)
- 月导购(含月指标完成率) / 月产品名称TOP10 / 月产品系列
- 排除4类人员
输出到 monthly_analysis.json
"""
import pandas as pd
import json
import os
import glob
from datetime import datetime

import wos_attr  # WOS 导购归因（含订单备注 [SA_NAME] 补充归因）

# ============================================================
# 配置
# ============================================================
STORE_ORDER = ['扬州万象汇', '扬州京华城', '扬州江都金鹰', '泰州万象城', '宿迁宝龙', '淮安新亚', '淮安万象城']
STORE_CODE_MAP = {
    'LCS133-LW': '扬州万象汇', 'LCS275-LW': '扬州京华城', 'LCS364-LW': '扬州江都金鹰',
    'LCS0632-LW': '泰州万象城', 'LCS249-LW': '宿迁宝龙', 'LCS305-LW': '淮安新亚',
    'LCS0420-LW': '淮安万象城'
}
NEW_STORE = '淮安万象城'
WK28_BASE = pd.Timestamp('2026-07-06')

# 排除4类人员（销售数据完全排除，不参与任何统计）
EXCLUDE_NAMES = {'李婷', '黄小莉', '扬州金鹰-长期兼职', '泰州万象城-长期兼职'}

# File paths — 动态查找，适配任意月份
csvs = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'), key=os.path.getmtime)
TODAY_XLSX = desktop_xlsxs[-1] if desktop_xlsxs else ''
TARGET_XLSX = sorted(glob.glob('/Users/a123/Desktop/。/**/*月指标同期.xlsx', recursive=True), key=os.path.getmtime)
TARGET_XLSX = TARGET_XLSX[-1] if TARGET_XLSX else ''
_cands = []
_cands += glob.glob('/Users/a123/Desktop/。/**/苏中区域*月员工指标分解.xlsx', recursive=True)
_cands += glob.glob('/Users/a123/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/*/Caches/Files/2026-*/*/苏中区域*月员工指标分解.xlsx')
GUIDE_XLSX = sorted(_cands, key=os.path.getmtime)[-1] if _cands else ''
LT_SKU_XLSX = sorted(glob.glob('/Users/a123/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/*/Caches/Files/2026-*/*/长尾款sku明细.xlsx'), key=os.path.getmtime)
LT_SKU_XLSX = LT_SKU_XLSX[-1] if LT_SKU_XLSX else ''

OUTPUT_JSON = '/Users/a123/WorkBuddy/Claw/outputs/monthly_analysis.json'

print(f"CSV文件: {len(csvs)}个")
print(f"今日文件: {os.path.basename(TODAY_XLSX) if TODAY_XLSX else '无(仅用CSV)'}")

# ============================================================
# 1. 读取+清洗订单数据（包含近3个月所有数据，动态适配）
# ============================================================
print("\n1. 读取订单数据...")
hist_dfs = [pd.read_csv(f, low_memory=False) for f in csvs]
hist_df = pd.concat(hist_dfs, ignore_index=True)
if TODAY_XLSX and os.path.exists(TODAY_XLSX):
    today_df = pd.read_excel(TODAY_XLSX)
    all_df = pd.concat([hist_df, today_df], ignore_index=True)
else:
    all_df = hist_df

# 先统一解析日期（混合格式容错：CSV 2026-07-20 13:22:00 / xlsx 2026/7/20 13:22）
all_df['销售日期_dt'] = pd.to_datetime(all_df['销售日期'], format='mixed', errors='coerce')

# 去重
all_df['_dedup_key'] = (
    all_df['单号'].astype(str) + '|' +
    all_df['乐高编号'].apply(lambda x: str(int(float(x))) if pd.notna(x) and x != '' else '__nan__') + '|' +
    all_df['商品名称'].fillna('').astype(str) + '|' +
    all_df['数量'].astype(float).round(2).astype(str) + '|' +
        all_df['结算金额'].astype(float).round(2).astype(str) + '|' +
        all_df['导购姓名'].fillna('').astype(str)
    )
all_df = all_df.drop_duplicates(subset=['_dedup_key'])
all_df = all_df.drop(columns=['_dedup_key'])

# 排除4类人员
before_exclude = len(all_df)
all_df = all_df[~all_df['导购姓名'].astype(str).str.strip().isin(EXCLUDE_NAMES)]
excluded_count = before_exclude - len(all_df)
print(f"   合并去重后: {before_exclude}行, 排除4类人员: {excluded_count}行, 剩余: {len(all_df)}行")

# 日期处理
all_df['销售日期_dt'] = pd.to_datetime(all_df['销售日期'], format='mixed', errors='coerce')
# 动态检测报告月份，保留近3个月数据
_max_date_raw = all_df['销售日期_dt'].max()
REPORT_MONTH = _max_date_raw.month
_date_3m_ago = _max_date_raw - pd.DateOffset(months=3)
all_df = all_df[all_df['销售日期_dt'] >= _date_3m_ago]

# 保存样品仓订单（用于样品销售模块），后续从主数据排除
sample_df = all_df[all_df['下单店仓名称'].str.contains('样品', na=False)].copy()
sample_df = sample_df[sample_df['是否赠品'] != '是']
sample_df = sample_df[sample_df['结算金额'].astype(float) != 0]

# 排除赠品、样品、零金额
all_df = all_df[all_df['是否赠品'] != '是']
all_df = all_df[~all_df['下单店仓名称'].str.contains('样品', na=False)]
all_df = all_df[all_df['结算金额'].astype(float) != 0]

# 门店映射
all_df['门店名'] = all_df['发货店仓编码'].map(STORE_CODE_MAP)
all_df = all_df[all_df['门店名'].notna()]
all_df['结算金额'] = all_df['结算金额'].astype(float)
# 上市日期解析（新品口径：年/月 == 报告年/月）
all_df['上市日期_dt'] = pd.to_datetime(all_df['上市日期'], errors='coerce')

# 样品数据也做同样的门店映射
sample_df['门店名'] = sample_df['发货店仓编码'].map(STORE_CODE_MAP)
sample_df = sample_df[sample_df['门店名'].notna()]
sample_df['结算金额'] = sample_df['结算金额'].astype(float)
sample_df['牌价额'] = sample_df['牌价额'].astype(float)

# 乐高编号统一
all_df['乐高编号_str'] = all_df['乐高编号'].apply(
    lambda v: str(int(v)) if pd.notna(v) and v == int(v) else str(v) if pd.notna(v) else ''
)

# 确定报告日期（从全量数据的最大日期推算）
max_date = all_df[all_df['销售日期_dt'].dt.month == REPORT_MONTH]['销售日期_dt'].max()
max_date_eod = pd.Timestamp(max_date.year, max_date.month, max_date.day, 23, 59, 59)
report_day = max_date.day
report_day_index = report_day + 1  # 月1日=col2, 月N日=col(N+1)
DAYS_IN_MONTH = max_date.days_in_month
print(f"   报告日期: {REPORT_MONTH}月{report_day}日 (col{report_day_index}), 当月{DAYS_IN_MONTH}天")
print(f"   数据范围: {all_df['销售日期_dt'].min().strftime('%Y-%m-%d')} ~ {all_df['销售日期_dt'].max().strftime('%Y-%m-%d')}")

# 分离当月和上月数据
cur_month_df = all_df[all_df['销售日期_dt'].dt.month == REPORT_MONTH]
# 上月：按日期范围取（处理跨年）
_month_start = pd.Timestamp(max_date.year, max_date.month, 1)
_prev_month_start = _month_start - pd.DateOffset(months=1)
prev_month_df = all_df[(all_df['销售日期_dt'] >= _prev_month_start) & (all_df['销售日期_dt'] < _month_start)]
# 上月同时间区间: prev 1st - prev (report_day)
jun_same_period = prev_month_df[prev_month_df['销售日期_dt'].dt.day <= report_day]

print(f"   当月({REPORT_MONTH}月)数据: {len(cur_month_df)}行, 上月数据: {len(prev_month_df)}行, 上月同期(1-{report_day}日): {len(jun_same_period)}行")

# Read LT SKUs
if LT_SKU_XLSX:
    lt_sku_df = pd.read_excel(LT_SKU_XLSX)
    lt_skus = set(lt_sku_df['乐高编号'].apply(lambda x: str(int(x)) if pd.notna(x) else '').tolist())
    # 提取近3个月新品（当月新品）
    sku_dates = pd.to_datetime(lt_sku_df['上市时间'], errors='coerce')
    recent_cutoff = pd.Timestamp('2026-05-01')
    new_products = lt_sku_df[sku_dates >= recent_cutoff]['商品名称'].dropna().tolist()
    print(f"   近3月新品: {len(new_products)}款")
else:
    lt_skus = set()
    new_products = []

# ============================================================
# 2. 读取指标同期Excel
# ============================================================
print("\n2. 读取指标同期数据...")
target_df = pd.read_excel(TARGET_XLSX, header=None)

# 门店月指标 (col1, rows 2-9)
store_month_target = {}
for i in range(2, 9):
    name = target_df.iloc[i, 0]
    target = float(target_df.iloc[i, 1]) if pd.notna(target_df.iloc[i, 1]) else 0
    store_month_target[name] = target
store_month_target['苏中区域'] = float(target_df.iloc[9, 1]) if pd.notna(target_df.iloc[9, 1]) else 0

# 同期数据 (rows 13-20, col2-col(report_day_index))
# 保留每日值 + 月累计
yoy_data = {}
yoy_daily_data = {}  # 每日同期数据
for i in range(13, 21):
    name = target_df.iloc[i, 0]
    if pd.isna(name):
        continue
    daily_vals = []
    for j in range(2, min(report_day_index + 1, target_df.shape[1])):
        v = target_df.iloc[i, j]
        if pd.notna(v):
            daily_vals.append(float(v))
        else:
            daily_vals.append(0)
    yoy_daily_data[str(name).strip()] = daily_vals
    yoy_data[name] = sum(daily_vals) if daily_vals else None

print(f"   月指标: {store_month_target}")
print(f"   同期月累计(去年): {yoy_data}")

# ============================================================
# 3. 月度门店KPI
# ============================================================
print("\n3. 计算月度门店KPI...")
store_data = {}
new_detail = {}  # 各店新品TOP明细 {store: [{name,amt,qty}, ...]}

for s in STORE_ORDER:
    s_jul = cur_month_df[cur_month_df['门店名'] == s]
    s_jun_sp = jun_same_period[jun_same_period['门店名'] == s]
    
    # 月累计
    month_amt = round(s_jul['结算金额'].sum(), 2)
    
    # 月环比 (上月同时间区间)
    jun_sp_amt = round(s_jun_sp['结算金额'].sum(), 2)
    mom = round((month_amt - jun_sp_amt) / jun_sp_amt * 100, 1) if jun_sp_amt > 0 else None
    
    # 月同比 (去年同月同期)
    yoy_total = yoy_data.get(s)
    if yoy_total and yoy_total > 0:
        yoy_pct = round((month_amt - yoy_total) / yoy_total * 100, 1)
    else:
        yoy_pct = None  # 淮安万象城无同期
    
    # 月指标 & 完成率
    target = store_month_target.get(s, 0)
    month_rate = round(month_amt / target * 100, 1) if target > 0 else 0
    
    # 客单价、连带率、笔数
    pos = s_jul[s_jul['结算金额'] > 0]
    month_cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    month_qty = int(pos['数量'].sum()) if len(pos) > 0 else 0
    month_atv = round(month_amt / month_cnt, 2) if month_cnt > 0 else 0
    month_jd = round(month_qty / month_cnt, 2) if month_cnt > 0 else 0

    # 新品（上市日期年/月 == 报告年/月）
    new_mask = (s_jul['上市日期_dt'].dt.year == max_date.year) & (s_jul['上市日期_dt'].dt.month == REPORT_MONTH)
    s_new = s_jul[new_mask]
    new_amt = round(s_new['结算金额'].sum(), 2)
    s_new_pos = s_new[s_new['结算金额'] > 0]
    new_qty = int(s_new_pos['数量'].sum()) if len(s_new_pos) > 0 else 0
    new_cnt = s_new_pos['单号'].nunique() if len(s_new_pos) > 0 else 0
    new_rate = round(new_amt / month_amt * 100, 1) if month_amt > 0 else 0
    # 新品TOP（按商品名称）
    s_new_list = []
    if len(s_new_pos) > 0:
        np_amt = s_new_pos.groupby('商品名称')['结算金额'].sum()
        np_qty = s_new_pos.groupby('商品名称')['数量'].sum()
        for nm in np_amt.index:
            s_new_list.append({'name': nm, 'amt': round(float(np_amt[nm]), 2), 'qty': int(np_qty[nm])})
        s_new_list.sort(key=lambda x: x['amt'], reverse=True)
        s_new_list = s_new_list[:10]
    new_detail[s] = s_new_list
    
    # WOS月累计
    month_wos = round(s_jul[s_jul['支付方式名称'].str.contains('WOS', na=False)]['结算金额'].sum(), 2)
    
    # 即时零售月累计
    jl_mask = (s_jul['支付方式名称'].str.contains('美团', na=False) |
               s_jul['特殊销售类型'].str.contains('饿了么', na=False) |
               s_jul['支付方式名称'].str.contains('淘宝', na=False) |
               s_jul['支付方式名称'].str.contains('京东', na=False))
    month_jl = round(s_jul[jl_mask]['结算金额'].sum(), 2)
    
    # 长尾款月累计
    s_lt = s_jul[s_jul['乐高编号_str'].isin(lt_skus)]
    month_lt = round(s_lt['结算金额'].sum(), 2)
    month_lt_qty = int(s_lt[s_lt['结算金额'] > 0]['数量'].sum()) if len(s_lt[s_lt['结算金额'] > 0]) > 0 else 0
    
    # 日销售趋势 — 对象数组（含同比、完成率、长尾款）
    s_yoy_daily = yoy_daily_data.get(s, [])
    daily_target = target / DAYS_IN_MONTH if target > 0 else 0
    daily = []
    daily_prev = []
    for day in range(1, report_day + 1):
        day_data = s_jul[s_jul['销售日期_dt'].dt.day == day]
        day_amt = day_data['结算金额'].sum()
        day_pos = day_data[day_data['结算金额'] > 0]
        day_lt = day_data[day_data['乐高编号_str'].isin(lt_skus)]
        day_lt_amt = day_lt['结算金额'].sum()
        day_lt_qty = int(day_lt[day_lt['结算金额'] > 0]['数量'].sum()) if len(day_lt[day_lt['结算金额'] > 0]) > 0 else 0
        day_wos = day_data[day_data['支付方式名称'].str.contains('WOS', na=False)]['结算金额'].sum()
        day_jl_mask = (day_data['支付方式名称'].str.contains('美团', na=False) |
                       day_data['特殊销售类型'].str.contains('饿了么', na=False) |
                       day_data['支付方式名称'].str.contains('淘宝', na=False) |
                       day_data['支付方式名称'].str.contains('京东', na=False))
        day_jl = day_data[day_jl_mask]['结算金额'].sum()
        yoy_amt = s_yoy_daily[day - 1] if day <= len(s_yoy_daily) else 0
        day_yoy_pct = round((day_amt - yoy_amt) / yoy_amt * 100, 1) if yoy_amt > 0 else None
        rate = round(day_amt / daily_target * 100, 1) if daily_target > 0 else 0
        daily.append({
            'day': day,
            'amt': round(day_amt, 2),
            'target': round(daily_target, 2),
            'rate': rate,
            'yoy_amt': round(yoy_amt, 2),
            'yoy_pct': day_yoy_pct,
            'lt': round(day_lt_amt, 2),
            'lt_qty': day_lt_qty,
            'wos': round(day_wos, 2),
            'jl': round(day_jl, 2),
            'tx_cnt': day_pos['单号'].nunique() if len(day_pos) > 0 else 0,
        })
        # 上月同日
        day_jun = s_jun_sp[s_jun_sp['销售日期_dt'].dt.day == day]
        daily_prev.append(round(day_jun['结算金额'].sum(), 2))
    
    store_data[s] = {
        'month_amt': month_amt,
        'mom': mom,
        'mom_prev': jun_sp_amt,
        'yoy': yoy_pct,
        'yoy_prev': yoy_total,
        'target': target,
        'month_rate': month_rate,
        'month_cnt': month_cnt,
        'month_qty': month_qty,
        'month_atv': month_atv,
        'month_jd': month_jd,
        'month_wos': month_wos,
        'new_amt': new_amt,
        'new_qty': new_qty,
        'new_cnt': new_cnt,
        'new_rate': new_rate,
        'month_jl': month_jl,
        'month_lt': month_lt,
        'month_lt_qty': month_lt_qty,
        'daily': daily,
        'daily_prev': daily_prev,
    }
    yoy_str = f"{yoy_pct:+.1f}%" if yoy_pct is not None else 'N/A'
    mom_str = f"{mom:+.1f}%" if mom is not None else 'N/A'
    print(f"   {s}: 月累计=¥{month_amt:,.0f} | 环比{mom_str} | 同比{yoy_str} | 完成率{month_rate}%")

# 区域合计
region_data = {
    'month_amt': sum(s['month_amt'] for s in store_data.values()),
    'mom_prev': sum(s['mom_prev'] for s in store_data.values()),
    'yoy_prev': yoy_data.get('苏中区域'),
    'target': store_month_target.get('苏中区域', 0),
    'month_cnt': sum(s['month_cnt'] for s in store_data.values()),
    'month_qty': sum(s['month_qty'] for s in store_data.values()),
    'month_wos': sum(s['month_wos'] for s in store_data.values()),
    'month_jl': sum(s['month_jl'] for s in store_data.values()),
    'month_lt': sum(s['month_lt'] for s in store_data.values()),
}
# 月环比：7店 vs 7店上月同期（淮安万象城上月有数据）
region_data['mom'] = round((region_data['month_amt'] - region_data['mom_prev']) / region_data['mom_prev'] * 100, 1) if region_data['mom_prev'] > 0 else None
# 月同比：排除淮安万象城（无同期数据），用6店月累计 ÷ 6店同期
region_month_no_new = sum(store_data[s]['month_amt'] for s in STORE_ORDER if s != NEW_STORE)
ry = region_data['yoy_prev']
region_data['yoy'] = round((region_month_no_new - ry) / ry * 100, 1) if ry and ry > 0 else None
region_data['month_rate'] = round(region_data['month_amt'] / region_data['target'] * 100, 1) if region_data['target'] > 0 else 0
region_data['month_atv'] = round(region_data['month_amt'] / region_data['month_cnt'], 2) if region_data['month_cnt'] > 0 else 0
region_data['month_jd'] = round(region_data['month_qty'] / region_data['month_cnt'], 2) if region_data['month_cnt'] > 0 else 0

# 新品区域汇总
region_new_mask = (cur_month_df['上市日期_dt'].dt.year == max_date.year) & (cur_month_df['上市日期_dt'].dt.month == REPORT_MONTH)
region_new_df = cur_month_df[region_new_mask]
region_new_amt = round(region_new_df['结算金额'].sum(), 2)
region_new_pos = region_new_df[region_new_df['结算金额'] > 0]
region_new_qty = int(region_new_pos['数量'].sum()) if len(region_new_pos) > 0 else 0
region_new_cnt = region_new_pos['单号'].nunique() if len(region_new_pos) > 0 else 0
region_new_rate = round(region_new_amt / region_data['month_amt'] * 100, 1) if region_data['month_amt'] > 0 else 0
region_data['new_amt'] = region_new_amt
region_data['new_qty'] = region_new_qty
region_data['new_cnt'] = region_new_cnt
region_data['new_rate'] = region_new_rate
# 区域新品TOP
region_new_list = []
if len(region_new_pos) > 0:
    rnp_amt = region_new_pos.groupby('商品名称')['结算金额'].sum()
    rnp_qty = region_new_pos.groupby('商品名称')['数量'].sum()
    for nm in rnp_amt.index:
        region_new_list.append({'name': nm, 'amt': round(float(rnp_amt[nm]), 2), 'qty': int(rnp_qty[nm])})
    region_new_list.sort(key=lambda x: x['amt'], reverse=True)
    region_new_list = region_new_list[:10]
new_products_detail = dict(new_detail)
new_products_detail['_region'] = region_new_list
print(f"   区域新品(上市日期={REPORT_MONTH}月): ¥{region_new_amt:,.0f} | 占比{region_new_rate}% | {region_new_qty}件 | TOP{len(region_new_list)}款")

# ============================================================
# 3.5 门店月成交率（成交率 = 月成交笔数 / 月客流）
# ============================================================
print("\n3.5 计算门店月成交率（笔数/客流）...")
FLOW_FILE = '/Users/a123/Desktop/8月客流（截止2号）.xlsx'
FLOW_MAP = {
    '扬州万象汇新店': '扬州万象汇', '扬州京华城': '扬州京华城', '扬州江都金鹰': '扬州江都金鹰',
    '泰州万象城': '泰州万象城', '宿迁宝龙': '宿迁宝龙', '淮安新亚广场': '淮安新亚',
    '淮安万象城': '淮安万象城'
}
try:
    # 8月客流(截止2号)文件：两列无表头 schema=门店名/客流值，已覆盖到报告日(8/1-8/2)，无需并入周客流
    flow_raw = pd.read_excel(FLOW_FILE, sheet_name=0, header=None)
    flow_raw = flow_raw.rename(columns={0: 'flow_store', 1: '客流'})
    flow_raw['门店名'] = flow_raw['flow_store'].map(FLOW_MAP)
    flow_map = flow_raw.dropna(subset=['门店名']).set_index('门店名')['客流'].to_dict()
    print(f"   客流文件读取(8月截止2号): 苏中7店客流合计 {int(sum(flow_map.values()))} 人次")
except Exception as e:
    print(f"   ⚠️ 客流文件读取失败: {e}")
    flow_map = {}
print(f"   月客流合计: {int(sum(flow_map.values()))} 人次")
conversion = {}
for s in STORE_ORDER:
    fl = flow_map.get(s, 0)
    cnt = store_data[s]['month_cnt']
    rate = round(cnt / fl * 100, 2) if fl > 0 else 0
    conversion[s] = {'flow': int(fl), 'cnt': int(cnt), 'rate': rate}
rf = sum(conversion[s]['flow'] for s in STORE_ORDER)
rc = region_data['month_cnt']
conversion['_region'] = {'flow': int(rf), 'cnt': int(rc), 'rate': round(rc / rf * 100, 2) if rf > 0 else 0}
print("   门店月成交率: " + " | ".join(f"{s} {conversion[s]['rate']}%" for s in STORE_ORDER) + f" | 区域 {conversion['_region']['rate']}%")

# ============================================================
# 4. 产品系列结构 (产品大类)
# ============================================================
print("\n4. 计算月度产品系列结构...")
series_by_store = {}
for s in STORE_ORDER:
    s_jul_pos = cur_month_df[(cur_month_df['门店名'] == s) & (cur_month_df['结算金额'] > 0)]
    series_amt = s_jul_pos.groupby('产品大类')['结算金额'].sum().round(2).to_dict()
    series_amt = dict(sorted(series_amt.items(), key=lambda x: x[1], reverse=True))
    series_by_store[s] = series_amt

r_jul_pos = cur_month_df[cur_month_df['结算金额'] > 0]
region_series = r_jul_pos.groupby('产品大类')['结算金额'].sum().round(2).to_dict()
region_series = dict(sorted(region_series.items(), key=lambda x: x[1], reverse=True))
print(f"   区域: {len(region_series)}个系列, TOP3: {list(region_series.items())[:3]}")

# ============================================================
# 5. 产品名称TOP10 (商品名称，含件数)
# ============================================================
print("\n5. 计算月度产品名称TOP10(含件数)...")
products_by_store = {}
for s in STORE_ORDER:
    s_jul_pos = cur_month_df[(cur_month_df['门店名'] == s) & (cur_month_df['结算金额'] > 0)]
    prod_amt = s_jul_pos.groupby('商品名称')['结算金额'].sum()
    prod_qty = s_jul_pos.groupby('商品名称')['数量'].sum()
    prod_dict = {}
    for name in prod_amt.index:
        prod_dict[name] = {'amt': round(float(prod_amt[name]), 2), 'qty': int(prod_qty[name])}
    prod_dict = dict(sorted(prod_dict.items(), key=lambda x: x[1]['amt'], reverse=True))
    products_by_store[s] = prod_dict

r_jul_pos = cur_month_df[cur_month_df['结算金额'] > 0]
r_prod_amt = r_jul_pos.groupby('商品名称')['结算金额'].sum()
r_prod_qty = r_jul_pos.groupby('商品名称')['数量'].sum()
region_products = {}
for name in r_prod_amt.index:
    region_products[name] = {'amt': round(float(r_prod_amt[name]), 2), 'qty': int(r_prod_qty[name])}
region_products = dict(sorted(region_products.items(), key=lambda x: x[1]['amt'], reverse=True))
products_by_store['_region'] = region_products
print(f"   区域: {len(region_products)}个商品, TOP3: {[(k, v['amt']) for k, v in list(region_products.items())[:3]]}")

# ============================================================
# 6. 导购月度数据（含月指标完成率）
# ============================================================
print("\n6. 计算导购月度数据...")

def normalize_guide_name(name):
    if name is None or str(name).strip() == '':
        return None
    name = str(name).strip()
    if name.startswith('长期兼职-'):
        name = name[len('长期兼职-'):]
    if name == '扬万支援销售':
        name = '李婷1'
    if name == '泰州万象城-支援2' or name == '泰州万象城支援2':
        name = '黄小莉1'
    return name

def normalize_store_name(name):
    if name is None:
        return None
    name = str(name).strip()
    variants = {
        '扬州万象汇店': '扬州万象汇', '扬州京华城店': '扬州京华城',
        '扬州江都金鹰店': '扬州江都金鹰', '泰州万象城': '泰州万象城',
        '宿迁宝龙店': '宿迁宝龙', '淮安新亚广场店': '淮安新亚',
        '淮安万象城店': '淮安万象城',
    }
    return variants.get(name, name)

guide_df = pd.read_excel(GUIDE_XLSX)
guide_info = {}
for _, row in guide_df.iterrows():
    gname = normalize_guide_name(row.get('姓名', row.get('导购姓名', '')))
    if gname is None:
        continue
    store = normalize_store_name(row.get('门店', row.get('门店名称', '')))
    if store not in STORE_ORDER:
        continue
    title = str(row.get('职务', row.get('職务', ''))).strip()
    target = float(row.get('员工个人指标', row.get('个人指标', 0)))
    if gname in guide_info:
        guide_info[gname]['target'] += target
    else:
        guide_info[gname] = {'store': store, 'title': title, 'target': target}

# 李婷1、黄小莉1 是扬万支援销售/泰州支援2重命名后的独立导购
# 导购指标表中只有李婷/黄小莉，需为李婷1/黄小莉1复制指标
if '李婷' in guide_info and '李婷1' not in guide_info:
    guide_info['李婷1'] = dict(guide_info['李婷'])
if '黄小莉' in guide_info and '黄小莉1' not in guide_info:
    guide_info['黄小莉1'] = dict(guide_info['黄小莉'])

# 去掉李婷、黄小莉本体（保留李婷1、黄小莉1）
guide_info.pop('李婷', None)
guide_info.pop('黄小莉', None)

def get_match_names(gname):
    match_names = [gname, f'长期兼职-{gname}']
    if gname == '李婷1':
        match_names += ['扬万支援销售']
    if gname == '黄小莉1':
        match_names += ['泰州万象城-支援2', '泰州万象城支援2']
    return match_names

def calc_guide_monthly(m_df, gname, info, wos_map):
    store = info['store']
    match_names = get_match_names(gname)
    g_m = m_df[(m_df['门店名'] == store) & (m_df['导购姓名'].astype(str).str.strip().isin(match_names))]
    if len(g_m) == 0:
        return {'m_sales': 0, 'm_qty': 0, 'm_cnt': 0, 'm_atv': 0, 'm_jd': 0, 'm_wos': 0, 'm_lt': 0, 'm_lt_qty': 0, 'm_np': 0}
    pos = g_m[g_m['结算金额'] > 0]
    m_sales = g_m['结算金额'].sum()
    m_qty = int(pos['数量'].sum()) if len(pos) > 0 else 0
    m_cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    m_atv = round(m_sales / m_cnt, 2) if m_cnt > 0 else 0
    m_jd = round(m_qty / m_cnt, 2) if m_cnt > 0 else 0
    # WOS（备注补充归因：导购姓名空 → 取订单备注 [SA_NAME]）
    m_wos = wos_attr.wos_lookup(wos_map, store, gname)
    g_lt = g_m[g_m['乐高编号_str'].isin(lt_skus)]
    m_lt = g_lt['结算金额'].sum()
    m_lt_qty = int(g_lt[g_lt['结算金额'] > 0]['数量'].sum()) if len(g_lt[g_lt['结算金额'] > 0]) > 0 else 0
    np_mask = (g_m['上市日期_dt'].dt.year == max_date.year) & (g_m['上市日期_dt'].dt.month == REPORT_MONTH)
    m_np = g_m[np_mask]['结算金额'].sum()
    return {
        'm_sales': round(m_sales, 2), 'm_qty': m_qty, 'm_cnt': m_cnt,
        'm_atv': m_atv, 'm_jd': m_jd, 'm_wos': round(m_wos, 2),
        'm_lt': round(m_lt, 2), 'm_lt_qty': m_lt_qty, 'm_np': round(m_np, 2),
    }

def calc_guide_monthly_prev(m_df, gname, info):
    """上月同时间区间"""
    store = info['store']
    match_names = get_match_names(gname)
    g_m = m_df[(m_df['门店名'] == store) & (m_df['导购姓名'].astype(str).str.strip().isin(match_names))]
    return round(g_m['结算金额'].sum(), 2) if len(g_m) > 0 else 0

guide_monthly = {}

# WOS 导购归因（含订单备注 [SA_NAME] 补充归因）—— 与导购WOS重算报表口径一致
cur_month_wos_map = wos_attr.build_wos_map(
    cur_month_df, store_col='门店名', guide_col='导购姓名', note_col='备注',
    amt_col='结算金额', wos_col='支付方式名称', normalizer=normalize_guide_name)

for gname, info in guide_info.items():
    m_data = calc_guide_monthly(cur_month_df, gname, info, cur_month_wos_map)
    m_prev = calc_guide_monthly_prev(jun_same_period, gname, info)
    target = info['target']
    m_rate = round(m_data['m_sales'] / target * 100, 1) if target > 0 else 0
    mom = round((m_data['m_sales'] - m_prev) / m_prev * 100, 1) if m_prev > 0 else None
    
    guide_monthly[gname] = {
        'name': gname, 'store': info['store'], 'title': info['title'],
        'target': target,
        'm_sales': m_data['m_sales'], 'm_qty': m_data['m_qty'],
        'm_cnt': m_data['m_cnt'], 'm_atv': m_data['m_atv'],
        'm_jd': m_data['m_jd'], 'm_wos': m_data['m_wos'],
        'm_lt': m_data['m_lt'], 'm_lt_qty': m_data['m_lt_qty'],
        'm_np': m_data['m_np'],
        'm_rate': m_rate, 'm_prev': m_prev, 'mom': mom,
    }
    mom_str = f"{mom:+.1f}%" if mom is not None else 'N/A'
    print(f"   {gname} ({info['store']}): 月=¥{m_data['m_sales']:,.0f} | 上月同期=¥{m_prev:,.0f} | 环比{mom_str} | 完成率{m_rate}%")

# 按门店分组
guide_by_store = {}
for gname, data in guide_monthly.items():
    store = data['store']
    if store not in guide_by_store:
        guide_by_store[store] = []
    guide_by_store[store].append(data)
for store in guide_by_store:
    guide_by_store[store].sort(key=lambda x: x['m_sales'], reverse=True)

series_by_store['_region'] = region_series

# 构建区域daily数组
r_yoy_daily = yoy_daily_data.get('苏中区域', [])
r_daily_target = region_data['target'] / 31 if region_data['target'] > 0 else 0
region_daily = []
for day in range(1, report_day + 1):
    day_jul = cur_month_df[cur_month_df['销售日期_dt'].dt.day == day]
    day_amt = day_jul['结算金额'].sum()
    # 同比排除淮安万象城（无同期数据）
    day_jul_no_new = day_jul[day_jul['门店名'] != NEW_STORE]
    day_amt_no_new = day_jul_no_new['结算金额'].sum()
    day_lt = day_jul[day_jul['乐高编号_str'].isin(lt_skus)]
    day_lt_amt = day_lt['结算金额'].sum()
    day_lt_qty = int(day_lt[day_lt['结算金额'] > 0]['数量'].sum()) if len(day_lt[day_lt['结算金额'] > 0]) > 0 else 0
    yoy_amt = r_yoy_daily[day - 1] if day <= len(r_yoy_daily) else 0
    yoy_pct = round((day_amt_no_new - yoy_amt) / yoy_amt * 100, 1) if yoy_amt > 0 else None
    rate = round(day_amt / r_daily_target * 100, 1) if r_daily_target > 0 else 0
    region_daily.append({
        'day': day,
        'amt': round(day_amt, 2),
        'target': round(r_daily_target, 2),
        'rate': rate,
        'yoy_amt': round(yoy_amt, 2),
        'yoy_pct': yoy_pct,
        'lt': round(day_lt_amt, 2),
        'lt_qty': day_lt_qty,
    })
region_data['daily'] = region_daily

# ============================================================
# 6.5 样品销售数据
# ============================================================
print("\n6.5 计算样品销售数据...")
month_start = pd.Timestamp(max_date.year, max_date.month, 1)
sample_data = {}
for store in STORE_ORDER:
    s_month = sample_df[(sample_df['门店名'] == store) & (sample_df['销售日期_dt'] >= month_start) & (sample_df['销售日期_dt'] <= max_date_eod)]
    month_pos = s_month[s_month['结算金额'] > 0]
    sample_data[store] = {
        'month_qty': int(month_pos['数量'].sum()) if len(month_pos) > 0 else 0,
        'month_retail': round(float(month_pos['牌价额'].sum()), 2) if len(month_pos) > 0 else 0,
        'month_amt': round(float(s_month['结算金额'].sum()), 2),
    }
sample_region = {
    'month_qty': sum(v['month_qty'] for v in sample_data.values()),
    'month_retail': round(sum(v['month_retail'] for v in sample_data.values()), 2),
    'month_amt': round(sum(v['month_amt'] for v in sample_data.values()), 2),
}
print(f"   区域: 月样品¥{sample_region['month_amt']}({sample_region['month_qty']}件)")

# ============================================================
# 7. 输出JSON
# ============================================================
print("\n7. 输出monthly_analysis.json...")
output = {
    **store_data,
    '_region': region_data,
    '_series': series_by_store,
    '_products': products_by_store,
    '_guides': guide_by_store,
    '_sample': sample_data,
    '_sample_region': sample_region,
    '_new_products': new_products,
    '_new_products_detail': new_products_detail,
    '_conversion': conversion,
    '_meta': {
        'report_date': f'{REPORT_MONTH}月{report_day}日',
        'report_day': report_day,
        'report_day_index': report_day_index,
        'max_date': max_date.strftime('%Y-%m-%d'),
        'store_order': STORE_ORDER,
    },
}

with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"\n✅ 完成!")
print(f"   门店数: {len(STORE_ORDER)} | 导购数: {len(guide_monthly)} | 系列数: {len(region_series)} | 商品数: {len(region_products)}")
print(f"   数据已写入: {OUTPUT_JSON}")
