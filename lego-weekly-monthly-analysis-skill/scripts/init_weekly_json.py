#!/usr/bin/env python3
"""创建weekly_analysis.json基础结构（含daily数组），供extract_weekly_enhanced.py更新"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import pandas as pd
import json
import glob
import os

STORE_ORDER = ['扬州万象汇', '扬州京华城', '扬州江都金鹰', '泰州万象城', '宿迁宝龙', '淮安新亚', '淮安万象城']
STORE_CODE_MAP = {
    'LCS133-LW': '扬州万象汇', 'LCS275-LW': '扬州京华城', 'LCS364-LW': '扬州江都金鹰',
    'LCS0632-LW': '泰州万象城', 'LCS249-LW': '宿迁宝龙', 'LCS305-LW': '淮安新亚',
    'LCS0420-LW': '淮安万象城'
}
NEW_STORE = '淮安万象城'
EXCLUDE_NAMES = {'李婷', '黄小莉', '扬州金鹰-长期兼职', '泰州万象城-长期兼职'}
WK28_BASE = pd.Timestamp('2026-07-06')

# Read data
csvs = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'))
TODAY_XLSX = desktop_xlsxs[-1] if desktop_xlsxs else ''

hist_dfs = [pd.read_csv(f, low_memory=False) for f in csvs]
hist_df = pd.concat(hist_dfs, ignore_index=True)
if TODAY_XLSX and os.path.exists(TODAY_XLSX):
    today_df = pd.read_excel(TODAY_XLSX)
    all_df = pd.concat([hist_df, today_df], ignore_index=True)
else:
    all_df = hist_df

# 先统一解析日期（混合格式容错：CSV 2026-07-20 13:22:00 / xlsx 2026/7/20 13:22）
all_df['销售日期_dt'] = pd.to_datetime(all_df['销售日期'], format='mixed', errors='coerce')

# Dedup
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

# Exclude
all_df = all_df[~all_df['导购姓名'].astype(str).str.strip().isin(EXCLUDE_NAMES)]

# Date
all_df['销售日期_dt'] = pd.to_datetime(all_df['销售日期'], format='mixed', errors='coerce')
# 动态检测报告月份，保留近3个月数据
_max_date_raw = all_df['销售日期_dt'].max()
REPORT_MONTH = _max_date_raw.month
_date_3m_ago = _max_date_raw - pd.DateOffset(months=3)
all_df = all_df[all_df['销售日期_dt'] >= _date_3m_ago]

# Clean
all_df = all_df[all_df['是否赠品'] != '是']
all_df = all_df[~all_df['下单店仓名称'].str.contains('样品', na=False)]
all_df = all_df[all_df['结算金额'].astype(float) != 0]

# Store mapping
all_df['门店名'] = all_df['发货店仓编码'].map(STORE_CODE_MAP)
all_df = all_df[all_df['门店名'].notna()]
all_df['结算金额'] = all_df['结算金额'].astype(float)

# Determine WK dates — 动态从数据最大日期推算
max_date = all_df[all_df['销售日期_dt'].dt.month == REPORT_MONTH]['销售日期_dt'].max()
max_date_eod = pd.Timestamp(max_date.year, max_date.month, max_date.day, 23, 59, 59)
weeks_since_wk28 = (max_date - WK28_BASE).days // 7
WK_START = WK28_BASE + pd.Timedelta(weeks=weeks_since_wk28)
print(f"Max date: {max_date.strftime('%Y-%m-%d')}, WK start: {WK_START.strftime('%Y-%m-%d')}, Month: {REPORT_MONTH}")

# LT SKUs
LT_SKU_XLSX = sorted(glob.glob('/Users/a123/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/*/Caches/Files/2026-*/*/长尾款sku明细.xlsx'), key=os.path.getmtime)
lt_skus = set()
if LT_SKU_XLSX:
    lt_sku_df = pd.read_excel(LT_SKU_XLSX[-1])
    lt_skus = set(lt_sku_df['乐高编号'].apply(lambda x: str(int(x)) if pd.notna(x) else '').tolist())

# ⚠️ 必须在 wk29_df 过滤之前创建乐高编号_str 列
all_df['乐高编号_str'] = all_df['乐高编号'].apply(
    lambda v: str(int(v)) if pd.notna(v) and v == int(v) else str(v) if pd.notna(v) else ''
)

# Filter WK29 data (此时 all_df 已含乐高编号_str)
wk29_df = all_df[(all_df['销售日期_dt'] >= WK_START) & (all_df['销售日期_dt'] <= max_date_eod)]

# Build daily arrays for each store
weekly = {}
weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']

for store in STORE_ORDER:
    s_wk29 = wk29_df[wk29_df['门店名'] == store]
    daily = []
    
    # Get unique DATES (not timestamps) in WK29
    s_wk29_dates = s_wk29['销售日期_dt'].dt.normalize()
    dates = sorted(s_wk29_dates.unique())
    if not dates:
        # Also check if there are dates from the WK29 range even with 0 sales
        current = WK_START
        while current <= max_date_eod:
            dates.append(current)
            current += pd.Timedelta(days=1)
    
    for d in dates:
        day_data = s_wk29[s_wk29['销售日期_dt'].dt.normalize() == d]
        total = day_data['结算金额'].sum()
        wos = day_data[day_data['支付方式名称'].str.contains('WOS', na=False)]['结算金额'].sum()
        jl_mask = (day_data['支付方式名称'].str.contains('美团', na=False) |
                   day_data['特殊销售类型'].str.contains('饿了么', na=False) |
                   day_data['支付方式名称'].str.contains('淘宝', na=False) |
                   day_data['支付方式名称'].str.contains('京东', na=False))
        jl = day_data[jl_mask]['结算金额'].sum()
        day_lt = day_data[day_data['乐高编号_str'].isin(lt_skus)]
        lt = day_lt['结算金额'].sum()
        lt_qty = int(day_lt[day_lt['结算金额'] > 0]['数量'].sum()) if len(day_lt[day_lt['结算金额'] > 0]) > 0 else 0
        pos = day_data[day_data['结算金额'] > 0]
        tx_cnt = pos['单号'].nunique() if len(pos) > 0 else 0
        qty = int(pos['数量'].sum()) if len(pos) > 0 else 0
        jd = round(qty / tx_cnt, 2) if tx_cnt > 0 else 0
        kdj = round(total / tx_cnt, 2) if tx_cnt > 0 else 0
        
        daily.append({
            'date': f'{d.month}/{d.day}',
            'weekday': weekday_names[d.weekday()],
            'total': round(total, 2),
            'wos': round(wos, 2),
            'jl': round(jl, 2),
            'lt': round(lt, 2),
            'lt_qty': lt_qty,
            'tx_cnt': tx_cnt,
            'jd': jd,
            'kdj': round(kdj, 2),
        })
    
    # Weekly totals
    wk_total = s_wk29['结算金额'].sum()
    wk_wos = s_wk29[s_wk29['支付方式名称'].str.contains('WOS', na=False)]['结算金额'].sum()
    wk_jl_mask = (s_wk29['支付方式名称'].str.contains('美团', na=False) |
                 s_wk29['特殊销售类型'].str.contains('饿了么', na=False) |
                 s_wk29['支付方式名称'].str.contains('淘宝', na=False) |
                 s_wk29['支付方式名称'].str.contains('京东', na=False))
    wk_jl = s_wk29[wk_jl_mask]['结算金额'].sum()
    s_lt = s_wk29[s_wk29['乐高编号_str'].isin(lt_skus)]
    wk_lt = s_lt['结算金额'].sum()
    
    # Month amount (from dashboard) — 动态月份
    s_cur = all_df[(all_df['门店名'] == store) & (all_df['销售日期_dt'].dt.month == REPORT_MONTH)]
    month_amt = s_cur['结算金额'].sum()
    
    weekly[store] = {
        'daily': daily,
        'wk_amt': round(wk_total, 2),
        'wk_wos': round(wk_wos, 2),
        'wk_jl': round(wk_jl, 2),
        'wk_lt': round(wk_lt, 2),
        'wk_total': round(wk_total, 2),
        'month_amt': round(month_amt, 2),
    }
    print(f"  {store}: {len(daily)} days, wk_amt=¥{wk_total:,.0f}")

# Region data (from dashboard_latest.json)
with open('os.path.join(OUTPUTS_DIR, '')') as f:
    dash = json.load(f)
weekly['_region'] = dict(dash['region'])

# Save
with open('os.path.join(OUTPUTS_DIR, '')', 'w') as f:
    json.dump(weekly, f, ensure_ascii=False, indent=2)

print(f"\n✅ weekly_analysis.json created with {len(STORE_ORDER)} stores")
