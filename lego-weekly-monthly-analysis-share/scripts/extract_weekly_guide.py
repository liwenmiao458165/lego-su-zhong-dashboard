#!/usr/bin/env python3
"""
提取周维度数据：导购周销售/笔数/件数/WOS/长尾款 + 门店周同比
输出更新到 weekly_analysis.json
"""
import pandas as pd
import json
import os
import glob
from datetime import datetime

import wos_attr  # WOS 导购归因（含订单备注 [SA_NAME] 补充归因）

# ============================================================
# 配置（与 generate_dashboard.py 一致）
# ============================================================
STORE_ORDER = ['扬州万象汇', '扬州京华城', '扬州江都金鹰', '泰州万象城', '宿迁宝龙', '淮安新亚', '淮安万象城']
STORE_CODE_MAP = {
    'LCS133-LW': '扬州万象汇', 'LCS275-LW': '扬州京华城', 'LCS364-LW': '扬州江都金鹰',
    'LCS0632-LW': '泰州万象城', 'LCS249-LW': '宿迁宝龙', 'LCS305-LW': '淮安新亚',
    'LCS0420-LW': '淮安万象城'
}
NEW_STORE = '淮安万象城'
WK28_BASE = pd.Timestamp('2026-07-06')

HIST_CSVS = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'))
TODAY_XLSX = desktop_xlsxs[-1] if desktop_xlsxs else ''
TARGET_XLSX = sorted(glob.glob('/Users/a123/Desktop/。/**/*月指标同期.xlsx', recursive=True), key=os.path.getmtime)
TARGET_XLSX = TARGET_XLSX[-1] if TARGET_XLSX else ''
_cands = []
_cands += glob.glob('/Users/a123/Desktop/。/**/苏中区域*月员工指标分解.xlsx', recursive=True)
_cands += glob.glob('/Users/a123/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/*/Caches/Files/2026-*/*/苏中区域*月员工指标分解.xlsx')
GUIDE_XLSX = sorted(_cands, key=os.path.getmtime)[-1] if _cands else ''
LT_SKU_XLSX = sorted(glob.glob('/Users/a123/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/*/Caches/Files/2026-*/*/长尾款sku明细.xlsx'), key=os.path.getmtime)
LT_SKU_XLSX = LT_SKU_XLSX[-1] if LT_SKU_XLSX else ''

OUTPUT_JSON = '/Users/a123/WorkBuddy/Claw/outputs/weekly_analysis.json'

# ============================================================
# 1. 读取+清洗订单数据
# ============================================================
print("1. 读取订单数据...")
hist_dfs = [pd.read_csv(f, low_memory=False) for f in HIST_CSVS]
hist_df = pd.concat(hist_dfs, ignore_index=True)
if TODAY_XLSX and os.path.exists(TODAY_XLSX):
    if TODAY_XLSX.endswith('.csv'):
        today_df = pd.read_csv(TODAY_XLSX, low_memory=False)
    else:
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
print(f"   合并后总行数: {len(all_df)}")

# 日期处理
all_df['销售日期_dt'] = pd.to_datetime(all_df['销售日期'], format='mixed', errors='coerce')
# 上市日期解析（新品口径：年/月 == 报告年/月）
all_df['上市日期_dt'] = pd.to_datetime(all_df['上市日期'], errors='coerce')

# 排除赠品、样品、零金额
all_df = all_df[all_df['是否赠品'] != '是']
all_df = all_df[~all_df['下单店仓名称'].str.contains('样品', na=False)]
all_df = all_df[all_df['结算金额'].astype(float) != 0]

# 排除4类人员（销售数据完全排除，不参与任何统计）
EXCLUDE_NAMES = {'李婷', '黄小莉', '扬州金鹰-长期兼职', '泰州万象城-长期兼职'}
all_df = all_df[~all_df['导购姓名'].astype(str).str.strip().isin(EXCLUDE_NAMES)]

# 门店映射
all_df['门店名'] = all_df['发货店仓编码'].map(STORE_CODE_MAP)
all_df = all_df[all_df['门店名'].notna()]
all_df['结算金额'] = all_df['结算金额'].astype(float)

# 乐高编号统一
all_df['乐高编号_str'] = all_df['乐高编号'].apply(lambda v: str(int(v)) if pd.notna(v) and v == int(v) else str(v) if pd.notna(v) else '')

# WK周 — 先算max_date和WK_START，再按日期范围过滤
max_date = all_df['销售日期_dt'].max()
max_date_eod = pd.Timestamp(max_date.year, max_date.month, max_date.day, 23, 59, 59)
REPORT_MONTH = max_date.month
REPORT_YEAR = max_date.year
report_day = max_date.day
weeks_since_wk28 = (max_date - WK28_BASE).days // 7
WK_START = WK28_BASE + pd.Timedelta(weeks=weeks_since_wk28)
WK_NUM_INT = 28 + weeks_since_wk28

# 保留近4周数据（跨月自动适配）
_data_cutoff = WK_START - pd.Timedelta(weeks=3)
all_df = all_df[all_df['销售日期_dt'] >= _data_cutoff]

print(f"   当前周: WK{WK_NUM_INT}, 周起始: {WK_START.strftime('%m.%d')}, 数据截至: {max_date.strftime('%m.%d')}")
all_df['is_current_wk'] = (all_df['销售日期_dt'] >= WK_START) & (all_df['销售日期_dt'] <= max_date_eod)

# ============================================================
# 2. 读取长尾款SKU
# ============================================================
print("2. 读取长尾款SKU...")
lt_sku_df = pd.read_excel(LT_SKU_XLSX)
lt_skus = set(lt_sku_df.iloc[:, 0].apply(lambda x: str(int(x)) if pd.notna(x) else '').tolist())
print(f"   长尾款SKU数: {len(lt_skus)}")

# ============================================================
# 3. 读取导购指标
# ============================================================
print("3. 读取导购指标...")
guide_df = pd.read_excel(GUIDE_XLSX)

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
    if name == '李婷':
        name = '李婷1'
    if name == '黄小莉':
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
print(f"   有效导购数: {len(guide_info)}")

# ============================================================
# 4. 计算导购周维度数据（含笔数、件数）
# ============================================================
print("4. 计算导购周维度数据...")
wk_df = all_df[all_df['is_current_wk']].copy()

# WOS 导购归因（含订单备注 [SA_NAME] 补充归因）—— 与导购WOS重算报表口径一致
_wos_map = wos_attr.build_wos_map(
    wk_df, store_col='门店名', guide_col='导购姓名', note_col='备注',
    amt_col='结算金额', wos_col='支付方式名称', normalizer=normalize_guide_name)

guide_weekly = {}
for gname, info in guide_info.items():
    store = info['store']
    # 匹配导购姓名（含长期兼职前缀和支援名称）
    match_names = [gname, f'长期兼职-{gname}']
    if gname == '李婷1':
        match_names += ['扬万支援销售']
    if gname == '黄小莉1':
        match_names += ['泰州万象城-支援2', '泰州万象城支援2']

    g_wk = wk_df[(wk_df['门店名'] == store) & (wk_df['导购姓名'].astype(str).str.strip().isin(match_names))]

    if len(g_wk) == 0:
        guide_weekly[gname] = {
            'name': gname, 'store': store, 'title': info['title'],
            'target': info['target'],
            'wk_sales': 0, 'wk_qty': 0, 'wk_cnt': 0,
            'wk_atv': 0, 'wk_jd': 0,
            'wk_wos': 0, 'wk_lt': 0, 'wk_lt_qty': 0,
            'wk_np': 0,
            'wk_rate': 0,
        }
        continue

    # 正单（结算金额 > 0）
    pos = g_wk[g_wk['结算金额'] > 0]
    wk_sales = g_wk['结算金额'].sum()  # 含退货扣减
    wk_qty = int(pos['数量'].sum()) if len(pos) > 0 else 0
    # 笔数 = 正单的去重单号数
    wk_cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    wk_atv = round(wk_sales / wk_cnt, 2) if wk_cnt > 0 else 0
    wk_jd = round(wk_qty / wk_cnt, 2) if wk_cnt > 0 else 0

    # WOS（备注补充归因：导购姓名空 → 取订单备注 [SA_NAME]）
    wk_wos = wos_attr.wos_lookup(_wos_map, store, gname)

    # 长尾款
    g_lt = g_wk[g_wk['乐高编号_str'].isin(lt_skus)]
    wk_lt = g_lt['结算金额'].sum()
    wk_lt_qty = int(g_lt[g_lt['结算金额'] > 0]['数量'].sum()) if len(g_lt[g_lt['结算金额'] > 0]) > 0 else 0

    # 新品（上市日期 年/月 == 报告年/月）
    np_mask = (g_wk['上市日期_dt'].dt.year == REPORT_YEAR) & (g_wk['上市日期_dt'].dt.month == REPORT_MONTH)
    wk_np = g_wk[np_mask]['结算金额'].sum()

    # 月完成率
    wk_rate = round(wk_sales / info['target'] * 100, 1) if info['target'] > 0 else 0

    guide_weekly[gname] = {
        'name': gname, 'store': store, 'title': info['title'],
        'target': info['target'],
        'wk_sales': round(wk_sales, 2),
        'wk_qty': wk_qty,
        'wk_cnt': wk_cnt,
        'wk_atv': wk_atv,
        'wk_jd': wk_jd,
        'wk_wos': round(wk_wos, 2),
        'wk_lt': round(wk_lt, 2),
        'wk_lt_qty': wk_lt_qty,
        'wk_np': round(wk_np, 2),
        'wk_rate': wk_rate,
    }
    print(f"   {gname} ({store}): 周¥{wk_sales:,.0f} | {wk_cnt}笔 | {wk_qty}件 | 客单价¥{wk_atv} | 连带{wk_jd} | WOS¥{wk_wos:,.0f} | 长尾¥{wk_lt:,.0f}")

# 按门店分组
guide_by_store = {}
for gname, data in guide_weekly.items():
    store = data['store']
    if store not in guide_by_store:
        guide_by_store[store] = []
    guide_by_store[store].append(data)
# 每店按周销售降序
for store in guide_by_store:
    guide_by_store[store].sort(key=lambda x: x['wk_sales'], reverse=True)

# ============================================================
# 5. 计算门店周同比
# ============================================================
print("\n5. 计算门店周同比...")
# 周同比(去年同时段): 按实际周窗口 WK_START~max_date 逐日取对应月份指标同期表的同期行
# 跨月周(WK31=7/27-8/2)自动同时取7月表(7/27-7/31)与8月表(8/1-8/2)，彻底修复"分母只取当月"导致的假高同比
def _find_target_file(month):
    cands = sorted(glob.glob('/Users/a123/Desktop/。/**/*月指标同期.xlsx', recursive=True))
    for f in cands:
        if f'{month}月' in os.path.basename(f):
            return f
    return TARGET_XLSX

def _load_yoy_map(path):
    """返回 {store: {day: 同期值}, '_region': {day: 同期值}}（按月表内 day1-31 索引）"""
    df = pd.read_excel(path, header=None); m = {}; ncols = df.shape[1]
    for i in range(13, 20):
        name = str(df.iloc[i, 0]).strip(); d = {}
        for day in range(1, 32):
            col = 1 + day
            if col < ncols:
                v = df.iloc[i, col]
                d[day] = float(v) if pd.notna(v) else 0
        m[name] = d
    reg = {}
    if df.shape[0] > 20:
        for day in range(1, 32):
            col = 1 + day
            if col < ncols:
                v = df.iloc[20, col]
                reg[day] = float(v) if pd.notna(v) else 0
    m['_region'] = reg
    return m

def _safe_yoy(month):
    p = _find_target_file(month)
    if not p or not os.path.exists(p):
        return {}
    try:
        return _load_yoy_map(p)
    except Exception as e:
        print(f"   ⚠️ 读取{month}月同期失败: {e}")
        return {}

jul_yoy_map = _safe_yoy(7)
aug_yoy_map = _safe_yoy(8)

# 实际周窗口逐日(M,D)列表
WK_DAYS = []
_cur = WK_START.normalize()
_end = max_date.normalize()
while _cur <= _end:
    WK_DAYS.append((_cur.month, _cur.day))
    _cur = _cur + pd.Timedelta(days=1)
print(f"   周窗口日期: {WK_DAYS}")

def _yoy_of(store):
    tot = 0
    for (mth, d) in WK_DAYS:
        mp = jul_yoy_map if mth == 7 else aug_yoy_map
        tot += mp.get(store, {}).get(d, 0)
    return tot

store_wk_yoy = {}
for store in STORE_ORDER:
    if store != NEW_STORE and (store in jul_yoy_map or store in aug_yoy_map):
        store_wk_yoy[store] = _yoy_of(store)
region_wk_yoy = _yoy_of('_region')

print(f"   门店周同比(去年同时段 {['%d/%d' % (m, d) for m, d in WK_DAYS]}): {store_wk_yoy}")
print(f"   区域周同比: {region_wk_yoy}")

# ============================================================
# 6. 更新 weekly_analysis.json
# ============================================================
print("\n6. 更新 weekly_analysis.json...")
with open(OUTPUT_JSON, 'r', encoding='utf-8') as f:
    weekly = json.load(f)

# 更新导购数据
weekly['_guides'] = guide_by_store

# 更新门店周同比
for store in STORE_ORDER:
    if store in weekly and store != NEW_STORE:
        wk_actual = weekly[store].get('wk_amt', 0)
        wk_yoy_last = store_wk_yoy.get(store, 0)
        if wk_yoy_last > 0:
            weekly[store]['wk_yoy'] = round((wk_actual - wk_yoy_last) / wk_yoy_last * 100, 1)
        else:
            weekly[store]['wk_yoy'] = None
    elif store in weekly:
        weekly[store]['wk_yoy'] = None  # 淮安万象城无同期

# 区域周同比 — 排除淮安万象城（无同期数据）
region = weekly.get('_region', {})
# 用WK29实际（排除淮安万象城）÷ 同期（6店）
region_wk_actual = sum(weekly[s].get('wk_amt', 0) for s in STORE_ORDER if s in weekly and s != NEW_STORE)
if region_wk_yoy > 0:
    region['wk_yoy'] = round((region_wk_actual - region_wk_yoy) / region_wk_yoy * 100, 1)
else:
    region['wk_yoy'] = None
weekly['_region'] = region

with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(weekly, f, ensure_ascii=False, indent=2)

print(f"\n✅ 完成! 导购周维度数据 + 门店周同比已更新到 weekly_analysis.json")
print(f"   导购数: {len(guide_weekly)} | 门店数: {len(guide_by_store)}")
