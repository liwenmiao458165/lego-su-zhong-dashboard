#!/usr/bin/env python3
"""
增强版周维度提取：WK28/WK29周环比 + 产品系列结构 + 导购个人周环比
输出更新到 weekly_analysis.json
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
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

# Find latest files automatically — 动态路径，适配任意月份
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

OUTPUT_JSON = 'os.path.join(OUTPUTS_DIR, '')'
OUTPUT_HTML = 'os.path.join(OUTPUTS_DIR, '')'

# 排除4类人员（销售数据完全排除，不参与任何统计）
EXCLUDE_NAMES = {'李婷', '黄小莉', '扬州金鹰-长期兼职', '泰州万象城-长期兼职'}

print(f"历史CSV: {len(csvs)}个")
print(f"今日文件: {os.path.basename(TODAY_XLSX) if TODAY_XLSX else '无(仅用CSV)'}")
print(f"导购指标: {os.path.basename(GUIDE_XLSX)}")
print(f"排除人员: {', '.join(sorted(EXCLUDE_NAMES))}")

# ============================================================
# 1. 读取+清洗订单数据
# ============================================================
print("\n1. 读取订单数据...")
hist_dfs = [pd.read_csv(f, low_memory=False) for f in csvs]
hist_df = pd.concat(hist_dfs, ignore_index=True)
if TODAY_XLSX and os.path.exists(TODAY_XLSX):
    today_df = pd.read_excel(TODAY_XLSX)
    all_df = pd.concat([hist_df, today_df], ignore_index=True)
else:
    all_df = hist_df
    print("   (无今日xlsx，仅使用CSV数据)")

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

# 保存样品仓订单（用于样品销售模块），后续从主数据排除
sample_df = all_df[all_df['下单店仓名称'].str.contains('样品', na=False)].copy()
sample_df = sample_df[sample_df['是否赠品'] != '是']
sample_df = sample_df[sample_df['结算金额'].astype(float) != 0]

# 排除赠品、样品、零金额
all_df = all_df[all_df['是否赠品'] != '是']
all_df = all_df[~all_df['下单店仓名称'].str.contains('样品', na=False)]
all_df = all_df[all_df['结算金额'].astype(float) != 0]

# 上市日期解析（新品口径：年/月 == 报告年/月）
all_df['上市日期_dt'] = pd.to_datetime(all_df['上市日期'], errors='coerce')

# 门店映射
all_df['门店名'] = all_df['发货店仓编码'].map(STORE_CODE_MAP)
all_df = all_df[all_df['门店名'].notna()]
all_df['结算金额'] = all_df['结算金额'].astype(float)

# 样品数据也做同样的门店映射
sample_df['门店名'] = sample_df['发货店仓编码'].map(STORE_CODE_MAP)
sample_df = sample_df[sample_df['门店名'].notna()]
sample_df['结算金额'] = sample_df['结算金额'].astype(float)
sample_df['牌价额'] = sample_df['牌价额'].astype(float)

# 乐高编号统一
all_df['乐高编号_str'] = all_df['乐高编号'].apply(
    lambda v: str(int(v)) if pd.notna(v) and v == int(v) else str(v) if pd.notna(v) else ''
)

# WK周期 — 先算max_date和WK_START，再按日期范围过滤
max_date = all_df['销售日期_dt'].max()
max_date_eod = pd.Timestamp(max_date.year, max_date.month, max_date.day, 23, 59, 59)
weeks_since_wk28 = (max_date - WK28_BASE).days // 7
WK_START = WK28_BASE + pd.Timedelta(weeks=weeks_since_wk28)
WK_NUM_INT = 28 + weeks_since_wk28

# 保留近4周数据（足够当前周+上周对比，跨月自动适配）
_data_cutoff = WK_START - pd.Timedelta(weeks=3)
all_df = all_df[all_df['销售日期_dt'] >= _data_cutoff]
sample_df = sample_df[sample_df['销售日期_dt'] >= _data_cutoff]
WK_PREV_START = WK_START - pd.Timedelta(days=7)
WK_PREV_END = WK_START - pd.Timedelta(seconds=1)   # 含上周日(7/19)全天，避免午夜边界漏掉当日致上周基数偏小、WoW虚高
WK_PREV_NUM = WK_NUM_INT - 1
print(f"   WK{WK_PREV_NUM}: {WK_PREV_START.strftime('%m.%d')}-{WK_PREV_END.strftime('%m.%d')}")
print(f"   WK{WK_NUM_INT}: {WK_START.strftime('%m.%d')}-{max_date.strftime('%m.%d')}")

all_df['is_wk28'] = (all_df['销售日期_dt'] >= WK_PREV_START) & (all_df['销售日期_dt'] <= WK_PREV_END)
all_df['is_wk29'] = (all_df['销售日期_dt'] >= WK_START) & (all_df['销售日期_dt'] <= max_date_eod)

# ============================================================
# 2. 上周/本周 门店周销售（周环比用）
# ============================================================
print("\n2. 计算上周/本周门店周销售...")
wk28_df = all_df[all_df['is_wk28']]
wk29_df = all_df[all_df['is_wk29']]

store_wk28 = {}
store_wk29 = {}
for s in STORE_ORDER:
    store_wk28[s] = round(wk28_df[wk28_df['门店名'] == s]['结算金额'].sum(), 2)
    store_wk29[s] = round(wk29_df[wk29_df['门店名'] == s]['结算金额'].sum(), 2)

region_wk28 = sum(store_wk28.values())
region_wk29 = sum(store_wk29.values())

for s in STORE_ORDER:
    wow = round((store_wk29[s] - store_wk28[s]) / store_wk28[s] * 100, 1) if store_wk28[s] > 0 else None
    print(f"   {s}: WK{WK_PREV_NUM}=¥{store_wk28[s]:,.0f} → WK{WK_NUM_INT}=¥{store_wk29[s]:,.0f} | 环比 {wow if wow is not None else 'N/A'}%")

region_wow = round((region_wk29 - region_wk28) / region_wk28 * 100, 1) if region_wk28 > 0 else None
print(f"   区域: WK{WK_PREV_NUM}=¥{region_wk28:,.0f} → WK{WK_NUM_INT}=¥{region_wk29:,.0f} | 环比 {region_wow}%")

# ============================================================
# 2.5 读取指标同期xlsx — 每日指标 + 去年同期每日数据
#     跨月周(WK31=7/27-8/2)需同时取7月表与8月表，按实际(月,日)索引，彻底修复错位
# ============================================================
print("\n2.5 读取每日指标和同期数据（7月+8月表合并，支持跨月周）...")

def _find_target_file(month):
    """按文件名含 '{month}月' 定位对应月份指标同期表"""
    cands = sorted(glob.glob('/Users/a123/Desktop/。/**/*月指标同期.xlsx', recursive=True))
    for f in cands:
        if f'{month}月' in os.path.basename(f):
            return f
    return TARGET_XLSX  # 回退

def _load_one_target(path, month):
    """读一张指标同期表 -> (store_tgt, store_yoy, reg_tgt, reg_yoy)，均按 day(1-31) 索引"""
    df = pd.read_excel(path, header=None)
    ncols = df.shape[1]
    st = {}; sy = {}; rt = {}; ry = {}
    for i in range(2, 9):
        name = str(df.iloc[i, 0]).strip()
        t = {}; y = {}
        for day in range(1, 32):
            col = 1 + day  # col2 = day1
            if col < ncols:
                vt = df.iloc[i, col]
                t[day] = float(vt) if pd.notna(vt) else 0
                vy = df.iloc[i + 11, col] if (i + 11) < df.shape[0] else None  # 同期行与门店行对齐: 行i -> 行i+11
                y[day] = float(vy) if (vy is not None and pd.notna(vy)) else 0
        st[name] = t; sy[name] = y
    for day in range(1, 32):
        col = 1 + day
        if col < ncols:
            vt = df.iloc[9, col]; rt[day] = float(vt) if pd.notna(vt) else 0
            vy = df.iloc[20, col] if df.shape[0] > 20 else None
            ry[day] = float(vy) if (vy is not None and pd.notna(vy)) else 0
    return st, sy, rt, ry

def _safe_load(month):
    p = _find_target_file(month)
    if not p or not os.path.exists(p):
        print(f"   ⚠️ 未找到{month}月指标同期表，回退TARGET_XLSX")
        p = TARGET_XLSX
    try:
        return _load_one_target(p, month)
    except Exception as e:
        print(f"   ⚠️ 读取{month}月表失败: {e}")
        return {}, {}, {}, {}

jul_st, jul_sy, jul_rt, jul_ry = _safe_load(7)
aug_st, aug_sy, aug_rt, aug_ry = _safe_load(8)

# 合并为 {name: {(month,day): value}}，跨月按实际月份取数（7月底取7月表，8月初取8月表）
store_daily_target = {}
store_daily_yoy = {}
for name in STORE_ORDER:
    td = {}; yd = {}
    for src_st, src_sy, mth in [(jul_st, jul_sy, 7), (aug_st, aug_sy, 8)]:
        if name in src_st:
            for d, v in src_st[name].items(): td[(mth, d)] = v
            for d, v in src_sy[name].items(): yd[(mth, d)] = v
    store_daily_target[name] = td
    store_daily_yoy[name] = yd
region_daily_target = {}
region_daily_yoy = {}
for d, v in jul_rt.items(): region_daily_target[(7, d)] = v
for d, v in jul_ry.items(): region_daily_yoy[(7, d)] = v
for d, v in aug_rt.items(): region_daily_target[(8, d)] = v
for d, v in aug_ry.items(): region_daily_yoy[(8, d)] = v

# 月指标(供其他逻辑/回退，无下游强依赖)
store_month_target = {n: 0 for n in STORE_ORDER}
region_month_target = 0

print(f"   每日指标: {len(store_daily_target)}家门店 (7月+8月表合并, 支持跨月周)")
print(f"   同期每日: {len(store_daily_yoy)}家门店")

# ============================================================
# 3. 本周产品系列结构（产品大类）
# ============================================================
print("\n3. 计算本周产品系列结构...")
series_by_store = {}
for s in STORE_ORDER:
    s_wk29 = wk29_df[wk29_df['门店名'] == s]
    series_amt = s_wk29.groupby('产品大类')['结算金额'].sum().round(2).to_dict()
    # Sort by amount descending
    series_amt = dict(sorted(series_amt.items(), key=lambda x: x[1], reverse=True))
    series_by_store[s] = series_amt
    top3 = list(series_amt.items())[:3]
    top3_str = ', '.join([f'{k}:¥{v:,.0f}' for k, v in top3])
    print(f"   {s}: {len(series_amt)}个系列, TOP3: {top3_str}")

# Region series
r_wk29 = wk29_df
region_series = r_wk29.groupby('产品大类')['结算金额'].sum().round(2).to_dict()
region_series = dict(sorted(region_series.items(), key=lambda x: x[1], reverse=True))
print(f"   区域: {len(region_series)}个系列, TOP3: {list(region_series.items())[:3]}")

# ============================================================
# 3.5 本周产品名称TOP10（商品名称，按结算金额，含件数）
# ============================================================
print("\n3.5 计算本周产品名称TOP10(含件数)...")
products_by_store = {}
for s in STORE_ORDER:
    s_wk29 = wk29_df[wk29_df['门店名'] == s]
    pos = s_wk29[s_wk29['结算金额'] > 0]
    prod_amt = pos.groupby('商品名称')['结算金额'].sum()
    prod_qty = pos.groupby('商品名称')['数量'].sum()
    prod_dict = {}
    for name in prod_amt.index:
        prod_dict[name] = {'amt': round(float(prod_amt[name]), 2), 'qty': int(prod_qty[name])}
    prod_dict = dict(sorted(prod_dict.items(), key=lambda x: x[1]['amt'], reverse=True))
    products_by_store[s] = prod_dict
    top3 = list(prod_dict.items())[:3]
    top3_str = ', '.join([f'{k}:¥{v["amt"]:,.0f}/{v["qty"]}件' for k, v in top3])
    print(f"   {s}: {len(prod_dict)}个商品, TOP3: {top3_str}")

# Region products with qty
r_pos = wk29_df[wk29_df['结算金额'] > 0]
r_prod_amt = r_pos.groupby('商品名称')['结算金额'].sum()
r_prod_qty = r_pos.groupby('商品名称')['数量'].sum()
region_products = {}
for name in r_prod_amt.index:
    region_products[name] = {'amt': round(float(r_prod_amt[name]), 2), 'qty': int(r_prod_qty[name])}
region_products = dict(sorted(region_products.items(), key=lambda x: x[1]['amt'], reverse=True))
products_by_store['_region'] = region_products
print(f"   区域: {len(region_products)}个商品, TOP3: {[(k, v['amt']) for k, v in list(region_products.items())[:3]]}")

# ============================================================
# 4. 导购数据：本周 + 上周（个人周环比）
# ============================================================
print("\n4. 计算导购周数据（含上周环比）...")
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

# 李婷1、黄小莉1 是扬万支援销售/泰州支援2重命名后的独立导购
# 导购指标表中只有李婷/黄小莉，需为李婷1/黄小莉1复制指标
if '李婷' in guide_info and '李婷1' not in guide_info:
    guide_info['李婷1'] = dict(guide_info['李婷'])
if '黄小莉' in guide_info and '黄小莉1' not in guide_info:
    guide_info['黄小莉1'] = dict(guide_info['黄小莉'])

# 去掉李婷、黄小莉本体（保留李婷1、黄小莉1）
guide_info.pop('李婷', None)
guide_info.pop('黄小莉', None)

# Read LT SKUs
if LT_SKU_XLSX:
    lt_sku_df = pd.read_excel(LT_SKU_XLSX)
    # 文件有8列：序号/乐高编号/商品名称/系列/零售价/上市时间/下市时间/产品等级
    # 必须用'乐高编号'列，不能用iloc[:,0]（那是序号1-79）
    lt_skus = set(lt_sku_df['乐高编号'].apply(lambda x: str(int(x)) if pd.notna(x) else '').tolist())
    # 提取近3个月新品（当月新品）
    sku_dates = pd.to_datetime(lt_sku_df['上市时间'], errors='coerce')
    recent_cutoff = pd.Timestamp('2026-05-01')
    new_products = lt_sku_df[sku_dates >= recent_cutoff]['商品名称'].dropna().tolist()
    print(f"   近3月新品: {len(new_products)}款")
else:
    lt_skus = set()
    new_products = []

def get_match_names(gname):
    """Get all possible name matches for a guide"""
    match_names = [gname, f'长期兼职-{gname}']
    if gname == '李婷1':
        match_names += ['扬万支援销售']
    if gname == '黄小莉1':
        match_names += ['泰州万象城-支援2', '泰州万象城支援2']
    return match_names

def calc_guide_weekly(wk_df, gname, info, wos_map):
    """Calculate guide weekly metrics from a filtered dataframe"""
    store = info['store']
    match_names = get_match_names(gname)
    g_wk = wk_df[(wk_df['门店名'] == store) & (wk_df['导购姓名'].astype(str).str.strip().isin(match_names))]

    if len(g_wk) == 0:
        return {
            'wk_sales': 0, 'wk_qty': 0, 'wk_cnt': 0,
            'wk_atv': 0, 'wk_jd': 0, 'wk_wos': 0, 'wk_lt': 0, 'wk_lt_qty': 0,
            'wk_np': 0,
        }

    pos = g_wk[g_wk['结算金额'] > 0]
    wk_sales = g_wk['结算金额'].sum()
    wk_qty = int(pos['数量'].sum()) if len(pos) > 0 else 0
    wk_cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    wk_atv = round(wk_sales / wk_cnt, 2) if wk_cnt > 0 else 0
    wk_jd = round(wk_qty / wk_cnt, 2) if wk_cnt > 0 else 0
    # WOS（备注补充归因：导购姓名空 → 取订单备注 [SA_NAME]）
    wk_wos = wos_attr.wos_lookup(wos_map, store, gname)
    g_lt = g_wk[g_wk['乐高编号_str'].isin(lt_skus)]
    wk_lt = g_lt['结算金额'].sum()
    wk_lt_qty = int(g_lt[g_lt['结算金额'] > 0]['数量'].sum()) if len(g_lt[g_lt['结算金额'] > 0]) > 0 else 0
    # 新品（上市日期 年/月 == 报告年/月）
    np_mask = (g_wk['上市日期_dt'].dt.year == max_date.year) & (g_wk['上市日期_dt'].dt.month == max_date.month)
    wk_np = g_wk[np_mask]['结算金额'].sum()
    return {
        'wk_sales': round(wk_sales, 2),
        'wk_qty': wk_qty,
        'wk_cnt': wk_cnt,
        'wk_atv': wk_atv,
        'wk_jd': wk_jd,
        'wk_wos': round(wk_wos, 2),
        'wk_lt': round(wk_lt, 2),
        'wk_lt_qty': wk_lt_qty,
        'wk_np': round(wk_np, 2),
    }

guide_weekly = {}

# WOS 导购归因（含订单备注 [SA_NAME] 补充归因）—— 与导购WOS重算报表口径一致
wk29_wos_map = wos_attr.build_wos_map(
    wk29_df, store_col='门店名', guide_col='导购姓名', note_col='备注',
    amt_col='结算金额', wos_col='支付方式名称', normalizer=normalize_guide_name)
wk28_wos_map = wos_attr.build_wos_map(
    wk28_df, store_col='门店名', guide_col='导购姓名', note_col='备注',
    amt_col='结算金额', wos_col='支付方式名称', normalizer=normalize_guide_name)

for gname, info in guide_info.items():
    wk29_data = calc_guide_weekly(wk29_df, gname, info, wk29_wos_map)
    wk28_data = calc_guide_weekly(wk28_df, gname, info, wk28_wos_map)
    
    wk_rate = round(wk29_data['wk_sales'] / info['target'] * 100, 1) if info['target'] > 0 else 0
    wk28_sales = wk28_data['wk_sales']
    wow = round((wk29_data['wk_sales'] - wk28_sales) / wk28_sales * 100, 1) if wk28_sales > 0 else None

    guide_weekly[gname] = {
        'name': gname, 'store': info['store'], 'title': info['title'],
        'target': info['target'],
        'wk_sales': wk29_data['wk_sales'],
        'wk_qty': wk29_data['wk_qty'],
        'wk_cnt': wk29_data['wk_cnt'],
        'wk_atv': wk29_data['wk_atv'],
        'wk_jd': wk29_data['wk_jd'],
        'wk_wos': wk29_data['wk_wos'],
        'wk_lt': wk29_data['wk_lt'],
        'wk_lt_qty': wk29_data['wk_lt_qty'],
        'wk_np': wk29_data['wk_np'],
        'wk_rate': wk_rate,
        'wk28_sales': round(wk28_sales, 2),
        'wow': wow,
    }
    wow_str = f"{wow:+.1f}%" if wow is not None else 'N/A'
    print(f"   {gname} ({info['store']}): 本周=¥{wk29_data['wk_sales']:,.0f} | 上周=¥{wk28_sales:,.0f} | 环比{wow_str}")

# 按门店分组
guide_by_store = {}
for gname, data in guide_weekly.items():
    store = data['store']
    if store not in guide_by_store:
        guide_by_store[store] = []
    guide_by_store[store].append(data)
for store in guide_by_store:
    guide_by_store[store].sort(key=lambda x: x['wk_sales'], reverse=True)

# ============================================================
# 4.5 样品销售数据
# ============================================================
print("\n4.5 计算样品销售数据...")
month_start = pd.Timestamp(max_date.year, max_date.month, 1)
sample_data = {}
for store in STORE_ORDER:
    s_wk = sample_df[(sample_df['门店名'] == store) & (sample_df['销售日期_dt'] >= WK_START) & (sample_df['销售日期_dt'] <= max_date_eod)]
    s_month = sample_df[(sample_df['门店名'] == store) & (sample_df['销售日期_dt'] >= month_start) & (sample_df['销售日期_dt'] <= max_date_eod)]
    wk_pos = s_wk[s_wk['结算金额'] > 0]
    month_pos = s_month[s_month['结算金额'] > 0]
    sample_data[store] = {
        'wk_qty': int(wk_pos['数量'].sum()) if len(wk_pos) > 0 else 0,
        'wk_retail': round(float(wk_pos['牌价额'].sum()), 2) if len(wk_pos) > 0 else 0,
        'wk_amt': round(float(s_wk['结算金额'].sum()), 2),
        'month_qty': int(month_pos['数量'].sum()) if len(month_pos) > 0 else 0,
        'month_retail': round(float(month_pos['牌价额'].sum()), 2) if len(month_pos) > 0 else 0,
        'month_amt': round(float(s_month['结算金额'].sum()), 2),
    }
sample_region = {
    'wk_qty': sum(v['wk_qty'] for v in sample_data.values()),
    'wk_retail': round(sum(v['wk_retail'] for v in sample_data.values()), 2),
    'wk_amt': round(sum(v['wk_amt'] for v in sample_data.values()), 2),
    'month_qty': sum(v['month_qty'] for v in sample_data.values()),
    'month_retail': round(sum(v['month_retail'] for v in sample_data.values()), 2),
    'month_amt': round(sum(v['month_amt'] for v in sample_data.values()), 2),
}
print(f"   区域: 周样品¥{sample_region['wk_amt']}({sample_region['wk_qty']}件), 月样品¥{sample_region['month_amt']}({sample_region['month_qty']}件)")

# ============================================================
# 4.6 门店周新品累计（新品口径：上市日期 年/月 == 报告年/月，与导购 wk_np 一致）
# ============================================================
print("\n4.6 计算门店周新品累计...")
np_store = {}
for store in STORE_ORDER:
    s_wk = wk29_df[(wk29_df['门店名'] == store) & (wk29_df['销售日期_dt'] >= WK_START) & (wk29_df['销售日期_dt'] <= max_date_eod)]
    np_mask = (s_wk['上市日期_dt'].dt.year == max_date.year) & (s_wk['上市日期_dt'].dt.month == max_date.month)
    np_store[store] = round(float(s_wk[np_mask]['结算金额'].sum()), 2)
np_region = round(sum(np_store.values()), 2)
region_wk_total = round(sum(store_wk29.values()), 2) if store_wk29 else 0
np_region_pct = round(np_region / region_wk_total * 100, 1) if region_wk_total > 0 else 0
print(f"   区域: 周新品累计¥{np_region}（占周销售{np_region_pct}%）")

# ============================================================
# 4.8 门店月成交率（成交率 = 月成交笔数 / 月客流，月累计）
# ============================================================
print("\n4.8 计算门店月成交率（笔数/客流，月累计）...")
# ⚠️ 月成交率分母 = 用户每月提供的全月累计客流文件(含7店|门店/客流两列)。
#    该文件已是全月累计(含本周WK), 直接采用, 不再叠加周客流(否则WK被重复计算)。
#    文件名按"M月客流.xlsx"惯例, 每月初更新。
FLOW_MONTH_FILE = '/Users/a123/Desktop/8月客流.xlsx'
FLOW_WK_FILE = None  # 已废弃: 历史逻辑=早期累计+本周客流拼接; 现改为单份全月文件, 不再使用
FLOW_MAP = {
    '扬州万象汇新店': '扬州万象汇', '扬州京华城': '扬州京华城', '扬州江都金鹰': '扬州江都金鹰',
    '泰州万象城': '泰州万象城', '宿迁宝龙': '宿迁宝龙', '淮安新亚广场': '淮安新亚',
    '淮安万象城': '淮安万象城'
}
REPORT_MONTH_WK = max_date.month
# ⚠️必须含年份过滤：本脚本glob了Downloads全部CSV(含2025年No.233906)，仅按month过滤会把去年同月混入当月累计
cur_month_df = all_df[(all_df['销售日期_dt'].dt.year == max_date.year) & (all_df['销售日期_dt'].dt.month == REPORT_MONTH_WK)]
flow_map = {}
try:
    # 8月客流.xlsx: 有表头(门店|客流)，全月累计客流(8/1起~截止上周末, 已含本周WK)
    # ⚠️ 该文件已是全月累计, 直接作为月成交率分母, 不再叠加周客流(避免WK重复计算)
    month_raw = pd.read_excel(FLOW_MONTH_FILE, sheet_name='Sheet1', header=0)
    month_raw = month_raw.rename(columns={month_raw.columns[0]: 'flow_store', month_raw.columns[1]: '客流'})
    month_raw['门店名'] = month_raw['flow_store'].map(FLOW_MAP)
    month_flow_map = month_raw.dropna(subset=['门店名']).set_index('门店名')['客流'].to_dict()
    flow_map = {s: int(month_flow_map.get(s, 0)) for s in STORE_ORDER}
    print(f"   月客流(全月累计)读取: 合计 {int(sum(month_flow_map.values()))} 人次, 覆盖 {len(month_flow_map)} 家店")
    print(f"   月累计客流(直接采用全月文件): 合计 {int(sum(flow_map.values()))} 人次")
    missing = [s for s in STORE_ORDER if flow_map.get(s, 0) == 0]
    if missing:
        print(f"   ⚠️ 以下门店客流无数据→月成交率显示'缺客流': {missing}")
except Exception as e:
    print(f"   ⚠️ 客流文件读取失败: {e}")
    flow_map = {}
print(f"   月客流合计: {int(sum(flow_map.values()))} 人次")
conversion = {}
for s in STORE_ORDER:
    s_m = cur_month_df[cur_month_df['门店名'] == s]
    pos = s_m[s_m['结算金额'] > 0]
    cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    fl = flow_map.get(s, 0)
    rate = round(cnt / fl * 100, 2) if fl > 0 else None  # 月客流缺失时置None，页面显示"缺客流"而非误导的0%
    conversion[s] = {'flow': int(fl), 'cnt': int(cnt), 'rate': rate}
rf = sum(conversion[s]['flow'] for s in STORE_ORDER)
rc = sum(conversion[s]['cnt'] for s in STORE_ORDER)
conversion['_region'] = {'flow': int(rf), 'cnt': int(rc), 'rate': round(rc / rf * 100, 2) if rf > 0 else 0}
print("   门店月成交率: " + " | ".join(f"{s} {conversion[s]['rate']}%" for s in STORE_ORDER) + f" | 区域 {conversion['_region']['rate']}%")

# ============================================================
# 4.9 门店WK周成交率（成交率 = 本周成交笔数 / 本周客流，周累计）
#    窗口跟随报告周动态计算（WK_START ~ max_date_eod），不再硬编码WK31
# ============================================================
print(f"\n4.9 计算门店WK{WK_NUM_INT}周成交率（笔数/客流，WK{WK_NUM_INT}周累计）...")
# ⚠️ 周客流文件名带周号(wkNN周客流.xlsx), 每周更新WK时需同步替换为当周文件, 否则会读到期周客流
WK_FLOW_FILE = '/Users/a123/Desktop/wk33周客流.xlsx'
FLOW_MAP_WK = {
    '扬州万象汇新店': '扬州万象汇', '扬州京华城': '扬州京华城', '扬州江都金鹰': '扬州江都金鹰',
    '泰州万象城': '泰州万象城', '宿迁宝龙': '宿迁宝龙', '淮安新亚广场': '淮安新亚',
    '淮安万象城': '淮安万象城'
}
wk_flow_df = all_df[(all_df['销售日期_dt'] >= WK_START) & (all_df['销售日期_dt'] <= max_date_eod)]
try:
    # 周客流.xlsx: Sheet1 = 苏中7店周客流(门店/客流)，与本WK窗口同口径；
    #   passenger_rank_* 大区排名表(含南京/合肥等)勿用，否则7店全落空→成交率显示"缺客流"
    wk_flow_raw = pd.read_excel(WK_FLOW_FILE, sheet_name='Sheet1', header=0)
    wk_flow_raw = wk_flow_raw.rename(columns={wk_flow_raw.columns[0]: 'flow_store', wk_flow_raw.columns[1]: '客流'})
    wk_flow_raw['门店名'] = wk_flow_raw['flow_store'].map(FLOW_MAP_WK)
    wk_flow_map = wk_flow_raw.dropna(subset=['门店名']).set_index('门店名')['客流'].to_dict()
    print(f"   WK{WK_NUM_INT}客流文件读取(Sheet1): 苏中7店WK{WK_NUM_INT}客流合计 {int(sum(wk_flow_map.values()))} 人次")
except Exception as e:
    print(f"   ⚠️ WK{WK_NUM_INT}客流文件读取失败: {e}")
    wk_flow_map = {}
weekly_conversion = {}
for s in STORE_ORDER:
    s_w = wk_flow_df[wk_flow_df['门店名'] == s]
    pos = s_w[s_w['结算金额'] > 0]
    cnt = pos['单号'].nunique() if len(pos) > 0 else 0
    fl = wk_flow_map.get(s, 0)
    rate = round(cnt / fl * 100, 2) if fl > 0 else None  # 客流缺失时置None，页面显示"缺客流"而非误导的0%
    weekly_conversion[s] = {'flow': int(fl), 'cnt': int(cnt), 'rate': rate}
rf = sum(weekly_conversion[s]['flow'] for s in STORE_ORDER)
rc = sum(weekly_conversion[s]['cnt'] for s in STORE_ORDER)
weekly_conversion['_region'] = {'flow': int(rf), 'cnt': int(rc), 'rate': round(rc / rf * 100, 2) if rf > 0 else None}
print("   WK%d门店周成交率: " % WK_NUM_INT + " | ".join(f"{s} {weekly_conversion[s]['rate']}%" for s in STORE_ORDER) + f" | 区域 {weekly_conversion['_region']['rate']}%")

# ============================================================
# 5. 更新 weekly_analysis.json
# ============================================================
print("\n5. 更新 weekly_analysis.json...")
with open(OUTPUT_JSON, 'r', encoding='utf-8') as f:
    weekly = json.load(f)

# 更新门店周环比和WK28 + 每日同比/完成率
for store in STORE_ORDER:
    if store in weekly:
        weekly[store]['wk28_amt'] = store_wk28.get(store, 0)
        weekly[store]['wk_np_amt'] = np_store.get(store, 0)
        wow = round((store_wk29[store] - store_wk28[store]) / store_wk28[store] * 100, 1) if store_wk28.get(store, 0) > 0 else None
        weekly[store]['wow'] = wow
        # 更新daily数组 — 添加target/rate/yoy_amt/yoy_pct
        daily_arr = weekly[store].get('daily', [])
        s_td = store_daily_target.get(store, {})
        s_yd = store_daily_yoy.get(store, {})
        for d_entry in daily_arr:
            # date="7/27" -> month=7, day=27；按实际月份到对应指标同期表取数（修复跨月错位）
            parts = d_entry['date'].split('/')
            mth = int(parts[0]); day = int(parts[1])
            t = s_td.get((mth, day), 0)
            y = s_yd.get((mth, day), 0)
            d_entry['target'] = round(t, 2)
            d_entry['rate'] = round(d_entry['total'] / t * 100, 1) if t > 0 else 0
            d_entry['yoy_amt'] = round(y, 2)
            d_entry['yoy_pct'] = round((d_entry['total'] - y) / y * 100, 1) if y > 0 else None

# 构建区域daily数组
region_daily = []
for d_entry in weekly[STORE_ORDER[0]].get('daily', []):
    parts = d_entry['date'].split('/')
    mth = int(parts[0]); day = int(parts[1])
    # 汇总所有门店当天数据
    r_total = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]]['total'] for s in STORE_ORDER if s in weekly)
    r_lt = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]].get('lt', 0) for s in STORE_ORDER if s in weekly)
    r_wos = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]].get('wos', 0) for s in STORE_ORDER if s in weekly)
    r_jl = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]].get('jl', 0) for s in STORE_ORDER if s in weekly)
    r_tx = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]].get('tx_cnt', 0) for s in STORE_ORDER if s in weekly)
    t = region_daily_target.get((mth, day), 0)
    y = region_daily_yoy.get((mth, day), 0)
    # 同比排除淮安万象城（无同期数据）
    r_total_no_new = sum(weekly[s]['daily'][[i for i, x in enumerate(weekly[s]['daily']) if x['date'] == d_entry['date']][0]]['total'] for s in STORE_ORDER if s in weekly and s != NEW_STORE)
    region_daily.append({
        'date': d_entry['date'],
        'weekday': d_entry['weekday'],
        'total': round(r_total, 2),
        'lt': round(r_lt, 2),
        'wos': round(r_wos, 2),
        'jl': round(r_jl, 2),
        'tx_cnt': r_tx,
        'target': round(t, 2),
        'rate': round(r_total / t * 100, 1) if t > 0 else 0,
        'yoy_amt': round(y, 2),
        'yoy_pct': round((r_total_no_new - y) / y * 100, 1) if y > 0 else None,
    })

# 更新产品系列结构
weekly['_series'] = series_by_store
series_by_store['_region'] = region_series

# 更新产品名称TOP10(含件数)
weekly['_products'] = products_by_store

# 更新导购数据
weekly['_guides'] = guide_by_store

# 更新样品销售数据
weekly['_sample'] = sample_data
weekly['_sample_region'] = sample_region
weekly['_region']['wk_np_amt'] = np_region

# 新增：当月新品列表
weekly['_new_products'] = new_products

# 新增：门店月成交率（笔数/客流）
weekly['_conversion'] = conversion

# 新增：门店WK周成交率（笔数/客流，窗口跟随报告周动态）
weekly['_weekly_conversion'] = weekly_conversion

# 更新区域数据
if '_region' in weekly:
    weekly['_region']['wk28_amt'] = region_wk28
    weekly['_region']['wk_amt'] = region_wk29
    weekly['_region']['wow'] = region_wow
    weekly['_region']['daily'] = region_daily
    weekly['_region']['wk_atv'] = round(region_wk29 / max(sum(d['tx_cnt'] for d in region_daily), 1), 2)

with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(weekly, f, ensure_ascii=False, indent=2)

print(f"\n✅ 完成!")
print(f"   门店数: {len(STORE_ORDER)} | 导购数: {len(guide_weekly)} | 系列数: {len(region_series)}")
print(f"   数据已写入: {OUTPUT_JSON}")
