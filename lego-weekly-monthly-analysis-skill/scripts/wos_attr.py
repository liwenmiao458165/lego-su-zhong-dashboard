"""WOS 导购归因共享模块。

规则：导购姓名有效时 → 用导购姓名；导购姓名空缺时 → 用订单备注 [SA_NAME=姓名] 补充归因。
一笔 WOS 只归一次，不会重复计。归因名会用调用方提供的 normalizer（各脚本的
normalize_guide_name）对齐到在册导购键（如 '扬万支援销售' → '李婷1'）。

数据特征（已验证，详见 导购WOS重算_含备注归因 报表）：
- 备注 [SA_NAME=] 提取的姓名与在册导购 100% 匹配，无噪音账号
- 凡导购姓名已填的订单，备注 SA_NAME 与之完全一致 → "仅补充空缺" 绝对安全，无冲突
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import re

import pandas as pd

# 需排除的账户（与项目内 EXCLUDE_NAMES 一致）
EXCLUDE_ACCOUNTS = {'李婷', '黄小莉', '扬州金鹰-长期兼职', '泰州万象城-长期兼职'}

_SA_RE = re.compile(r'\[SA_NAME=([^\]]+)\]')


def get_sa_name(note):
    """从订单备注提取 [SA_NAME=姓名] 的真实导购姓名；无则返回 ''。"""
    if note is None:
        return ''
    if isinstance(note, float) and pd.isna(note):
        return ''
    m = _SA_RE.search(str(note))
    return m.group(1).strip() if m else ''


def derive_wos_guide_name(guide_raw, note):
    """WOS 导购归因名：优先用订单导购姓名（有效且非排除账户），空缺则用备注 SA_NAME。

    返回 '' 表示该行无法归因（导购姓名空且备注无 SA_NAME）。
    """
    g = ''
    if guide_raw is not None and not (isinstance(guide_raw, float) and pd.isna(guide_raw)):
        g = str(guide_raw).strip()
    if g and g not in EXCLUDE_ACCOUNTS:
        return g
    return get_sa_name(note)


def build_wos_map(df, store_col='门店名', guide_col='导购姓名', note_col='备注',
                  amt_col='结算金额', wos_col='支付方式名称', normalizer=None):
    """对 df 内所有 WOS 行，按 (门店, 派生导购名) 归集金额。

    返回 dict: {(store, guide): amount}。guide 已按 normalizer 归一（若提供）。
    仅补充归因，不重复计：每行只归一次。
    """
    w = df[df[wos_col].fillna('').str.contains('WOS', na=False)]
    result = {}
    for _, r in w.iterrows():
        attr = derive_wos_guide_name(r[guide_col], r[note_col])
        if not attr:
            continue
        if normalizer:
            attr = normalizer(attr)
        if not attr:
            continue
        key = (r[store_col], attr)
        try:
            amt = float(r[amt_col])
        except (TypeError, ValueError):
            continue
        result[key] = result.get(key, 0) + amt
    return result


def wos_lookup(wos_map, store, gname):
    """查某导购的 WOS 金额（无则 0）。"""
    return wos_map.get((store, gname), 0)
