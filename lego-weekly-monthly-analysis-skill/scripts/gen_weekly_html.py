#!/usr/bin/env python3
"""Generate restructured weekly_analysis.html"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import os
import sys
import glob
import pandas as pd

OUTPUT = 'os.path.join(OUTPUTS_DIR, '')'

# ============================================================
# 动态检测数据日期（与extract脚本一致）
# ============================================================
csvs = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'))
all_files = csvs + desktop_xlsxs

max_date = None
for f in all_files:
    try:
        if f.endswith('.csv'):
            _df = pd.read_csv(f, low_memory=False)
        else:
            _df = pd.read_excel(f)
        if '销售日期' in _df.columns:
            _dates = pd.to_datetime(_df['销售日期'], errors='coerce')
            _max = _dates.max()
            if pd.notna(_max) and (max_date is None or _max > max_date):
                max_date = _max
    except Exception:
        continue

if max_date is None:
    print("ERROR: cannot determine max_date from data files")
    sys.exit(1)

WK28_BASE = pd.Timestamp('2026-07-06')
_weeks_since = (max_date - WK28_BASE).days // 7
WK_START = WK28_BASE + pd.Timedelta(weeks=_weeks_since)
WK_NUM_INT = 28 + _weeks_since
WK_PREV_NUM = WK_NUM_INT - 1

_DAY_NAMES = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
_start_str = f"{WK_START.month}月{WK_START.day}日({_DAY_NAMES[WK_START.dayofweek]})"
_end_str = f"{max_date.month}月{max_date.day}日({_DAY_NAMES[max_date.dayofweek]})"
_start_md = f"{WK_START.month}/{WK_START.day}"

print(f"  max_date={max_date.strftime('%Y-%m-%d')}, WK{WK_NUM_INT}: {WK_START.strftime('%m/%d')} - {max_date.strftime('%m/%d')}")

CSS = """<style>
:root{--bg:#f8fafc;--card-bg:#fff;--border:#e2e8f0;--text:#1e293b;--text-muted:#64748b;--red:#dc2626;--green:#16a34a;--blue:#0284c7;--orange:#ea580c;--purple:#7c3aed}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text)}
.header{background:linear-gradient(135deg,#1e293b 0%,#334155 100%);color:#fff;padding:20px 30px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.header h1{font-size:22px;font-weight:700}
.header .subtitle{font-size:13px;opacity:.7;margin-top:4px}
.store-tabs{display:flex;gap:8px;padding:12px 30px;background:#fff;border-bottom:1px solid var(--border);overflow-x:auto;position:sticky;top:64px;z-index:99}
.store-tab{padding:8px 20px;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;white-space:nowrap;border:2px solid var(--border);background:#fff;transition:all .2s}
.store-tab:hover{border-color:var(--purple)}
.store-tab.active{color:#fff;border-color:transparent}
.store-tab[data-store="扬州万象汇"].active{background:#7c3aed}
.store-tab[data-store="扬州京华城"].active{background:#0ea5e9}
.store-tab[data-store="扬州江都金鹰"].active{background:#059669}
.store-tab[data-store="泰州万象城"].active{background:#dc2626}
.store-tab[data-store="宿迁宝龙"].active{background:#ea580c}
.store-tab[data-store="淮安新亚"].active{background:#0891b2}
.store-tab[data-store="淮安万象城"].active{background:#9333ea}
.store-tab[data-store="区域对比"].active{background:#1e293b}
.container{max-width:1400px;margin:0 auto;padding:20px 30px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:24px}
.kpi-card{background:var(--card-bg);border-radius:12px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.08);border:1px solid var(--border);text-align:center}
.kpi-card .label{font-size:12px;color:var(--text-muted);margin-bottom:6px}
.kpi-card .value{font-size:24px;font-weight:800}
.kpi-card .sub{font-size:11px;margin-top:4px}
.kpi-card .value.red{color:var(--red)}.kpi-card .value.green{color:var(--green)}.kpi-card .value.blue{color:var(--blue)}.kpi-card .value.orange{color:var(--orange)}.kpi-card .value.purple{color:var(--purple)}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px}
.chart-grid.full{grid-template-columns:1fr}
.chart-card{background:var(--card-bg);border-radius:12px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);border:1px solid var(--border)}
.chart-card h3{font-size:15px;font-weight:700;margin-bottom:14px;display:flex;align-items:center;gap:6px}
.chart-card h3 .icon{font-size:18px}
.chart-wrap{position:relative;height:320px}
.chart-wrap.tall{height:380px}
.summary-table{width:100%;border-collapse:collapse;font-size:13px}
.summary-table th{background:#f1f5f9;padding:10px 12px;text-align:center;font-weight:700;border-bottom:2px solid var(--border)}
.summary-table td{padding:8px 12px;text-align:center;border-bottom:1px solid var(--border)}
.summary-table tr:hover{background:#f8fafc}
.summary-table .store-name{text-align:left;font-weight:600}
.summary-table .pos{color:var(--green);font-weight:700}.summary-table .neg{color:var(--red);font-weight:700}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700}
.badge.green{background:#dcfce7;color:#166534}.badge.red{background:#fee2e2;color:#991b1b}.badge.gray{background:#f1f5f9;color:#64748b}
.conclusion{font-size:13px;color:#475569;margin-top:10px;padding:0;background:#f8fafc;border-radius:6px;border-left:3px solid var(--purple);line-height:1.5;overflow:hidden}
.conc-title{padding:8px 12px 4px 12px;font-weight:700;color:#1e293b}
.conc-body{padding:0 12px 8px 12px}
.conc-opportunity{padding:6px 12px;background:#fef3c7;border-top:1px solid #f59e0b;font-size:12px;color:#92400e;line-height:1.6}
.conc-opportunity .conc-label{font-weight:700;color:#b45309}
.footer{text-align:center;padding:30px;color:var(--text-muted);font-size:12px}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(3,1fr)}.chart-grid{grid-template-columns:1fr}}
</style>"""

HTML_BODY = f"""<div class="header"><div style="display:flex;align-items:center;gap:16px;">
<a href="index.html" style="background:rgba(255,255,255,.15);color:#fff;padding:8px 16px;border-radius:8px;font-size:14px;font-weight:600;text-decoration:none;white-space:nowrap;">← 返回看板</a>
<div><h1>📊 苏中区域 WK{WK_NUM_INT} 门店周维度交互分析</h1>
<div class="subtitle">数据周期：2026年{_start_str} - {_end_str} | 生成时间：<span id="genTime"></span></div>
</div></div></div>
<div class="store-tabs" id="storeTabs">
<div class="store-tab active" data-store="区域对比" onclick="selectStore('区域对比')">📈 区域对比</div>
<div class="store-tab" data-store="扬州万象汇" onclick="selectStore('扬州万象汇')">🏪 扬州万象汇</div>
<div class="store-tab" data-store="扬州京华城" onclick="selectStore('扬州京华城')">🏪 扬州京华城</div>
<div class="store-tab" data-store="扬州江都金鹰" onclick="selectStore('扬州江都金鹰')">🏪 扬州江都金鹰</div>
<div class="store-tab" data-store="泰州万象城" onclick="selectStore('泰州万象城')">🏪 泰州万象城</div>
<div class="store-tab" data-store="宿迁宝龙" onclick="selectStore('宿迁宝龙')">🏪 宿迁宝龙</div>
<div class="store-tab" data-store="淮安新亚" onclick="selectStore('淮安新亚')">🏪 淮安新亚</div>
<div class="store-tab" data-store="淮安万象城" onclick="selectStore('淮安万象城')">🏪 淮安万象城</div>
</div>
<div class="container" id="content"></div>
<div class="footer">
<p>🧱 苏中区域销售看板 - 周维度分析 | 数据口径：结算金额(含退货扣减)，排除赠品/样品/零金额</p>
<p>WK{WK_NUM_INT}定义：{WK_START.month}月{WK_START.day}日({_DAY_NAMES[WK_START.dayofweek]})起 | 淮安万象城为新店，无同期数据(N/A)</p>
</div>"""

# Write the file in parts
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write('<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n')
    f.write('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
    f.write(f'<title>苏中区域 WK{WK_NUM_INT} 门店周维度交互分析</title>\n')
    f.write('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>\n')
    f.write(CSS + '\n</head>\n<body>\n')
    f.write(HTML_BODY + '\n')
    f.write('<script>\n')
    f.write('const DATA = {};\n')  # placeholder for inject_json.py

print(f"Part 1 written to {OUTPUT}")
print(f"File size so far: {os.path.getsize(OUTPUT)} bytes")
