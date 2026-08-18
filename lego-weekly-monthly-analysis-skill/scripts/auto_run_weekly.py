#!/usr/bin/env python3
"""
周维度+月维度分析 — 独立更新脚本
每周运行一次，生成周/月分析HTML页面
用法: python auto_run_weekly.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import glob
import os
import subprocess
import sys

import pandas as pd

# ============================================================
# 固定配置
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.environ.get('PROJECT_DIR', _HERE)
OUTPUTS_DIR = os.path.join(PROJECT_DIR, 'outputs')
ANALYSIS_DIR = os.path.join(PROJECT_DIR, 'outputs_analysis')
PYTHON_BIN = '/Users/a123/.workbuddy/binaries/python/envs/default/bin/python'

DOWNLOADS_DIR = '/Users/a123/Downloads'
DESKTOP_DIR = '/Users/a123/Desktop'

# ============================================================
# 1. 检测数据文件（仅用于日志确认，extract脚本自行glob查找）
# ============================================================
print("=" * 50)
print("📊 周维度+月维度分析 — 每周更新")
print("=" * 50)

# 找最新订单文件
csvs = sorted(glob.glob(os.path.join(DOWNLOADS_DIR, '销售订单*.csv')))
desktop_xlsxs = sorted(glob.glob(os.path.join(DESKTOP_DIR, '销售订单明细查询*.xlsx')))

if not csvs and not desktop_xlsxs:
    print("❌ 未找到销售订单文件！请确保Downloads或Desktop有数据文件")
    sys.exit(1)

# 检测报告日期
max_date = None
all_files = csvs + desktop_xlsxs
latest_file = max(all_files, key=os.path.getmtime)
try:
    if latest_file.endswith('.csv'):
        df = pd.read_csv(latest_file, low_memory=False)
    else:
        df = pd.read_excel(latest_file)
    if '销售日期' in df.columns:
        df['销售日期_dt'] = pd.to_datetime(df['销售日期'], format='mixed', errors='coerce')
        max_date = df['销售日期_dt'].max()
        print(f"✅ 最新数据日期: {max_date.strftime('%Y-%m-%d')} ({max_date.strftime('%A')})")
    else:
        print("⚠️ 文件缺少'销售日期'列，跳过日期检测")
except Exception as e:
    print(f"⚠️ 日期检测失败: {e}")

print(f"✅ 历史CSV: {len(csvs)}个")
print(f"✅ 今日文件: {os.path.basename(latest_file)}")

# ============================================================
# 2. 运行分析脚本
# ============================================================
print("\n🚀 运行周/月维度分析（HTML模板生成 + JSON提取 + 注入）...")

ANALYSIS_SCRIPTS = [
    # Step 1: 生成HTML模板（必须先于JSON提取，inject_json依赖模板中的const DATA占位符）
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

failed = []
for script_name, desc in ANALYSIS_SCRIPTS:
    script_path = os.path.join(PROJECT_DIR, script_name)
    print(f"\n   ▶ {desc} ({script_name})...")
    result = subprocess.run(
        [PYTHON_BIN, script_path],
        capture_output=True,
        text=True,
        cwd=PROJECT_DIR
    )
    if result.stdout:
        for line in result.stdout.strip().split('\n')[-5:]:
            print(f"     {line}")
    if result.stderr:
        print(f"     ⚠️ STDERR: {result.stderr[:500]}")
    if result.returncode != 0:
        print(f"     ❌ {script_name} 失败 (exit code {result.returncode})")
        failed.append(script_name)
    else:
        print(f"     ✅ {script_name} 完成")

# ============================================================
# 3. JS语法校验（防止变量重复声明等导致页面空白）
# ============================================================
print("\n🔍 JS语法校验...")
NODE_BIN = '/Users/a123/.workbuddy/binaries/node/versions/22.22.2/bin/node'
js_check_script = r"""
const fs = require('fs');
const files = [
  ['os.path.join(OUTPUTS_DIR, '')', 'weekly'],
  ['os.path.join(OUTPUTS_DIR, '')', 'monthly']
];
let hasError = false;
files.forEach(([path, label]) => {
  if (!fs.existsSync(path)) { console.log(label + ': file not found'); hasError = true; return; }
  const html = fs.readFileSync(path, 'utf8');
  const lines = html.split('\n');
  const dataLineIdx = lines.findIndex(l => l.includes('const DATA'));
  const jsLines = lines.filter((l, i) => i !== dataLineIdx);
  let js = jsLines.join('\n');
  const m = js.match(/<script>([\s\S]*)<\/script>/);
  if (!m) { console.log(label + ': no script tag found'); hasError = true; return; }
  js = m[1];
  js = 'const Chart=function(){};const document={getElementById:()=>({}),querySelectorAll:()=>[],addEventListener:()=>{}};' + js;
  try {
    new Function(js);
    console.log(label + ' JS: OK');
  } catch(e) {
    console.log(label + ' JS ERROR: ' + e.message);
    hasError = true;
  }
});
process.exit(hasError ? 1 : 0);
"""

js_result = subprocess.run(
    [NODE_BIN, '-e', js_check_script],
    capture_output=True, text=True
)
if js_result.stdout:
    for line in js_result.stdout.strip().split('\n'):
        print(f"   {line}")
if js_result.returncode != 0:
    print(f"\n❌ JS语法校验失败！页面将空白，请修复后再部署")
    print(f"   常见原因：const变量重复声明（如 const sd 在同一函数内声明两次）")
    sys.exit(1)
else:
    print("   ✅ JS语法校验通过")

# ============================================================
# 4. 检查输出文件 & 复制到独立分析目录
# ============================================================
import shutil

print("\n" + "=" * 50)
weekly_html = os.path.join(OUTPUTS_DIR, 'weekly_analysis.html')
monthly_html = os.path.join(OUTPUTS_DIR, 'monthly_analysis.html')

# 确保分析目录存在
os.makedirs(ANALYSIS_DIR, exist_ok=True)

for f in [weekly_html, monthly_html]:
    if os.path.exists(f):
        size_kb = os.path.getsize(f) // 1024
        print(f"✅ {os.path.basename(f)} ({size_kb}KB)")
        # 复制到独立分析目录
        dst = os.path.join(ANALYSIS_DIR, os.path.basename(f))
        shutil.copy2(f, dst)
        print(f"   → 已复制到 {ANALYSIS_DIR}")
    else:
        print(f"⚠️ 缺失: {os.path.basename(f)}")

# ============================================================
# 5. 生成落地页 index.html（动态日期，根治标签过期误导）
#    旧版 index.html 是静态手建文件，auto_run_weekly 只覆盖
#    weekly/monthly 两页、从不重写它，导致日期标签永远停在
#    手建时的 7/23。这里每次用最新数据日期重新生成，避免误导。
# ============================================================
# 落地页日期以月度分析 JSON 的 max_date 为准（extract 脚本合并全量数据后的真实报告日），
# 避免只用“最新单文件”的 max_date —— 当用户补传 6月文件时，单文件 max_date=6/30 会导致
# 整页标签误标成 6月（WK27），但分析数据实为 7月（WK30）。
import json as _json
_meta_path = os.path.join(OUTPUTS_DIR, 'monthly_analysis.json')
if os.path.exists(_meta_path):
    try:
        _m = _json.load(open(_meta_path, encoding='utf-8'))
        _md = _m.get('_meta', {}).get('max_date')
        if _md:
            max_date = pd.Timestamp(_md)
            print(f"✅ 落地页日期取自月度分析JSON: {max_date:%Y-%m-%d}")
    except Exception as e:
        print(f"⚠️ 读取月度JSON日期失败，回退到单文件检测: {e}")

if max_date is not None:
    wk_base = pd.Timestamp('2026-07-06')  # WK28 基准（周一）
    weeks_since = (max_date - wk_base).days // 7
    wk_num = 28 + weeks_since
    wk_start = max_date - pd.Timedelta(days=max_date.weekday())  # 本周一
    date_sub = f"数据更新于 {max_date:%Y-%m-%d} 收盘（WK{wk_num} / 7月1日–{max_date.month}月{max_date.day}日）"
    wk_label = f"WK{wk_num}（{wk_start:%m.%d}–{max_date:%m.%d}）门店/导购/系列/区域对比，含成交率、连带、长尾款、WOS 机会点诊断。"
    month_label = f"7月1日–{max_date.day}日 vs 上月同期，门店到区域机会点诊断（9 维度 Excel + 5 Tab 交互看板）。"
    index_html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>苏中区域 · 周/月维度分析</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#f5f6fa; font-family:-apple-system,BlinkMacSystemFont,'PingFang SC','Microsoft YaHei',sans-serif; color:#1a1a2e; padding:40px 20px; }
  .wrap { max-width:760px; margin:0 auto; }
  h1 { font-size:26px; font-weight:700; margin-bottom:6px; }
  .sub { color:#6b7280; font-size:14px; margin-bottom:32px; }
  .cards { display:grid; grid-template-columns:1fr 1fr; gap:20px; }
  a.card { text-decoration:none; color:inherit; background:#fff; border:1px solid #e8eaed; border-radius:14px; padding:28px 24px; transition:transform .15s, box-shadow .15s; display:block; }
  a.card:hover { transform:translateY(-4px); box-shadow:0 12px 28px rgba(0,0,0,0.10); }
  .icon { font-size:34px; }
  .t { font-size:19px; font-weight:700; margin:12px 0 6px; }
  .d { font-size:13px; color:#6b7280; line-height:1.6; }
  .bar { height:4px; border-radius:4px; margin-top:16px; }
  .bar.w { background:linear-gradient(90deg,#059669,#10b981); }
  .bar.m { background:linear-gradient(90deg,#7c3aed,#9333ea); }
  .foot { margin-top:36px; font-size:12px; color:#9ca3af; text-align:center; }
</style>
</head>
<body>
<div class="wrap">
  <h1>📊 苏中区域 · 周/月维度分析</h1>
  <div class="sub">__DATE_SUB__</div>
  <div class="cards">
    <a class="card" href="weekly_analysis.html">
      <div class="icon">🟢</div>
      <div class="t">周维度分析</div>
      <div class="d">__WK_LABEL__</div>
      <div class="bar w"></div>
    </a>
    <a class="card" href="monthly_analysis.html">
      <div class="icon">🟣</div>
      <div class="t">月维度分析</div>
      <div class="d">__MONTH_LABEL__</div>
      <div class="bar m"></div>
    </a>
  </div>
  <div class="foot">每日主看板请返回日维度链接 · 本页为周/月专项分析</div>
</div>
</body>
</html>'''
    index_html = (index_html
                  .replace('__DATE_SUB__', date_sub)
                  .replace('__WK_LABEL__', wk_label)
                  .replace('__MONTH_LABEL__', month_label))
    index_path = os.path.join(ANALYSIS_DIR, 'index.html')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(index_html)
    # 同时在根目录生成 analysis.html 着陆页（根目录显式文件名，绕开 OSS 子目录索引坑）
    root_index_path = os.path.join(OUTPUTS_DIR, 'analysis.html')
    with open(root_index_path, 'w', encoding='utf-8') as f:
        f.write(index_html)
    # 同步写进部署目录 ANALYSIS_DIR（修复根因：否则 outputs_analysis/ 缺 analysis.html，云端着陆页永远停在旧版）
    with open(os.path.join(ANALYSIS_DIR, 'analysis.html'), 'w', encoding='utf-8') as f:
        f.write(index_html)
    print(f"✅ 落地页 index.html / analysis.html 已生成: {max_date:%Y-%m-%d} (WK{wk_num})")

if failed:
    print(f"\n⚠️ {len(failed)}个脚本失败: {', '.join(failed)}")
    print("请检查错误信息并重试")
else:
    print("\n✅ 周维度+月维度分析更新完成！")

# ============================================================
# 6. 自动部署到固定链接（根目录显式文件，与主看板共存）
#    主通道：腾讯云 COS（默认域名直接渲染，不下载，企业网通）—— 2026-07-26 启用
#    回退：阿里云 OSS（默认域名强制下载，仅兜底）
#    再回退：GitHub Pages（企业网可能打不开）
# ============================================================
_analysis_link = None

def _read_secret(p):
    if os.path.exists(p):
        with open(p) as _f:
            return _f.read().strip()
    return ""

# ========== 部署：按「自定义域名」文件决定主通道，绝不输出会下载的链接 ==========
_cos_id = os.environ.get("COS_SECRET_ID", "") or _read_secret(os.path.join(PROJECT_DIR, '.cos_secret_id'))
_cos_key = os.environ.get("COS_SECRET_KEY", "") or _read_secret(os.path.join(PROJECT_DIR, '.cos_secret_key'))
_cos_bucket = os.environ.get("COS_BUCKET", "") or _read_secret(os.path.join(PROJECT_DIR, '.cos_bucket')) or "suzhong-hk-1458473628"
_cos_region = os.environ.get("COS_REGION", "") or _read_secret(os.path.join(PROJECT_DIR, '.cos_region')) or "ap-hongkong"
_cos_custom = os.environ.get("COS_CUSTOM_DOMAIN", "") or _read_secret(os.path.join(PROJECT_DIR, '.cos_custom_domain'))
_oss_ak = _read_secret(os.path.join(PROJECT_DIR, '.oss_ak'))
_oss_sk = _read_secret(os.path.join(PROJECT_DIR, '.oss_sk'))
_oss_bucket = os.environ.get("OSS_BUCKET", "") or _read_secret(os.path.join(PROJECT_DIR, '.oss_bucket')) or "suzhongquyu"
_oss_endpoint = os.environ.get("OSS_ENDPOINT", "") or _read_secret(os.path.join(PROJECT_DIR, '.oss_endpoint')) or "oss-cn-shanghai.aliyuncs.com"
_oss_custom = os.environ.get("OSS_CUSTOM_DOMAIN", "") or _read_secret(os.path.join(PROJECT_DIR, '.oss_custom_domain'))
_gh_owner = 'liwenmiao458165'
_gh_token_path = os.path.join(PROJECT_DIR, '.gh_token')
_gh_has_token = os.path.exists(_gh_token_path)
_gh_custom = os.environ.get("GH_CUSTOM_DOMAIN", "") or _read_secret(os.path.join(PROJECT_DIR, '.gh_custom_domain'))

if _cos_custom:
    ACTIVE = 'cos'
elif _gh_custom:
    ACTIVE = 'github'
elif _oss_custom:
    ACTIVE = 'oss'
else:
    ACTIVE = 'github' if _gh_has_token else 'cos'

_analysis_link = None

if ACTIVE == 'cos' and _cos_id and _cos_key:
    _env = dict(os.environ, COS_SECRET_ID=_cos_id, COS_SECRET_KEY=_cos_key,
                COS_BUCKET=_cos_bucket, COS_REGION=_cos_region,
                COS_CUSTOM_DOMAIN=_cos_custom, LOCAL_DIR=ANALYSIS_DIR, SUBPATH='', EXCLUDE_INDEX='true')
    print("\n🚀 部署周/月分析到腾讯云 COS 香港桶根目录" + ("（自定义域名直开不下载）..." if _cos_custom else "（待绑定自定义域名，仅备料）..."))
    _r = subprocess.run([PYTHON_BIN, os.path.join(PROJECT_DIR, 'deploy_cos.py')], env=_env, capture_output=True, text=True, timeout=300)
    print(_r.stdout)
    if _r.returncode != 0:
        print("⚠️ COS 部署失败："); print(_r.stderr[-800:])
    elif _cos_custom:
        _analysis_link = f"https://{_cos_custom}/analysis.html"
elif ACTIVE == 'github' and _gh_has_token:
    _gh_token = open(_gh_token_path).read().strip()
    _env = dict(os.environ, GH_TOKEN=_gh_token, GH_OWNER=_gh_owner, LOCAL_DIR=ANALYSIS_DIR, SUBPATH='', EXCLUDE_INDEX='true')
    print("\n🚀 部署周/月分析到 GitHub Pages（API 通道）" + ("（自定义域名直开不下载）..." if _gh_custom else "（github.io 直开，个人网络可用）..."))
    _r = subprocess.run([PYTHON_BIN, os.path.join(PROJECT_DIR, 'deploy_via_api.py')], env=_env, capture_output=True, text=True, timeout=300)
    print(_r.stdout)
    if _r.returncode != 0:
        print("⚠️ Pages 部署失败："); print(_r.stderr[-800:])
    elif _gh_custom:
        _analysis_link = f"https://{_gh_custom}/analysis.html"
    else:
        _analysis_link = f"https://{_gh_owner}.github.io/lego-su-zhong-dashboard/analysis.html"
elif ACTIVE == 'oss' and _oss_ak and _oss_sk:
    _env = dict(os.environ, OSS_AK=_oss_ak, OSS_SK=_oss_sk, OSS_BUCKET=_oss_bucket, OSS_ENDPOINT=_oss_endpoint, LOCAL_DIR=OUTPUTS_DIR, SUBPATH='')
    print("\n🚀 部署周/月分析到阿里云 OSS" + ("（自定义域名直开）..." if _oss_custom else "（默认域名会下载，仅兜底）..."))
    _r = subprocess.run([PYTHON_BIN, os.path.join(PROJECT_DIR, 'deploy_oss.py')], env=_env, capture_output=True, text=True, timeout=300)
    print(_r.stdout)
    if _r.returncode != 0:
        print("⚠️ OSS 部署失败："); print(_r.stderr[-800:])
    elif _oss_custom:
        _analysis_link = f"https://{_oss_custom}/analysis.html"
else:
    print("\n（无可用部署通道或缺少凭证，跳过自动部署）")

if _analysis_link:
    print(f"\n✅ 周/月分析固定链接: {_analysis_link}")
    print("（内容每次更新，地址不变）")
else:
    print("\n⚠️ 暂无可直开的周/月分析链接：当前仅备料，未配置自定义域名。")
    print("   绑定自定义域名后（建 .cos_custom_domain 或 .gh_custom_domain = suzhong-dashboard.com），重跑即得直开链接。")
    print("（内容每次更新，地址不变）")
