#!/usr/bin/env python3
"""Part 1: HTML structure + CSS + body + script + constants + helpers + conclusions"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import os
import sys
import glob
from calendar import monthrange
import pandas as pd

# ============================================================
# 动态检测数据日期
# ============================================================
csvs = sorted(glob.glob('/Users/a123/Downloads/销售订单*.csv'))
desktop_xlsxs = sorted(glob.glob('/Users/a123/Desktop/销售订单明细查询*.xlsx'))
all_files = csvs + desktop_xlsxs
max_date = None
for f in all_files:
    try:
        _df = pd.read_csv(f, low_memory=False) if f.endswith('.csv') else pd.read_excel(f)
        if '销售日期' in _df.columns:
            _dates = pd.to_datetime(_df['销售日期'], errors='coerce')
            _max = _dates.max()
            if pd.notna(_max) and (max_date is None or _max > max_date):
                max_date = _max
    except Exception:
        continue
if max_date is None:
    sys.exit(1)

REPORT_MONTH = max_date.month
REPORT_DAY = max_date.day
# 上月同期
_prev_month = REPORT_MONTH - 1
_prev_year = max_date.year
if _prev_month == 0:
    _prev_month = 12
    _prev_year -= 1
_prev_last_day = monthrange(_prev_year, _prev_month)[1]
_prev_end_day = min(REPORT_DAY, _prev_last_day)

_month_title = f'{REPORT_MONTH}月'
_date_range = f'2026年{REPORT_MONTH}月1日-{REPORT_MONTH}月{REPORT_DAY}日'
_prev_range = f'{_prev_month}月1-{_prev_end_day}日'

print(f"  monthly: {_date_range} | 环比=上月同期({_prev_range})")

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>苏中区域 7月门店月度交互分析</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
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
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700}
.badge.green{background:#dcfce7;color:#166534}.badge.red{background:#fee2e2;color:#991b1b}.badge.gray{background:#f1f5f9;color:#64748b}
.conclusion{font-size:13px;color:#475569;margin-top:10px;padding:0;background:#f8fafc;border-radius:6px;border-left:3px solid var(--purple);line-height:1.5;overflow:hidden}
.conc-title{padding:8px 12px 4px 12px;font-weight:700;color:#1e293b}
.conc-body{padding:0 12px 8px 12px}
.conc-opportunity{padding:6px 12px;background:#fef3c7;border-top:1px solid #f59e0b;font-size:12px;color:#92400e;line-height:1.6}
.conc-opportunity .conc-label{font-weight:700;color:#b45309}
.footer{text-align:center;padding:30px;color:var(--text-muted);font-size:12px}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(3,1fr)}.chart-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="header"><div style="display:flex;align-items:center;gap:16px;">
<a href="index.html" style="background:rgba(255,255,255,.15);color:#fff;padding:8px 16px;border-radius:8px;font-size:14px;font-weight:600;text-decoration:none;white-space:nowrap;">← 返回看板</a>
<div><h1>📊 苏中区域 7月门店月度交互分析</h1>
<div class="subtitle">数据周期：2026年7月1日-7月18日 | 月环比=上月同时间区间(6月1-18日) | 生成时间：<span id="genTime"></span></div>
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
<p>🧱 苏中区域销售看板 - 月度分析 | 数据口径：结算金额(含退货扣减)，排除赠品/样品/零金额</p>
<p>月环比=vs 6月1-18日同期 | 淮安万象城为新店，无同比数据(N/A)</p>
</div>
<script>
const DATA = {};

const STORE_ORDER=['扬州万象汇','扬州京华城','扬州江都金鹰','泰州万象城','宿迁宝龙','淮安新亚','淮安万象城'];
const STORE_COLORS={'扬州万象汇':'#7c3aed','扬州京华城':'#0ea5e9','扬州江都金鹰':'#059669','泰州万象城':'#dc2626','宿迁宝龙':'#ea580c','淮安新亚':'#0891b2','淮安万象城':'#9333ea'};
const SC=['#7c3aed','#0ea5e9','#059669','#dc2626','#ea580c','#0891b2','#9333ea','#d97706','#65a30d','#4f46e5','#e11d48','#0d9488','#c026d3','#7c2d12','#1e40af','#831843','#365314','#854d0e','#3730a3','#9f1239'];
let charts={};
let currentStore='区域对比';

function fmtAmt(v){if(v===null||v===undefined)return 'N/A';return '¥'+Math.round(v).toLocaleString()}
function fmtPct(v){if(v===null||v===undefined)return 'N/A';return v.toFixed(1)+'%'}
function colorPct(v){if(v===null||v===undefined)return '';return v>=0?'green':'red'}
function isNA(v){return v===null||v===undefined}
function regVal(fn){return STORE_ORDER.reduce((a,s)=>a+fn(DATA[s]),0)}
function conc(c,o){if(!o)return '<div class="conclusion"><div class="conc-title">📋 结论</div><div class="conc-body">'+c+'</div></div>';return '<div class="conclusion"><div class="conc-title">📋 结论</div><div class="conc-body">'+c+'</div><div class="conc-opportunity"><span class="conc-label">💡 需关注的机会点</span><br>'+o+'</div></div>'}

function concKPI(s){
  const d=DATA[s],r=DATA['_region'];
  const atv=d.month_atv,jd=d.month_jd,ratv=r.month_atv,rjd=r.month_jd;
  const wpct=d.month_amt>0?d.month_wos/d.month_amt*100:0,rwpct=r.month_amt>0?r.month_wos/r.month_amt*100:0;
  const lpct=d.month_amt>0?d.month_lt/d.month_amt*100:0,rlpct=r.month_amt>0?r.month_lt/r.month_amt*100:0;
  const rate=d.month_rate,rrate=r.month_rate;
  const below=[];
  if(rate<rrate)below.push('月完成率'+rate.toFixed(1)+'%（区域'+rrate.toFixed(1)+'%，差'+(rrate-rate).toFixed(1)+'pct）');
  if(!isNA(d.yoy)&&!isNA(r.yoy)&&d.yoy<r.yoy)below.push('月同比'+d.yoy.toFixed(1)+'%（区域'+r.yoy.toFixed(1)+'%）');
  if(!isNA(d.mom)&&!isNA(r.mom)&&d.mom<r.mom)below.push('月环比'+d.mom.toFixed(1)+'%（区域'+r.mom.toFixed(1)+'%）');
  if(atv<ratv)below.push('客单价¥'+Math.round(atv)+'（区域¥'+Math.round(ratv)+'，差¥'+Math.round(ratv-atv)+'）');
  if(jd<rjd)below.push('连带率'+jd.toFixed(2)+'（区域'+rjd.toFixed(2)+'，差'+(rjd-jd).toFixed(2)+'）');
  if(wpct<rwpct)below.push('WOS占比'+wpct.toFixed(1)+'%（区域'+rwpct.toFixed(1)+'%）');
  if(lpct<rlpct)below.push('长尾款占比'+lpct.toFixed(1)+'%（区域'+rlpct.toFixed(1)+'%）');
  const c=s+'KPI汇总：月销售'+fmtAmt(d.month_amt)+'，完成率'+rate.toFixed(1)+'%，连带率'+jd.toFixed(2)+'，客单价¥'+Math.round(atv)+'。';
  let o='';
  if(below.length>0)o='低于区域整体的KPI：'+below.join('；')+'。需逐项分析差距原因，制定提升计划。';
  else o='各项KPI均达到或优于区域整体水平，保持当前运营策略。';
  return conc(c,o);
}

function concSeries(s){
  const d=DATA[s];const sd=(DATA['_series']||{})[s]||{};
  const ents=Object.entries(sd).sort((a,b)=>b[1]-a[1]);
  if(ents.length===0)return conc('暂无系列数据。');
  const total=d.month_amt||1;
  const top3=ents.slice(0,3);
  const pct=top3.reduce((a,[k,v])=>a+v,0)/total*100;
  const topPcts=ents.slice(0,10).map(([k,v])=>k+'('+fmtAmt(v)+',占比'+(v/total*100).toFixed(1)+'%)');
  const c='TOP10系列：'+topPcts.join('、')+'；TOP3合计占比'+pct.toFixed(1)+'%。';
  let o='';
  if(pct>70)o='TOP3系列过于集中，建议丰富中长尾系列推荐降低依赖风险。';else if(pct<35)o='系列分布较散，建议打造1-2个核心系列重点陈列。';else o='系列结构良好，保持均衡推荐同时强化核心系列。';
  const focusTop3=['CITY','TECHNIC','NINJAGO'];
  const missingTop3=focusTop3.filter(n=>!top3.some(([k])=>k.toUpperCase().includes(n)));
  if(missingTop3.length>0)o+='重点系列'+missingTop3.join('、')+'未进TOP3（排名需提升），建议加强陈列和推荐。';
  const inTop3=focusTop3.filter(n=>top3.some(([k])=>k.toUpperCase().includes(n)));
  if(inTop3.length>0)o+='重点系列'+inTop3.join('、')+'在TOP3表现良好。';
  const top5=ents.slice(0,5);
  const minecraftInTop5=top5.some(([k])=>k.toUpperCase().includes('MINECRAFT'));
  if(!minecraftInTop5)o+='MINECRAFT未进TOP5，需关注沙盒游戏IP粉丝群体引流和推荐。';
  // BOTANICALS植物系列销售占比
  const botanicalAmt=ents.reduce((a,[k,v])=>k.toUpperCase().includes('BOTANICAL')?a+v:a,0);
  const botanicalPct=botanicalAmt/total*100;
  const botanicalStr=botanicalAmt>0?'BOTANICALS植物系列占比'+botanicalPct.toFixed(1)+'%。':'BOTANICALS植物系列本月无销售。';
  // 结论中加入BOTANICALS占比
  const c2=c+botanicalStr;
  if(botanicalAmt>0){
    o+='BOTANICALS植物系列销售占比'+botanicalPct.toFixed(1)+'%（指标4.1%）';
    if(botanicalPct<4.1)o+='，低于指标，需加强植物系列陈列和推荐，抓住礼品/家居装饰客群。';
    else o+='，达到指标。';
  }else{
    o+='BOTANICALS植物系列本月无销售，低于指标4.1%，建议关注植物系列铺货和推荐。';
  }
  return conc(c2,o);
}
function concRadar(s){
  const d=DATA[s];const r=DATA['_region'];
  const atv=d.month_atv,ratv=r.month_atv,jd=d.month_jd,rjd=r.month_jd;
  const cnt=d.month_cnt,rcnt=r.month_cnt/7;
  const wpct=d.month_amt>0?d.month_wos/d.month_amt*100:0,rwpct=r.month_amt>0?r.month_wos/r.month_amt*100:0;
  const lpct=d.month_amt>0?d.month_lt/d.month_amt*100:0,rlpct=r.month_amt>0?r.month_lt/r.month_amt*100:0;
  const items=[['客单价',atv>ratv],['连带率',jd>rjd],['笔数',cnt>rcnt],['WOS占比',wpct>rwpct],['长尾款占比',lpct>rlpct]];
  const lead=items.filter(x=>x[1]).map(x=>x[0]),weak=items.filter(x=>!x[1]).map(x=>x[0]);
  const c=s+'在'+(lead.length>0?lead.join('、'):'无')+'维度领先区域'+(weak.length>0?'，'+weak.join('、')+'维度低于区域需提升':'，各维度均优于区域，综合能力突出')+'。';
  let o='';
  if(weak.length>0){
    const gaps=[];
    if(weak.includes('客单价'))gaps.push('客单价差¥'+Math.round(ratv-atv)+'（¥'+Math.round(atv)+' vs ¥'+Math.round(ratv)+'）');
    if(weak.includes('连带率'))gaps.push('连带率差'+(rjd-jd).toFixed(2)+'（'+jd.toFixed(2)+' vs '+rjd.toFixed(2)+'）');
    if(weak.includes('笔数'))gaps.push('月笔数差'+(rcnt-cnt).toFixed(0)+'（'+cnt+' vs '+rcnt.toFixed(0)+'）');
    if(weak.includes('WOS占比'))gaps.push('WOS占比差'+(rwpct-wpct).toFixed(1)+'%（'+wpct.toFixed(1)+'% vs '+rwpct.toFixed(1)+'%）');
    if(weak.includes('长尾款占比'))gaps.push('长尾款占比差'+(rlpct-lpct).toFixed(1)+'%（'+lpct.toFixed(1)+'% vs '+rlpct.toFixed(1)+'%）');
    o='低于区域均值的维度：'+gaps.join('；')+'。';
  }
  if(lead.length>=4&&!weak.length)o+='多维度领先区域，保持优势策略可推广至其他门店。';
  if(!o)o='各维度均优于区域水平，综合运营能力突出，保持当前策略。';
  return conc(c,o);
}
function concLongTail(s){
  const d=DATA[s],r=DATA['_region'];
  const lpct=d.month_amt>0?d.month_lt/d.month_amt*100:0;
  const rl=r.month_amt>0?r.month_lt/r.month_amt*100:0;
  const c=s+'本月长尾款'+fmtAmt(d.month_lt)+'，占比'+lpct.toFixed(1)+'%；区域长尾款占比'+rl.toFixed(1)+'%。';
  let o='';if(lpct<rl-2)o='长尾款占比低于区域均值'+(rl-lpct).toFixed(1)+'个百分点，建议加强79款SKU推荐话术培训，收银台增加长尾小件陈列位。';else if(lpct>rl+2)o='长尾款占比高于区域均值'+(lpct-rl).toFixed(1)+'个百分点，继续保持关联推荐策略。';else o='长尾款与区域持平，关注件单价提升和连带搭配。';
  const rlt=regVal(x=>x.month_lt)/7;
  if(d.month_lt<rlt*0.7)o+='长尾款金额低于区域门店均值（¥'+Math.round(rlt)+'），需重点跟进长尾款推荐。';
  if(d.month_lt_qty<10)o+='月度长尾件数偏低，需持续推荐确保占比达标。';
  return conc(c,o);
}
function concMom(s){
  const d=DATA[s],r=DATA['_region'];
  const c=s+'月环比'+(isNA(d.mom)?'N/A':(d.mom>=0?'+':'')+d.mom.toFixed(1)+'%')+'，'+(isNA(r.mom)?'区域N/A':'区域'+(r.mom>=0?'+':'')+r.mom.toFixed(1)+'%')+'。';
  let o='';if(!isNA(d.mom)){if(d.mom<-10)o='月环比下滑明显，需复盘本月客流和转化率下降原因，排查天气/活动/竞品影响。';else if(d.mom<-3)o='月环比小幅下滑，关注下半月恢复节奏。';else if(d.mom<3)o='月环比基本持平，需寻找增长突破点。';else o='月环比增长良好，总结本月成功经验复制。';}else o='新店需建立基准月数据后再分析环比趋势。';
  return conc(c,o);
}
function concTop(s){
  const d=DATA[s];const pd=(DATA['_products']||{})[s]||{};
  const ents=Object.entries(pd).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]);
  if(ents.length===0)return conc('暂无产品数据。');
  const [n,a,q]=ents[0];const pct=d.month_amt>0?a/d.month_amt*100:0;
  const c='TOP1产品「'+n+'」销售'+fmtAmt(a)+'（'+q+'件），占月销售'+pct.toFixed(1)+'%。';
  let o='';if(pct>25)o='TOP1单品依赖度过高（>25%），警惕库存不足和缺货风险，建议培养第二梯队爆款。';
  // check 当月新品是否进入TOP10
  const newProds=DATA['_new_products']||[];
  if(newProds.length>0){
    const top10Names=ents.slice(0,10).map(([k])=>k);
    const newInTop=newProds.filter(np=>top10Names.some(tn=>tn.includes(np)||np.includes(tn)));
    if(newInTop.length>0)o+='当月新品「'+newInTop.join('、')+'」进入TOP10，表现突出，建议加大陈列和推荐力度。';
    else o+='当月新品无一款进入TOP10，需关注新品陈列位置和推荐话术，加强新品首周曝光。';
  }
  // check 区域TOP10产品门店是否也在TOP10
  const rpd=(DATA['_products']||{})['_region']||{};
  const rEnts=Object.entries(rpd).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const storeTop10=new Set(ents.slice(0,10).map(([k])=>k));
  const rInStore=rEnts.filter(([k])=>storeTop10.has(k));
  const rNotInStore=rEnts.filter(([k])=>!storeTop10.has(k));
  if(rInStore.length>0)o+='区域TOP10中'+rInStore.length+'款门店也有在TOP10（'+rInStore.map(([k])=>k).join('、')+'）';
  if(rNotInStore.length>0)o+=(rInStore.length>0?'，':'')+'区域TOP10中'+rNotInStore.length+'款门店未进入TOP10（'+rNotInStore.map(([k])=>k).join('、')+'），建议对齐区域爆款推荐策略。';
  else if(rInStore.length===10)o+='门店TOP10与区域完全重合，产品结构高度一致。';
  return conc(c,o);
}
function concNew(key){
  const isR = (key==='区域对比'||key==='_region');
  const d = isR ? DATA['_region'] : DATA[key];
  const r = DATA['_region'];
  const detailKey = isR ? '_region' : key;
  const na=d.new_amt||0, rate=d.new_rate||0, rRate=r.new_rate||0;
  const label = isR ? '区域' : key;
  const c = label+'本月新品（上市日期在本月）销售'+fmtAmt(na)+'，占月销售'+rate.toFixed(1)+'%（'+(isR?('区域'+rate.toFixed(1)):'区域'+rRate.toFixed(1))+'%）。';
  let o='';
  if(!isR){
    if(rate < rRate-1) o='新品占比低于区域'+(rRate-rate).toFixed(1)+'个百分点，建议加强当月新品首周陈列、堆头和导购主推话术。';
    else if(rate > rRate+1) o='新品占比高于区域，新品推荐动销突出，保持并提炼陈列经验推广。';
    else o='新品占比与区域持平，维持当前新品铺货和推荐节奏。';
  } else {
    if(rRate>=10) o='新品占比达'+rRate.toFixed(1)+'%，动销良好，建议持续强化新品首周曝光与跨店调货保障。';
    else o='新品占比偏低（'+rRate.toFixed(1)+'%），需统筹区域新品铺货、陈列标准与导购主推话术。';
  }
  const detail=(DATA['_new_products_detail']||{})[detailKey]||[];
  if(detail.length>0) o+='TOP新品：'+detail.slice(0,3).map(x=>x.name+'('+fmtAmt(x.amt)+')').join('、')+'。';
  return conc(c,o);
}
function concGuideSales(s){
  const gs=(DATA['_guides']||{})[s]||[];if(gs.length===0)return conc('暂无导购数据。');
  const top=gs.reduce((a,b)=>b.m_sales>a.m_sales?b:a);
  const tot=gs.reduce((a,g)=>a+g.m_sales,0);
  const c='TOP1导购'+top.name+'月销售'+fmtAmt(top.m_sales)+'（完成率'+top.m_rate.toFixed(1)+'%），门店导购合计'+fmtAmt(tot)+'。';
  let o='';const stAvg=gs.length>0?tot/gs.length:0;
  const allRegGuides=[];STORE_ORDER.forEach(t=>{(DATA['_guides']||{})[t]?.forEach(g=>allRegGuides.push(g));});
  const regAvg=allRegGuides.length>0?allRegGuides.reduce((a,g)=>a+g.m_sales,0)/allRegGuides.length:0;
  const topPct=tot>0?top.m_sales/tot*100:0;if(topPct>50)o='TOP导购销售占门店'+topPct.toFixed(0)+'%，建议头部经验分享复制，缩小导购间差距。';
  const belowSt=gs.filter(g=>g.m_sales>0&&g.m_sales<stAvg*0.7);
  if(belowSt.length>0)o+=belowSt.map(g=>g.name+'('+fmtAmt(g.m_sales)+')').join('、')+'低于门店均值'+fmtAmt(stAvg)+'，需重点跟进销售技巧辅导。';
  const belowReg=gs.filter(g=>g.m_sales>0&&g.m_sales<regAvg*0.7);
  if(belowReg.length>0)o+=belowReg.map(g=>g.name).join('、')+'低于区域导购均值'+fmtAmt(regAvg)+'，建议安排导师带教和排班优化。';
  const low=gs.filter(g=>g.m_rate<30&&g.m_sales>0);if(low.length>0)o+=low.length+'位导购完成率不足30%，需一对一辅导。';
  if(!o)o='导购团队整体表现均衡（门店人均'+fmtAmt(stAvg)+'，区域人均'+fmtAmt(regAvg)+'），继续保持团队协作氛围。';
  return conc(c,o);
}
function concGuideJdKdj(s){
  const d=DATA[s],r=DATA['_region'];
  const c=s+'月连带率'+d.month_jd.toFixed(2)+'，客单价¥'+Math.round(d.month_atv)+'；区域连带率'+r.month_jd.toFixed(2)+'，客单价¥'+Math.round(r.month_atv)+'。';
  let o='';if(d.month_jd<r.month_jd-0.1)o='连带率低于区域均值'+(r.month_jd-d.month_jd).toFixed(2)+'（'+d.month_jd.toFixed(2)+' vs '+r.month_jd.toFixed(2)+'），建议强化跨品类搭配推荐话术培训。';else if(d.month_jd<r.month_jd)o='连带率略低于区域（差'+(r.month_jd-d.month_jd).toFixed(2)+'），关注搭配推荐微调。';else o='连带率表现良好（高于区域'+(d.month_jd-r.month_jd).toFixed(2)+'），关注高端产品推荐提升客单价。';
  if(d.month_atv<r.month_atv*0.85)o+='客单价低于区域¥'+Math.round(r.month_atv-d.month_atv)+'（¥'+Math.round(d.month_atv)+' vs ¥'+Math.round(r.month_atv)+'），建议引导客户关注高价套装和限定款。';
  else if(d.month_atv<r.month_atv)o+='客单价略低于区域¥'+Math.round(r.month_atv-d.month_atv)+'，适度推荐高单价产品。';
  return conc(c,o);
}
function concGuideLt(s){
  const gs=(DATA['_guides']||{})[s]||[];if(gs.length===0)return conc('暂无导购数据。');
  const top=gs.reduce((a,b)=>b.m_lt>a.m_lt?b:a);
  const tot=gs.reduce((a,g)=>a+g.m_lt,0),tq=gs.reduce((a,g)=>a+g.m_lt_qty,0);
  const c='长尾款TOP导购'+top.name+'：'+fmtAmt(top.m_lt)+'（'+top.m_lt_qty+'件）；门店导购长尾款合计'+fmtAmt(tot)+'（'+tq+'件）。';
  let o='';const ltAvg=gs.length>0?tot/gs.length:0;
  const lowLt=gs.filter(g=>g.m_lt===0||(g.m_sales>0&&g.m_lt/g.m_sales<0.03));
  if(lowLt.length>0)o+=lowLt.map(g=>g.name+'(¥'+(g.m_lt||0).toFixed(0)+')').join('、')+'长尾款占比偏低（门店均值¥'+Math.round(ltAvg)+'），建议专项培训79款SKU知识和搭配推荐。';
  const zeroLt=gs.filter(g=>g.m_lt===0&&g.m_sales>0);
  if(zeroLt.length>0)o+=zeroLt.map(g=>g.name).join('、')+'本月零长尾款销售，需重点跟进79款SKU推荐。';
  if(!o)o='导购长尾推荐整体较好（门店人均¥'+Math.round(ltAvg)+'），保持79款SKU的日常陈列和推荐习惯。';
  return conc(c,o);
}
"""

# 替换硬编码日期为动态值
HTML = HTML.replace('7月门店月度交互分析', f'{_month_title}门店月度交互分析')
HTML = HTML.replace('数据周期：2026年7月1日-7月18日 | 月环比=上月同时间区间(6月1-18日)', f'数据周期：{_date_range} | 月环比=上月同时间区间({_prev_range})')
HTML = HTML.replace('月环比=vs 6月1-18日同期', f'月环比=vs {_prev_range}同期')

with open('os.path.join(OUTPUTS_DIR, '')', 'w', encoding='utf-8') as f:
    f.write(HTML)

import os
sz = os.path.getsize('os.path.join(OUTPUTS_DIR, '')')
print(f"Part 1 written. File size: {sz:,} bytes")
