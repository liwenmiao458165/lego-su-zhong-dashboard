#!/usr/bin/env python3
"""Part 3: renderComparison + selectStore + init + closing tags"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import os
import sys
import glob
from calendar import monthrange
import pandas as pd

# 动态检测数据日期
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
_prev_month = REPORT_MONTH - 1
if _prev_month == 0:
    _prev_month = 12
_prev_last_day = monthrange(max_date.year if REPORT_MONTH > 1 else max_date.year - 1, _prev_month)[1]
_prev_end_day = min(REPORT_DAY, _prev_last_day)
_prev_range = f'{_prev_month}月1-{_prev_end_day}日'
_curr_range = f'{REPORT_MONTH}月1-{REPORT_DAY}日'
print(f"  monthly p3: {_curr_range} vs {_prev_range}")

HTML = r"""
function renderComparison(){
  const r=DATA['_region'];
  const totAmt=regVal(x=>x.month_amt);
  const totTarget=regVal(x=>x.target);
  const totRate=totTarget>0?totAmt/totTarget*100:0;
  const totCnt=regVal(x=>x.month_cnt);
  const totWos=regVal(x=>x.month_wos);
  const totJl=regVal(x=>x.month_jl);
  const totLt=regVal(x=>x.month_lt);
  const totLtQty=regVal(x=>x.month_lt_qty);
  const totNewAmt=regVal(x=>(x.new_amt||0));
  const totNewRate=r.new_rate||0;
  const sReg=DATA['_sample_region']||{month_amt:0,month_qty:0};
  let html='<div class="kpi-row">'+
    '<div class="kpi-card"><div class="label">区域月累计</div><div class="value" style="color:#1e293b">'+fmtAmt(totAmt)+'</div><div class="sub">月目标'+fmtAmt(totTarget)+' | <span style="color:'+(totRate>=50?'#16a34a':'#dc2626')+';font-weight:700">完成率'+totRate.toFixed(1)+'%</span></div></div>'+
    '<div class="kpi-card"><div class="label">区域月环比</div><div class="value '+(isNA(r.mom)?'':colorPct(r.mom))+'">'+(isNA(r.mom)?'N/A':(r.mom>=0?'+':'')+r.mom.toFixed(1)+'%')+'</div><div class="sub">vs上月同期</div></div>'+
    '<div class="kpi-card"><div class="label">区域月同比</div><div class="value '+(isNA(r.yoy)?'':colorPct(r.yoy))+'">'+(isNA(r.yoy)?'N/A':fmtPct(r.yoy))+'</div><div class="sub">vs去年同期</div></div>'+
    '<div class="kpi-card"><div class="label">区域月连带率</div><div class="value blue">'+r.month_jd.toFixed(2)+'</div><div class="sub">'+totCnt+'笔</div></div>'+
    '<div class="kpi-card"><div class="label">区域月客单价</div><div class="value blue">¥'+Math.round(r.month_atv)+'</div><div class="sub">7家门店</div></div>'+
    '<div class="kpi-card"><div class="label">区域长尾款月累计</div><div class="value purple">'+fmtAmt(totLt)+'</div><div class="sub">'+totLtQty+'件 | 占比'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">区域样品月累计</div><div class="value" style="color:#9333ea">'+fmtAmt(sReg.month_amt||0)+'</div><div class="sub">'+(sReg.month_qty||0)+'件（独立口径，不计入门店销售额）</div></div>'+
    '<div class="kpi-card"><div class="label">区域新品月累计</div><div class="value" style="color:#db2777">'+fmtAmt(r.new_amt||0)+'</div><div class="sub">'+(r.new_qty||0)+'件 | 占比'+(r.new_rate||0).toFixed(1)+'%</div></div>'+
    '</div>';

  html+=conc('区域本月月累计'+fmtAmt(totAmt)+'，月目标'+fmtAmt(totTarget)+'，完成率'+totRate.toFixed(1)+'%；月同比'+fmtPct(r.yoy)+'，月环比'+(isNA(r.mom)?'N/A':(r.mom>=0?'+':'')+r.mom.toFixed(1)+'%。'),function(){const rates=STORE_ORDER.map(s=>({s,rate:DATA[s].month_rate}));const avg=rates.reduce((a,x)=>a+x.rate,0)/rates.length;const below=rates.filter(x=>x.rate<avg).sort((a,b)=>a.rate-b.rate);const wy=STORE_ORDER.map(s=>({s,v:DATA[s].yoy})).filter(x=>!isNA(x.v)).sort((a,b)=>a.v-b.v)[0];let o='';if(below.length>0)o+='完成率低于区域均值（'+avg.toFixed(1)+'%）的门店：'+below.map(x=>x.s+'('+x.rate.toFixed(1)+'%)').join('、')+'，是区域达标主要短板，建议城市经理驻店帮扶并复盘目标拆解。';if(wy)o+=' '+wy.s+'月同比最低（'+fmtPct(wy.v)+'），需重点排查品类缺货与竞品分流。';return o||'各门店完成率均衡，区域整体达标进度健康。';}());

  // Pre-compute ltSampleData for use in conc() and chart
  const ltSampleData=STORE_ORDER.map(s=>{const d=DATA[s];const sampD=(DATA['_sample']||{})[s]||{};return{s,ltAmt:d.month_lt||0,ltQty:d.month_lt_qty||0,ltPct:d.month_amt>0?(d.month_lt||0)/d.month_amt*100:0,sampAmt:sampD.month_amt||0,sampQty:sampD.month_qty||0};}).sort((a,b)=>b.ltAmt-a.ltAmt);

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📈</span> 各门店月累计销售&完成率对比</h3><div class="chart-wrap"><canvas id="cmpAmt"></canvas></div>'+
    conc(function(){const sr=STORE_ORDER.map(s=>{const d=DATA[s];return{s,amt:d.month_amt,target:d.target,rate:d.month_rate};}).sort((a,b)=>b.rate-a.rate);return '完成率排名：'+sr.map(x=>x.s+' '+x.rate.toFixed(1)+'%').join('、')+'，区域合计'+fmtAmt(totAmt)+'。';}(),function(){const sr=STORE_ORDER.map(s=>{const d=DATA[s];return{s,amt:d.month_amt,target:d.target,rate:d.month_rate};}).sort((a,b)=>b.rate-a.rate);const avgRate=sr.reduce((a,x)=>a+x.rate,0)/sr.length;const below=sr.filter(x=>x.rate<avgRate);const above=sr.filter(x=>x.rate>=avgRate);let o='';if(below.length>0)o+='低于区域均值（'+avgRate.toFixed(1)+'%）的门店：'+below.map(x=>x.s+'('+x.rate.toFixed(1)+'%,差'+(avgRate-x.rate).toFixed(1)+'pct)').join('、')+'，需重点跟进目标达成。';if(above.length>0)o+=above.map(x=>x.s+'('+x.rate.toFixed(1)+'%)').join('、')+'完成率领先，可提炼经验推广。';return o||'各门店完成率较为均衡。';}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">📊</span> 各门店月同比对比</h3><div class="chart-wrap"><canvas id="cmpYoy"></canvas></div>'+
    conc(function(){const s=STORE_ORDER.map(x=>({s:x,v:DATA[x].yoy})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v);if(s.length===0)return '暂无同比数据。';return '月同比最高'+s[0].s+'('+fmtPct(s[0].v)+')，最低'+s[s.length-1].s+'('+fmtPct(s[s.length-1].v)+')，区域'+fmtPct(r.yoy)+'。';}(),function(){const s=STORE_ORDER.map(x=>({s:x,v:DATA[x].yoy})).filter(x=>!isNA(x.v));const neg=s.filter(x=>x.v<0);if(neg.length>0)return neg.map(x=>x.s+'('+fmtPct(x.v)+')').join('、')+'同比下滑，需关注其品类结构调整和竞品分流影响。';if(r.yoy<0)return '区域整体同比下滑，检视去年同期爆款是否缺货。';return '区域同比表现良好，保持品类结构和渠道策略。';}())+'</div></div>';

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">🎯</span> 各门店客单价&连带率对比</h3><div class="chart-wrap"><canvas id="cmpAtvJd"></canvas></div>'+
    conc(function(){const atv=STORE_ORDER.map(s=>({s,v:DATA[s].month_atv})).sort((a,b)=>b.v-a.v);const jd=STORE_ORDER.map(s=>({s,v:DATA[s].month_jd})).sort((a,b)=>b.v-a.v);return '客单价最高'+atv[0].s+'(¥'+Math.round(atv[0].v)+')，连带率最高'+jd[0].s+'('+jd[0].v.toFixed(2)+')。';}(),function(){const jdV=STORE_ORDER.map(s=>({s,v:DATA[s].month_jd})).sort((a,b)=>b.v-a.v);const atvV=STORE_ORDER.map(s=>({s,v:DATA[s].month_atv})).sort((a,b)=>b.v-a.v);const ratv=r.month_atv,rjd=r.month_jd;const atvBelow=atvV.filter(x=>x.v<ratv);const jdBelow=jdV.filter(x=>x.v<rjd);let t='';if(atvBelow.length>0)t+=atvBelow.map(x=>x.s+'(¥'+Math.round(x.v)+',差¥'+Math.round(ratv-x.v)+')').join('、')+'客单价低于区域均值¥'+Math.round(ratv)+'，建议引导高端产品推荐。';if(jdBelow.length>0)t+=jdBelow.map(x=>x.s+'('+x.v.toFixed(2)+',差'+(rjd-x.v).toFixed(2)+')').join('、')+'连带率低于区域均值'+rjd.toFixed(2)+'，建议强化搭配推荐培训。';return t||'各门店客单价和连带率均达区域均值，保持现有推荐策略。';}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">📉</span> 各门店月环比对比</h3><div class="chart-wrap"><canvas id="cmpMom"></canvas></div>'+
    conc(function(){const s=STORE_ORDER.map(x=>({s:x,v:DATA[x].mom})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v);if(s.length===0)return '暂无环比数据。';return '月环比最佳'+s[0].s+'('+(s[0].v>=0?'+':'')+s[0].v.toFixed(1)+'%)，区域'+(isNA(r.mom)?'N/A':(r.mom>=0?'+':'')+r.mom.toFixed(1)+'%')+'。';}(),function(){const s=STORE_ORDER.map(x=>({s:x,v:DATA[x].mom})).filter(x=>!isNA(x.v));const reMom=isNA(r.mom)?0:r.mom;const below=s.filter(x=>x.v<reMom);const neg=s.filter(x=>x.v<-8);let o='';if(below.length>0&&!isNA(r.mom))o+='低于区域月环比（'+fmtPct(r.mom)+'）的门店：'+below.map(x=>x.s+'('+fmtPct(x.v)+')').join('、')+'，需复盘客流转化变化。';if(neg.length>0)o+=neg.map(x=>x.s+'('+x.v.toFixed(1)+'%)').join('、')+'环比下滑较大，需关注竞品及促销影响。';return o||'各门店月环比基本平稳，关注下半月冲量节奏。';}())+'</div></div>';


  // 门店月成交率 & 长尾款 对比（半宽并排，第三排）
  const cvo=(DATA['_conversion']||{});
  const convData=STORE_ORDER.map(s=>({s,v:cvo[s]?cvo[s].rate:0,flow:cvo[s]?cvo[s].flow:0,cnt:cvo[s]?cvo[s].cnt:0})).sort((a,b)=>b.v-a.v);
  const rConv=cvo._region||{};
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📦</span> 各门店长尾款 & 样品销售月累计对比（标注占比）</h3><div class="chart-wrap"><canvas id="cmpLtSample"></canvas></div>'+
    conc('区域长尾款月累计'+fmtAmt(totLt)+'（'+totLtQty+'件，占比'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%）。各门店长尾款占比：'+ltSampleData.filter(x=>x.ltAmt>0).map(x=>x.s+x.ltPct.toFixed(1)+'%').join('、')+'。样品月累计'+fmtAmt(sReg.month_amt||0)+'（'+(sReg.month_qty||0)+'件）。',function(){const lpct=totAmt>0?totLt/totAmt*100:0;if(lpct<8)return '区域长尾款占比偏低（'+lpct.toFixed(1)+'%），建议统一部署79款SKU推荐话术培训和收银台陈列优化。';if(lpct>15)return '区域长尾款占比优秀（'+lpct.toFixed(1)+'%），保持当前关联推荐策略。';return '长尾款占比适中，关注各门店间均衡性。';}())+
    '</div>'+
    '<div class="chart-card"><h3><span class="icon">🎯</span> 各门店月成交率对比（成交笔数/客流量，本月累计）</h3><div class="chart-wrap"><canvas id="cmpConv"></canvas></div>'+
    conc('区域月成交率'+(rConv.rate?rConv.rate.toFixed(2):'N/A')+'%（月成交'+(rConv.cnt||0)+'笔 ÷ 月客流'+(rConv.flow||0)+'人次）。各店：'+convData.map(x=>x.s+(x.v===null?'N/A':x.v.toFixed(2)+'%')).join('、')+'。',function(){const valid=convData.filter(x=>x.v!==null);const avg=valid.length?valid.reduce((a,x)=>a+x.v,0)/valid.length:0;const below=valid.filter(x=>x.v<avg);let o='';if(below.length>0)o+='低于区域均值（'+avg.toFixed(2)+'%）的门店：'+below.map(x=>x.s+'('+x.v.toFixed(2)+'%,差'+(avg-x.v).toFixed(2)+'pct)').join('、')+'，客流未有效转化为笔数，需提升进店转化（橱窗陈列/试玩体验/导购拦截话术）。';else o+='各门店成交率较为均衡。';if(valid.length){const best=valid[0],worst=valid[valid.length-1];o+='最高'+best.s+'('+best.v.toFixed(2)+'%)与最低'+worst.s+'('+worst.v.toFixed(2)+'%)相差'+(best.v-worst.v).toFixed(2)+'pct，差距明显可对标提升。';}return o;}())+
    '</div></div>';

  // WOS & JL comparison + Series proportion
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">⚡</span> 各门店WOS & 即时零售月累计对比</h3><div class="chart-wrap"><canvas id="cmpWosJl"></canvas></div>'+
    conc('WOS月累计'+fmtAmt(totWos)+'（占比'+(totAmt>0?(totWos/totAmt*100).toFixed(1):0)+'%），即时零售'+fmtAmt(totJl)+'（占比'+(totAmt>0?(totJl/totAmt*100).toFixed(1):0)+'%）。',function(){const wpct=totAmt>0?totWos/totAmt*100:0,jpct=totAmt>0?totJl/totAmt*100:0;const lowW=STORE_ORDER.filter(s=>DATA[s].month_amt>0&&DATA[s].month_wos/DATA[s].month_amt*100<3);let t='';if(wpct+jpct<10)t='WOS+即时零售合计占比不足10%，建议加大线上渠道投放力度，特别是美团/饿了么专属活动。';else if(wpct+jpct>20)t='线上渠道占比超过20%，多渠道运营成熟。';else t='线上渠道占比中等，适度增加即时零售平台曝光和WOS直播频次。';if(lowW.length>0)t+=' '+lowW.join('、')+'WOS占比偏低，需重点提升。';return t;}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">📊</span> 区域产品系列结构（月累计）</h3><div class="chart-wrap"><canvas id="cmpRegSeries"></canvas></div>'+
    conc(function(){const rs=(DATA['_series']||{})['_region']||{};const ents=Object.entries(rs).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '暂无系列数据。';const top10=ents.slice(0,10);const t3pct=totAmt>0?ents.slice(0,3).reduce((a,[,v])=>a+v,0)/totAmt*100:0;const botAmt=ents.reduce((a,[k,v])=>k.toUpperCase().includes('BOTANICAL')?a+v:a,0);const botPct=totAmt>0?botAmt/totAmt*100:0;return '区域TOP10系列：'+top10.map(([k,v])=>k+'('+fmtAmt(v)+',占比'+(totAmt>0?(v/totAmt*100).toFixed(1):0)+'%)').join('、')+'；TOP3合计占比'+t3pct.toFixed(1)+'%。BOTANICALS植物系列占比'+botPct.toFixed(1)+'%（指标4.1%）。';}(),function(){const rs=(DATA['_series']||{})['_region']||{};const ents=Object.entries(rs).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '';const top3=ents.slice(0,3);const top5=ents.slice(0,5);const t3pct=totAmt>0?top3.reduce((a,[,v])=>a+v,0)/totAmt*100:0;let o='';if(t3pct>65)o+='区域TOP3系列占比过高（'+t3pct.toFixed(1)+'%），建议丰富中长尾系列陈列和推荐。';else if(t3pct<30)o+='系列分布过于分散，建议集中打造1-2个主推系列。';else o+='系列结构均衡，保持当前推荐策略。';const focusTop3=['CITY','TECHNIC','NINJAGO'];const missingTop3=focusTop3.filter(n=>!top3.some(([k])=>k.toUpperCase().includes(n)));if(missingTop3.length>0)o+='重点系列'+missingTop3.join('、')+'未进TOP3，建议加强陈列和推荐。';const inTop3=focusTop3.filter(n=>top3.some(([k])=>k.toUpperCase().includes(n)));if(inTop3.length>0)o+='重点系列'+inTop3.join('、')+'在TOP3表现良好。';if(!top5.some(([k])=>k.toUpperCase().includes('MINECRAFT')))o+='MINECRAFT未进TOP5，需关注沙盒游戏IP粉丝群体引流。';const botAmt=ents.reduce((a,[k,v])=>k.toUpperCase().includes('BOTANICAL')?a+v:a,0);const botPct=totAmt>0?botAmt/totAmt*100:0;if(botAmt>0){o+='BOTANICALS植物系列占比'+botPct.toFixed(1)+'%（指标4.1%）';if(botPct<4.1)o+='，低于指标，需加强植物系列陈列和推荐。';else o+='，达到指标。';}else{o+='BOTANICALS植物系列无销售，低于指标4.1%，建议关注植物系列铺货和推荐。';}return o;}())+'</div></div>';



  // Regional TOP10 products
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">🏆</span> 区域TOP10产品月销排名（含件数）</h3><div class="chart-wrap"><canvas id="cmpRegProducts"></canvas></div>'+
    conc(function(){const rp=(DATA['_products']||{})['_region']||{};const ents=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,3);if(ents.length===0)return '暂无产品数据。';return '区域TOP3产品：'+ents.map(([n,a,q])=>n+'('+fmtAmt(a)+','+q+'件)').join('、')+'。';}(),function(){const rp=(DATA['_products']||{})['_region']||{};const ents=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '';const t1=ents[0];const t1pct=totAmt>0?t1[1]/totAmt*100:0;let o='';if(t1pct>20)o+='TOP1产品占比'+t1pct.toFixed(1)+'%，依赖度高，需关注库存和缺货风险。';else o+='区域产品结构多元，无过度依赖单品的风险。';const newProds=DATA['_new_products']||[];if(newProds.length>0){const top10Names=ents.slice(0,10).map(([k])=>k);const newInTop=newProds.filter(np=>top10Names.some(tn=>tn.includes(np)||np.includes(tn)));if(newInTop.length>0)o+='当月新品「'+newInTop.join('、')+'」进入区域TOP10，表现突出，建议加大陈列推荐。';else o+='当月新品无进入区域TOP10，需关注新品陈列位置和推荐话术。';}return o;}())+'</div></div>';

  // 新品模块：各店新品占比对比 + 区域新品TOP
  const newShareData=STORE_ORDER.map(s=>{const d=DATA[s];return{s,v:d.new_rate||0,amt:d.new_amt||0};}).sort((a,b)=>b.v-a.v);
  const rNewRate=r.new_rate||0;
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">🆕</span> 各门店新品销售占比对比（上市日期=本月）</h3><div class="chart-wrap"><canvas id="cmpNewShare"></canvas></div>'+
    conc(function(){const rs=STORE_ORDER.map(s=>({s,v:DATA[s].new_rate||0})).sort((a,b)=>b.v-a.v);if(rs[0].v===0)return '本月暂无新品销售记录。';return '新品占比最高'+rs[0].s+'('+rs[0].v.toFixed(1)+'%)。';}(),function(){const rs=STORE_ORDER.map(s=>({s,v:DATA[s].new_rate||0,amt:DATA[s].new_amt||0}));const zero=rs.filter(x=>x.amt===0);const below=rs.filter(x=>x.v<rNewRate&&x.amt>0);let o='';if(zero.length>0)o+=zero.map(x=>x.s).join('、')+'本月无新品销售，需紧急排查新品铺货与陈列。';if(below.length>0)o+=below.map(x=>x.s+'('+x.v.toFixed(1)+'%)').join('、')+'新品占比低于区域('+rNewRate.toFixed(1)+'%)。';if(!o)o='各门店均有新品销售且占比达区域水平，动销良好。';return o;}())+
    '</div>'+
    '<div class="chart-card"><h3><span class="icon">🆕</span> 区域新品TOP销售排名（上市日期=本月）</h3><div class="chart-wrap"><canvas id="cmpNewProducts"></canvas></div>'+concNew('区域对比')+'</div></div>';

  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">📋</span> 区域月维度汇总表</h3><table class="summary-table" id="cmpTable"></table></div></div>';

  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">👩</span> 全区域导购月度完成率排名</h3><div class="chart-wrap tall" style="height:600px;"><canvas id="cmpGuides"></canvas></div>'+
    conc(function(){const all=[];STORE_ORDER.forEach(s=>{(DATA['_guides']||{})[s]?.forEach(g=>all.push(g));});all.sort((a,b)=>b.m_rate-a.m_rate);if(all.length===0)return '暂无导购数据。';const t=all[0];return '完成率第一'+t.name+'('+t.store+')'+t.m_rate.toFixed(1)+'%，月销售'+fmtAmt(t.m_sales)+'。';}(),function(){const all=[];STORE_ORDER.forEach(s=>{(DATA['_guides']||{})[s]?.forEach(g=>all.push(g));});all.sort((a,b)=>b.m_rate-a.m_rate);if(all.length===0)return '';const regAvg=all.filter(g=>g.m_sales>0).reduce((a,g)=>a+g.m_sales,0)/all.filter(g=>g.m_sales>0).length;const below=all.filter(g=>g.m_sales>0&&g.m_sales<regAvg*0.7);const high=all.filter(g=>g.m_rate>=80);const low=all.filter(g=>g.m_rate<30&&g.m_sales>0);let o='';if(below.length>0)o+='低于区域均值¥'+Math.round(regAvg)+'的导购：'+below.map(g=>g.name+'('+g.store+',¥'+Math.round(g.m_sales)+')').join('、')+'，需重点跟进销售辅导和排班优化。';if(high.length>0)o+=high.map(g=>g.name+'('+g.store+')').join('、')+'完成率优秀(≥80%)，建议经验分享。';if(low.length>0)o+=low.map(g=>g.name+'('+g.store+')').join('、')+'完成率不足30%，需一对一辅导。';return o||'各导购完成率较为均衡（区域人均¥'+Math.round(regAvg)+'）。';}())+'</div></div>';

  document.getElementById('content').innerHTML=html;
  Object.values(charts).forEach(c=>c&&c.destroy());charts={};

  // Chart 1: Monthly amount + completion rate (sorted by rate, dual axis)
  const storeRates=STORE_ORDER.map(s=>{const d=DATA[s];return{s,v:d.month_amt,target:d.target,rate:d.month_rate};}).sort((a,b)=>b.rate-a.rate);
  charts.cmpAmt=new Chart(document.getElementById('cmpAmt'),{
    data:{labels:storeRates.map(x=>x.s+'\n'+x.rate.toFixed(1)+'%'),datasets:[
      {type:'bar',label:'月累计',data:storeRates.map(x=>x.v),backgroundColor:storeRates.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:storeRates.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6,yAxisID:'y'},
      {type:'line',label:'完成率',data:storeRates.map(x=>x.rate),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:5,pointBackgroundColor:'#dc2626',tension:.3,yAxisID:'y1'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const x=storeRates[ctx.dataIndex];if(ctx.dataset.type==='line')return '完成率: '+x.rate.toFixed(1)+'%';return '月销售: '+fmtAmt(x.v)+' | 月目标: '+fmtAmt(x.target)+' | 完成率: '+x.rate.toFixed(1)+'%';}}}},
      scales:{y:{position:'left',title:{display:true,text:'月累计(¥)'},ticks:{callback:v=>fmtAmt(v)}},y1:{position:'right',title:{display:true,text:'完成率(%)'},grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}}
  });

  // Chart 2: YoY
  const yoyData=STORE_ORDER.map(s=>({s,v:DATA[s].yoy}));
  charts.cmpYoy=new Chart(document.getElementById('cmpYoy'),{
    type:'bar',
    data:{labels:yoyData.map(x=>x.s),datasets:[{label:'月同比',data:yoyData.map(x=>x.v),backgroundColor:yoyData.map(x=>isNA(x.v)?'#e2e8f0':(x.v>=0?'#16a34acc':'#dc2626cc')),borderColor:yoyData.map(x=>isNA(x.v)?'#94a3b8':(x.v>=0?'#16a34a':'#dc2626')),borderWidth:1,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>isNA(ctx.raw)?'N/A':'月同比: '+(ctx.raw>=0?'+':'')+ctx.raw.toFixed(1)+'%'}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}}
  });

  // Chart 3: ATV & JD (dual axis)
  charts.cmpAtvJd=new Chart(document.getElementById('cmpAtvJd'),{
    data:{labels:STORE_ORDER,datasets:[
      {type:'bar',label:'客单价',data:STORE_ORDER.map(s=>DATA[s].month_atv),backgroundColor:'#0ea5e9cc',borderColor:'#0ea5e9',borderWidth:1,borderRadius:4,yAxisID:'y'},
      {type:'line',label:'连带率',data:STORE_ORDER.map(s=>DATA[s].month_jd),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#dc2626',tension:.3,yAxisID:'y1'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label.includes('客单'))return '客单价: ¥'+Math.round(ctx.raw);return '连带率: '+ctx.raw.toFixed(2)}}}},
      scales:{y:{position:'left',title:{display:true,text:'客单价(¥)'},ticks:{callback:v=>'¥'+v}},y1:{position:'right',title:{display:true,text:'连带率'},grid:{drawOnChartArea:false},ticks:{callback:v=>v.toFixed(1)}}}}
  });

  // Chart 4: MoM
  const momData=STORE_ORDER.map(s=>({s,v:DATA[s].mom}));
  charts.cmpMom=new Chart(document.getElementById('cmpMom'),{
    type:'bar',
    data:{labels:momData.map(x=>x.s),datasets:[
      {label:'6月1-18日',data:momData.map(x=>DATA[x.s].mom_prev||0),backgroundColor:'#94a3b880',borderColor:'#64748b',borderWidth:1,borderRadius:4},
      {label:'7月1-18日',data:momData.map(x=>DATA[x.s].month_amt||0),backgroundColor:momData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:momData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmtAmt(ctx.raw)}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 5: WOS & JL comparison
  charts.cmpWosJl=new Chart(document.getElementById('cmpWosJl'),{
    type:'bar',
    data:{labels:STORE_ORDER,datasets:[
      {label:'WOS',data:STORE_ORDER.map(s=>DATA[s].month_wos),backgroundColor:'#0284c7cc',borderColor:'#0284c7',borderWidth:1,borderRadius:4},
      {label:'即时零售',data:STORE_ORDER.map(s=>DATA[s].month_jl),backgroundColor:'#ea580ccc',borderColor:'#ea580c',borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmtAmt(ctx.raw)}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: Combined long-tail & sample comparison
  charts.cmpLtSample=new Chart(document.getElementById('cmpLtSample'),{
    type:'bar',
    data:{labels:ltSampleData.map(x=>x.s),datasets:[
      {label:'长尾款',data:ltSampleData.map(x=>x.ltAmt),backgroundColor:'#ea580ccc',borderColor:'#ea580c',borderWidth:1,borderRadius:4},
      {label:'样品',data:ltSampleData.map(x=>x.sampAmt),backgroundColor:'#9333eacc',borderColor:'#9333ea',borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const x=ltSampleData[ctx.dataIndex];if(ctx.datasetIndex===0)return '长尾款: '+fmtAmt(x.ltAmt)+' ('+x.ltQty+'件, 占比'+x.ltPct.toFixed(1)+'%)';return '样品: '+fmtAmt(x.sampAmt)+' ('+x.sampQty+'件)';}}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 6: Regional series doughnut
  const rs=(DATA['_series']||{})['_region']||{};
  const rse=Object.entries(rs).sort((a,b)=>b[1]-a[1]);
  const rtop10=rse.slice(0,10),roth=rse.slice(10).reduce((a,[,v])=>a+v,0);
  charts.cmpRegSeries=new Chart(document.getElementById('cmpRegSeries'),{
    type:'doughnut',
    data:{labels:rtop10.map(([k])=>k).concat(roth>0?['其他']:[]),datasets:[{data:rtop10.map(([,v])=>v).concat(roth>0?[roth]:[]),backgroundColor:SC.slice(0,rtop10.length+(roth>0?1:0)),borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},padding:6}},tooltip:{callbacks:{label:ctx=>{const pct=totAmt>0?(ctx.raw/totAmt*100).toFixed(1):0;return ctx.label+': '+fmtAmt(ctx.raw)+' ('+pct+'%)'}}}},cutout:'55%'}
  });

  // Chart 7: Regional TOP10 products
  const rp=(DATA['_products']||{})['_region']||{};
  const rpEnts=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  charts.cmpRegProducts=new Chart(document.getElementById('cmpRegProducts'),{
    type:'bar',
    data:{labels:rpEnts.map(([k])=>k),datasets:[{label:'月销售',data:rpEnts.map(([k,a,q])=>a),backgroundColor:SC.slice(0,rpEnts.length).map(c=>c+'cc'),borderColor:SC.slice(0,rpEnts.length),borderWidth:1,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const idx=ctx.dataIndex;const [n,a,q]=rpEnts[idx];const pct=totAmt>0?(a/totAmt*100).toFixed(1):0;return n+': '+fmtAmt(a)+' ('+pct+'%, '+q+'件)'}}}},
      scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: 门店月成交率
  charts.cmpConv=new Chart(document.getElementById('cmpConv'),{
    type:'bar',
    data:{labels:convData.map(x=>x.s),datasets:[{label:'月成交率',data:convData.map(x=>x.v),backgroundColor:convData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:convData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const x=convData[ctx.dataIndex];return '月成交率: '+(x.v===null?'缺客流':x.v.toFixed(2)+'%')+' | 笔数'+x.cnt+' / 客流'+x.flow;}}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}}
  });

  // Chart: New-product share by store (上市日期=本月)
  charts.cmpNewShare=new Chart(document.getElementById('cmpNewShare'),{
    type:'bar',
    data:{labels:newShareData.map(x=>x.s),datasets:[
      {label:'新品占比',data:newShareData.map(x=>x.v),backgroundColor:newShareData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:newShareData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6},
      {label:'区域均值',type:'line',data:newShareData.map(x=>rNewRate),borderColor:'#1e293b',backgroundColor:'transparent',borderWidth:2,pointRadius:0,borderDash:[6,4]}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const x=newShareData[ctx.dataIndex];if(ctx.dataset.label==='新品占比')return '新品占比: '+x.v.toFixed(1)+'% | 新品额: '+fmtAmt(x.amt);return '区域均值: '+ctx.raw.toFixed(1)+'%';}}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}}
  });

  // Chart: Regional new-product TOP
  const npRegDetail=(DATA['_new_products_detail']||{})['_region']||[];
  if(npRegDetail.length>0){
    charts.cmpNewProducts=new Chart(document.getElementById('cmpNewProducts'),{
      type:'bar',
      data:{labels:npRegDetail.map(x=>x.name),datasets:[{label:'新品销售',data:npRegDetail.map(x=>x.amt),backgroundColor:SC.slice(0,npRegDetail.length).map(c=>c+'cc'),borderColor:SC.slice(0,npRegDetail.length),borderWidth:1,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const x=npRegDetail[ctx.dataIndex];const pct=totAmt>0?(x.amt/totAmt*100).toFixed(1):0;return x.name+': '+fmtAmt(x.amt)+' ('+pct+'%, '+x.qty+'件)';}}}},
        scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
    });
  }else{
    document.getElementById('cmpNewProducts').parentNode.innerHTML='<div style="padding:40px;text-align:center;color:#94a3b8">本月无新品销售记录</div>';
  }

  // Summary table
  let tHtml='<thead><tr><th>门店</th><th>月累计</th><th>月目标</th><th>月完成率</th><th>月同比</th><th>月环比</th><th>笔数</th><th>客单价</th><th>连带率</th><th>成交率</th><th>WOS</th><th>即时零售</th><th>长尾款</th><th>长尾件数</th><th>长尾占比</th><th>新品额</th><th>新品占比</th><th>样品</th><th>样品件数</th></tr></thead><tbody>';
  // 区域均值（区域汇总表低于均值标红）
  const regMonthAmtAvg=totAmt/STORE_ORDER.length;
  const regMonthCntAvg=totCnt/STORE_ORDER.length;
  const regMonthATV=r.month_atv;
  const regMonthJD=r.month_jd;
  const regMonthLtPctAvg=STORE_ORDER.reduce((a,s)=>a+(DATA[s].month_amt>0?(DATA[s].month_lt/DATA[s].month_amt*100):0),0)/STORE_ORDER.length;
  const regMonthNewPctAvg=STORE_ORDER.reduce((a,s)=>a+(DATA[s].month_amt>0?((DATA[s].new_amt||0)/DATA[s].month_amt*100):0),0)/STORE_ORDER.length;
  STORE_ORDER.forEach(s=>{
    const d=DATA[s];
    const lpct=d.month_amt>0?(d.month_lt/d.month_amt*100):0;
    const sd=((DATA['_sample']||{})[s])||{};
    const yb=isNA(d.yoy)?'<span class="badge gray">N/A</span>':(d.yoy>=0?'<span class="badge green">+'+d.yoy.toFixed(1)+'%</span>':'<span class="badge red">'+d.yoy.toFixed(1)+'%</span>');
    const mb=isNA(d.mom)?'<span class="badge gray">N/A</span>':(d.mom>=0?'<span class="badge green">+'+d.mom.toFixed(1)+'%</span>':'<span class="badge red">'+d.mom.toFixed(1)+'%</span>');
    tHtml+='<tr><td class="store-name" style="color:'+STORE_COLORS[s]+'">'+s+'</td><td style="font-weight:700'+(d.month_amt<regMonthAmtAvg?';color:#dc2626':'')+'">'+fmtAmt(d.month_amt)+'</td><td>'+fmtAmt(d.target)+'</td><td style="color:'+(d.month_rate>=50?'#16a34a':'#dc2626')+';font-weight:700">'+d.month_rate.toFixed(1)+'%</td><td>'+yb+'</td><td>'+mb+'</td><td'+(d.month_cnt<regMonthCntAvg?' style="color:#dc2626;font-weight:700"':'')+'>'+d.month_cnt+'</td><td'+(d.month_atv<regMonthATV?' style="color:#dc2626;font-weight:700"':'')+'>¥'+Math.round(d.month_atv)+'</td><td'+(d.month_jd<regMonthJD?' style="color:#dc2626;font-weight:700"':'')+'>'+d.month_jd.toFixed(2)+'</td><td style="color:'+((cvo[s]&&cvo[s].rate>=rConv.rate)?'#16a34a':'#dc2626')+';font-weight:700">'+((cvo[s]&&cvo[s].rate)?cvo[s].rate.toFixed(2):'N/A')+'%</td><td>'+fmtAmt(d.month_wos)+'</td><td>'+fmtAmt(d.month_jl)+'</td><td>'+fmtAmt(d.month_lt)+'</td><td>'+(d.month_lt_qty||0)+'件</td><td style="color:'+(lpct<regMonthLtPctAvg?'#dc2626':'#16a34a')+';font-weight:700">'+lpct.toFixed(1)+'%</td>'+
'<td'+(d.new_amt<regMonthAmtAvg?' style="color:#dc2626;font-weight:700"':'')+'>'+fmtAmt(d.new_amt||0)+'</td>'+
'<td style="color:'+((d.new_rate||0)<regMonthNewPctAvg?'#dc2626':'#16a34a')+';font-weight:700">'+((d.new_rate||0)).toFixed(1)+'%</td>'+
'<td style="color:#9333ea">'+fmtAmt(sd.month_amt||0)+'</td><td>'+(sd.month_qty||0)+'件</td></tr>';
  });
  tHtml+='<tr style="background:#7c3aed10;font-weight:700"><td class="store-name">区域合计</td><td>'+fmtAmt(totAmt)+'</td><td>'+fmtAmt(totTarget)+'</td><td style="color:'+(totRate>=50?'#16a34a':'#dc2626')+'">'+totRate.toFixed(1)+'%</td><td>'+(isNA(r.yoy)?'<span class="badge gray">N/A</span>':(r.yoy>=0?'<span class="badge green">+'+r.yoy.toFixed(1)+'%</span>':'<span class="badge red">'+r.yoy.toFixed(1)+'%</span>'))+'</td><td>'+(isNA(r.mom)?'<span class="badge gray">N/A</span>':(r.mom>=0?'<span class="badge green">+'+r.mom.toFixed(1)+'%</span>':'<span class="badge red">'+r.mom.toFixed(1)+'%</span>'))+'</td><td>'+totCnt+'</td><td>¥'+Math.round(r.month_atv)+'</td><td>'+r.month_jd.toFixed(2)+'</td><td style="color:#16a34a;font-weight:700">'+(rConv.rate?rConv.rate.toFixed(2):'N/A')+'%</td><td>'+fmtAmt(totWos)+'</td><td>'+fmtAmt(totJl)+'</td><td>'+fmtAmt(totLt)+'</td><td>'+totLtQty+'件</td><td style="color:'+(totAmt>0&&(totLt/totAmt*100)>=10?'#16a34a':'#dc2626')+';font-weight:700">'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%</td>'+
'<td>'+fmtAmt(totNewAmt)+'</td>'+
'<td style="color:#16a34a;font-weight:700">'+totNewRate.toFixed(1)+'%</td>'+
'<td style="color:#9333ea">'+fmtAmt(sReg.month_amt||0)+'</td><td>'+(sReg.month_qty||0)+'件</td></tr>';
  tHtml+='</tbody>';
  document.getElementById('cmpTable').innerHTML=tHtml;
  // 区域均值标红说明
  let cmpNote=document.getElementById('cmpTableNote');
  if(cmpNote) cmpNote.remove();
  cmpNote=document.createElement('div');
  cmpNote.id='cmpTableNote';
  cmpNote.style='margin-top:10px;padding:8px 12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;font-size:13px;color:#dc2626;font-weight:700';
  cmpNote.textContent='🔴 月累计/笔数/客单价/连带率/长尾占比/新品占比 低于区域均值已标红（月累计¥'+Math.round(regMonthAmtAvg)+'、笔数'+Math.round(regMonthCntAvg)+'、客单价¥'+Math.round(regMonthATV)+'、连带率'+regMonthJD.toFixed(2)+'、长尾占比'+regMonthLtPctAvg.toFixed(1)+'%、新品占比'+regMonthNewPctAvg.toFixed(1)+'%）';
  document.getElementById('cmpTable').parentNode.appendChild(cmpNote);
  let _cmpConc=document.createElement('div');
  _cmpConc.style.cssText='margin-top:10px';
  const _below=STORE_ORDER.filter(s=>DATA[s].month_amt<regMonthAmtAvg);
  const _dual=STORE_ORDER.filter(s=>!isNA(DATA[s].yoy)&&DATA[s].yoy<0&&!isNA(DATA[s].mom)&&DATA[s].mom<0);
  const _topRate=STORE_ORDER.map(s=>({s,rate:DATA[s].month_rate})).sort((a,b)=>b.rate-a.rate)[0];
  _cmpConc.innerHTML=conc('区域7店中'+_below.length+'家月累计低于区域均值（¥'+Math.round(regMonthAmtAvg)+'）：'+(_below.length?_below.map(s=>s).join('、')+'。':'均超均值。')+_topRate.s+'完成率领先（'+_topRate.rate.toFixed(1)+'%）。',(_dual.length>0?(_dual.map(s=>s).join('、')+'同环比双降，是区域主要拖累，建议集中资源帮扶并复盘品类与客流。'):(_below.length===0?'各门店月累计与同环比表现均衡，区域整体健康。':'')));
  document.getElementById('cmpTable').parentNode.appendChild(_cmpConc);

  // All guides ranking chart — sorted by completion rate
  const allGuides=[];
  STORE_ORDER.forEach(s=>{((DATA['_guides']||{})[s]||[]).forEach(g=>allGuides.push(g));});
  allGuides.sort((a,b)=>b.m_rate-a.m_rate);
  if(allGuides.length>0){
    charts.cmpGuides=new Chart(document.getElementById('cmpGuides'),{
      type:'bar',
      data:{labels:allGuides.map(g=>g.name+'('+g.store+') '+g.m_rate.toFixed(1)+'%'),datasets:[{label:'月度完成率(%)',data:allGuides.map(g=>g.m_rate),backgroundColor:allGuides.map(g=>g.m_rate>=50?STORE_COLORS[g.store]+'cc':'#dc2626cc'),borderColor:allGuides.map(g=>g.m_rate>=50?STORE_COLORS[g.store]:'#dc2626'),borderWidth:1,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const g=allGuides[ctx.dataIndex];return g.name+'('+g.store+'): 完成率'+g.m_rate.toFixed(1)+'% | 月销售'+fmtAmt(g.m_sales)+' | 月目标'+fmtAmt(g.target)+' | 月环比'+(isNA(g.mom)?'N/A':(g.mom>=0?'+':'')+g.mom.toFixed(1)+'%')}}}},
        scales:{x:{title:{display:true,text:'完成率(%)'},ticks:{callback:v=>v+'%'}}}}
    });
  }
}

function selectStore(store){
  currentStore=store;
  document.querySelectorAll('.store-tab').forEach(t=>t.classList.toggle('active',t.dataset.store===store));
  if(store==='区域对比'){renderComparison();}
  else{renderStore(store);}
}

document.addEventListener('DOMContentLoaded',()=>{
  document.getElementById('genTime').textContent=new Date().toLocaleString('zh-CN');
  selectStore(currentStore);
});
</script>
</body>
</html>
"""

# 替换硬编码日期为动态值
HTML = HTML.replace('6月1-18日', _prev_range).replace('7月1-18日', _curr_range)

with open('os.path.join(OUTPUTS_DIR, '')', 'a', encoding='utf-8') as f:
    f.write(HTML)

import os
sz = os.path.getsize('os.path.join(OUTPUTS_DIR, '')')
print(f"Part 3 appended. File size: {sz:,} bytes")
