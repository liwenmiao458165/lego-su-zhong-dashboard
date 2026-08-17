#!/usr/bin/env python3
"""Part 4: renderComparison + selectStore + init + closing tags"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import os
import sys
import glob
import pandas as pd

# 动态检测WK编号
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
WK28_BASE = pd.Timestamp('2026-07-06')
WK_NUM_INT = 28 + (max_date - WK28_BASE).days // 7
WK_PREV_NUM = WK_NUM_INT - 1
print(f"  p4: WK{WK_NUM_INT} (prev WK{WK_PREV_NUM})")

HTML = r"""
function renderComparison(){
  const r=DATA['_region'];
  const cvo=(DATA['_conversion']||{});
  // Region KPI cards
  const totAmt=regVal(x=>x.wk_amt);
  const totW28=regVal(x=>x.wk28_amt||0);
  const totWow=totW28>0?(totAmt-totW28)/totW28*100:null;
  const totCnt=regVal(getWeeklyCnt);
  const totWos=regVal(x=>x.wk_wos);
  const totLt=regVal(x=>x.wk_lt);
  const totLtQty=STORE_ORDER.reduce((sum,s)=>sum+DATA[s].daily.reduce((a,x)=>a+(x.lt_qty||0),0),0);
  const sReg=DATA['_sample_region']||{wk_amt:0,wk_qty:0};
  const npRegAmt=STORE_ORDER.reduce((sum,s)=>sum+(DATA[s].wk_np_amt||0),0);
  const regYoy=r.wk_yoy;
  const regWkTarget=STORE_ORDER.reduce((sum,s)=>sum+DATA[s].daily.reduce((a,x)=>a+(x.target||0),0),0);
  const regWkRate=regWkTarget>0?totAmt/regWkTarget*100:0;
  let html='<div class="kpi-row">'+
    '<div class="kpi-card"><div class="label">区域WK29周累计</div><div class="value" style="color:#1e293b">'+fmtAmt(totAmt)+'</div><div class="sub">周目标'+fmtAmt(regWkTarget)+' | <span style="color:'+(regWkRate>=50?'#16a34a':'#dc2626')+';font-weight:700">完成率'+regWkRate.toFixed(1)+'%</span></div></div>'+
    '<div class="kpi-card"><div class="label">区域周环比</div><div class="value '+(isNA(r.wow)?'':colorPct(r.wow))+'">'+(isNA(r.wow)?'N/A':(r.wow>=0?'+':'')+r.wow.toFixed(1)+'%')+'</div><div class="sub">vs WK28</div></div>'+
    '<div class="kpi-card"><div class="label">区域周同比</div><div class="value '+(isNA(regYoy)?'':colorPct(regYoy))+'">'+(isNA(regYoy)?'N/A':fmtPct(regYoy))+'</div><div class="sub">vs去年同期</div></div>'+
    '<div class="kpi-card"><div class="label">区域周连带率</div><div class="value blue">'+getRegWkJD().toFixed(2)+'</div><div class="sub">'+totCnt+'笔</div></div>'+
    '<div class="kpi-card"><div class="label">区域周客单价</div><div class="value blue">\u00a5'+Math.round(getRegWkATV())+'</div><div class="sub">7家门店</div></div>'+
    '<div class="kpi-card"><div class="label">区域WOS周累计</div><div class="value orange">'+fmtAmt(totWos)+'</div><div class="sub">占比'+(totAmt>0?(totWos/totAmt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">区域长尾款周累计</div><div class="value purple">'+fmtAmt(totLt)+'</div><div class="sub">'+totLtQty+'件 | 占比'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">区域样品周累计</div><div class="value" style="color:#9333ea">'+fmtAmt(sReg.wk_amt||0)+'</div><div class="sub">'+(sReg.wk_qty||0)+'件（独立口径，不计入门店销售额）</div></div>'+
    '<div class="kpi-card"><div class="label">区域新品周累计</div><div class="value" style="color:#db2777">'+fmtAmt(npRegAmt)+'</div><div class="sub">占比'+(totAmt>0?(npRegAmt/totAmt*100).toFixed(1):0)+'%</div></div>'+
    '</div>';

  html+=conc('区域本周周累计'+fmtAmt(totAmt)+'，周目标'+fmtAmt(regWkTarget)+'，完成率'+regWkRate.toFixed(1)+'%；周同比'+fmtPct(regYoy)+'，周环比'+(isNA(r.wow)?'N/A':(r.wow>=0?'+':'')+r.wow.toFixed(1)+'%。'),function(){const rates=STORE_ORDER.map(s=>{const d=DATA[s];const wt=d.daily.reduce((a,x)=>a+(x.target||0),0);return{s,rate:wt>0?d.wk_amt/wt*100:0};});const avg=rates.reduce((a,x)=>a+x.rate,0)/rates.length;const below=rates.filter(x=>x.rate<avg).sort((a,b)=>a.rate-b.rate);const wy=STORE_ORDER.map(s=>({s,v:DATA[s].wk_yoy})).filter(x=>!isNA(x.v)).sort((a,b)=>a.v-b.v)[0];let o='';if(below.length>0)o+='完成率低于区域均值（'+avg.toFixed(1)+'%）的门店：'+below.map(x=>x.s+'('+x.rate.toFixed(1)+'%)').join('、')+'，是区域达标主要短板，建议城市经理驻店帮扶并复盘目标拆解。';if(wy)o+=' '+wy.s+'周同比最低（'+fmtPct(wy.v)+'），需重点排查品类缺货与竞品分流。';return o||'各门店完成率均衡，区域整体达标进度健康。';}());

  // Pre-compute ltSampleData for use in conc() and chart
  const ltSampleData=STORE_ORDER.map(s=>{
    const d=DATA[s];
    const ltQty=d.daily.reduce((a,x)=>a+(x.lt_qty||0),0);
    const sampD=(DATA['_sample']||{})[s]||{};
    return{s,ltAmt:d.wk_lt||0,ltQty,ltPct:d.wk_amt>0?(d.wk_lt||0)/d.wk_amt*100:0,sampAmt:sampD.wk_amt||0,sampQty:sampD.wk_qty||0};
  }).sort((a,b)=>b.ltAmt-a.ltAmt);

  // Chart 1: Weekly amount comparison (bar)
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">\ud83d\udcc8</span> \u5404\u95e8\u5e97WK29\u5468\u7d2f\u8ba1\u9500\u552e&\u5b8c\u6210\u7387\u5bf9\u6bd4</h3><div class="chart-wrap"><canvas id="cmpAmt"></canvas></div>'+
    conc(function(){const sr=STORE_ORDER.map(s=>{const d=DATA[s];const wt=d.daily.reduce((a,x)=>a+(x.target||0),0);return{s,amt:d.wk_amt,wt,rate:wt>0?d.wk_amt/wt*100:0};}).sort((a,b)=>b.rate-a.rate);return '完成率排名：'+sr.map(x=>x.s+' '+x.rate.toFixed(1)+'%').join('、')+'。';}(),function(){const sr=STORE_ORDER.map(s=>{const d=DATA[s];const wt=d.daily.reduce((a,x)=>a+(x.target||0),0);return{s,amt:d.wk_amt,wt,rate:wt>0?d.wk_amt/wt*100:0};}).sort((a,b)=>b.rate-a.rate);const avgRate=sr.reduce((a,x)=>a+x.rate,0)/sr.length;const below=sr.filter(x=>x.rate<avgRate);const above=sr.filter(x=>x.rate>=avgRate);let o='';if(below.length>0)o+='低于区域均值（'+avgRate.toFixed(1)+'%）的门店：'+below.map(x=>x.s+'('+x.rate.toFixed(1)+'%,差'+(avgRate-x.rate).toFixed(1)+'pct)').join('、')+'，需重点跟进目标达成。';if(above.length>0)o+=above.map(x=>x.s+'('+x.rate.toFixed(1)+'%)').join('、')+'完成率领先，可提炼经验推广。';return o||'各门店完成率较为均衡。';}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">\ud83d\udcca</span> \u5404\u95e8\u5e97\u5468\u540c\u6bd4\u5bf9\u6bd4</h3><div class="chart-wrap"><canvas id="cmpYoy"></canvas></div>'+
    conc('\u5468\u540c\u6bd4\u6700\u9ad8'+STORE_ORDER.map(s=>({s:s,v:DATA[s].wk_yoy})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v)[0]?.s+'('+fmtPct(STORE_ORDER.map(s=>({s:s,v:DATA[s].wk_yoy})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v)[0]?.v)+')\uff0c\u533a\u57df'+fmtPct(regYoy)+'\u3002',function(){const sy=STORE_ORDER.map(s=>({s,v:DATA[s].wk_yoy})).filter(x=>!isNA(x.v));const neg=sy.filter(x=>x.v<0);if(neg.length>0)return neg.map(x=>x.s+'('+fmtPct(x.v)+')').join('、')+'同比下滑，需重点关注其品类结构调整和竞品分流影响。';if(regYoy<0)return '区域整体同比下滑，需检视去年同期爆款是否缺货，调整品类结构。';return '区域同比表现良好，保持品类结构和渠道策略。';}())+'</div></div>';

  // Chart 3: ATV & JD comparison
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">\ud83d\udd29</span> \u5404\u95e8\u5e97\u5ba2\u5355\u4ef7&\u8fde\u5e26\u7387\u5bf9\u6bd4</h3><div class="chart-wrap"><canvas id="cmpAtvJd"></canvas></div>'+
    conc('\u5ba2\u5355\u4ef7\u6700\u9ad8'+STORE_ORDER.map(s=>({s:s,v:getWeeklyATV(s)})).sort((a,b)=>b.v-a.v)[0]?.s+'(\u00a5'+Math.round(STORE_ORDER.map(s=>({s:s,v:getWeeklyATV(s)})).sort((a,b)=>b.v-a.v)[0]?.v)+')\uff0c\u8fde\u5e26\u7387\u6700\u9ad8'+STORE_ORDER.map(s=>({s:s,v:getWeeklyJD(s)})).sort((a,b)=>b.v-a.v)[0]?.s+'('+STORE_ORDER.map(s=>({s:s,v:getWeeklyJD(s)})).sort((a,b)=>b.v-a.v)[0]?.v.toFixed(2)+')\u3002',function(){const atvV=STORE_ORDER.map(s=>({s,v:getWeeklyATV(s)})).sort((a,b)=>b.v-a.v);const jdV=STORE_ORDER.map(s=>({s,v:getWeeklyJD(s)})).sort((a,b)=>b.v-a.v);const ratv=getRegWkATV(),rjd=getRegWkJD();const atvBelow=atvV.filter(x=>x.v<ratv);const jdBelow=jdV.filter(x=>x.v<rjd);let t='';if(atvBelow.length>0)t+=atvBelow.map(x=>x.s+'(\u00a5'+Math.round(x.v)+',差\u00a5'+Math.round(ratv-x.v)+')').join('\u3001')+'\u5ba2\u5355\u4ef7\u4f4e\u4e8e\u533a\u57df\u5747\u503c\u00a5'+Math.round(ratv)+'\uff0c\u5efa\u8bae\u5f15\u5bfc\u9ad8\u4ef7\u5957\u88c5\u548c\u9650\u5b9a\u6b3e\u63a8\u8d2d\u3002';if(jdBelow.length>0)t+=jdBelow.map(x=>x.s+'('+x.v.toFixed(2)+',差'+(rjd-x.v).toFixed(2)+')').join('\u3001')+'\u8fde\u5e26\u7387\u4f4e\u4e8e\u533a\u57df\u5747\u503c'+rjd.toFixed(2)+'\uff0c\u5efa\u8bae\u5f3a\u5316\u8d44\u5bd3\u63a8\u8d2d\u8bdd\u672f\u3002';return t||'\u5404\u95e8\u5e97\u5ba2\u5355\u4ef7\u548c\u8fde\u5e26\u7387\u5747\u8fbe\u533a\u57df\u5747\u503c\uff0c\u4fdd\u6301\u63a8\u8d2d\u7b56\u7565\u3002';}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">\ud83d\udcc9</span> \u5404\u95e8\u5e97WK28 vs WK29\u5468\u73af\u6bd4</h3><div class="chart-wrap"><canvas id="cmpWow"></canvas></div>'+
    conc('\u5468\u73af\u6bd4\u6700\u4f73'+STORE_ORDER.map(s=>({s:s,v:DATA[s].wow})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v)[0]?.s+'('+(STORE_ORDER.map(s=>({s:s,v:DATA[s].wow})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v)[0]?.v>=0?'+':'')+STORE_ORDER.map(s=>({s:s,v:DATA[s].wow})).filter(x=>!isNA(x.v)).sort((a,b)=>b.v-a.v)[0]?.v.toFixed(1)+'%)\uff0c\u533a\u57df'+(isNA(r.wow)?'N/A':(r.wow>=0?'+':'')+r.wow.toFixed(1)+'%')+'\u3002',function(){const wv=STORE_ORDER.map(s=>({s,v:DATA[s].wow})).filter(x=>!isNA(x.v));const reWow=isNA(r.wow)?0:r.wow;const below=wv.filter(x=>x.v<reWow);const neg=wv.filter(x=>x.v<-5);let o='';if(below.length>0&&!isNA(r.wow))o+='低于区域周环比（'+fmtPct(r.wow)+'）的门店：'+below.map(x=>x.s+'('+fmtPct(x.v)+')').join('\u3001')+'\uff0c\u9700\u590d\u76d8\u5ba2\u6d41\u8f6c\u5316\u53d8\u5316\u3002';if(neg.length>0)o+=neg.map(x=>x.s+'('+x.v.toFixed(1)+'%)').join('\u3001')+'\u73af\u6bd4\u4e0b\u6ed1\u8f83\u5927\uff0c\u91cd\u70b9\u5173\u6ce8\u7ade\u54c1\u548c\u4fc3\u9500\u5f71\u54cd\u3002';return o||'\u5404\u95e8\u5e97\u73af\u6bd4\u57fa\u672c\u5e73\u7a33\uff0c\u5173\u6ce8\u5468\u672b\u9ad8\u5cf0\u51b2\u91cf\u3002';}())+'</div></div>';


  // 门店WK29周成交率 & 长尾款 对比（半宽并排，第三排）
  const wcvo=(DATA['_weekly_conversion']||{});
  const convData=STORE_ORDER.map(s=>({s,v:(wcvo[s]&&wcvo[s].flow>0?wcvo[s].rate:null),flow:wcvo[s]?wcvo[s].flow:0,cnt:wcvo[s]?wcvo[s].cnt:0})).sort((a,b)=>b.v-a.v);
  const rConv=wcvo._region||{};
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📦</span> 各门店长尾款 & 样品销售周累计对比（标注占比）</h3><div class="chart-wrap"><canvas id="cmpLtSample"></canvas></div>'+
    conc('区域长尾款周累计'+fmtAmt(totLt)+'（'+totLtQty+'件，占比'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%）。各门店长尾款占比：'+ltSampleData.filter(x=>x.ltAmt>0).map(x=>x.s+x.ltPct.toFixed(1)+'%').join('、')+'。样品周累计'+fmtAmt(sReg.wk_amt||0)+'（'+(sReg.wk_qty||0)+'件）。',function(){const lpct=totAmt>0?totLt/totAmt*100:0;if(lpct<8)return '区域长尾款占比偏低（'+lpct.toFixed(1)+'%），建议统一部署79款SKU推荐话术培训和收银台陈列优化。';if(lpct>15)return '区域长尾款占比优秀（'+lpct.toFixed(1)+'%），保持当前关联推荐策略。';return '长尾款占比适中，关注各门店间均衡性。';}())+
    '</div>'+
    '<div class="chart-card"><h3><span class="icon">🎯</span> 各门店WK29周成交率对比（WK29笔数/客流量）</h3><div class="chart-wrap"><canvas id="cmpConv"></canvas></div>'+
    conc('区域WK29周成交率'+(rConv.rate?rConv.rate.toFixed(2):'N/A')+'%（ZK29成交'+(rConv.cnt||0)+'笔 ÷ ZK29客流'+(rConv.flow||0)+'人次）。各店：'+convData.map(x=>x.s+(x.v===null?'N/A':x.v.toFixed(2)+'%')).join('、')+'。',function(){const valid=convData.filter(x=>x.v!==null);const avg=valid.length?valid.reduce((a,x)=>a+x.v,0)/valid.length:0;const below=valid.filter(x=>x.v<avg);let o='';if(below.length>0)o+='低于区域均值（'+avg.toFixed(2)+'%）的门店：'+below.map(x=>x.s+'('+x.v.toFixed(2)+'%,差'+(avg-x.v).toFixed(2)+'pct)').join('、')+'，客流未有效转化为笔数，需提升进店转化（橱窗陈列/试玩体验/导购拦截话术）。';else o+='各门店成交率较为均衡。';if(valid.length){const best=valid[0],worst=valid[valid.length-1];o+='最高'+best.s+'('+best.v.toFixed(2)+'%)与最低'+worst.s+'('+worst.v.toFixed(2)+'%)相差'+(best.v-worst.v).toFixed(2)+'pct，差距明显可对标提升。';}return o;}())+
    '</div></div>';

  // 新品 & 样品已改为各店KPI卡展示，区域对比图(cmpNewSample)已取消

  // WOS & JL comparison + Series proportion
  const totJl=regVal(x=>x.wk_jl);
  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">\u26a1</span> \u5404\u95e8\u5e97WOS & \u5373\u65f6\u96f6\u552e\u5468\u7d2f\u8ba1\u5bf9\u6bd4</h3><div class="chart-wrap"><canvas id="cmpWosJl"></canvas></div>'+
    conc('WOS\u5468\u7d2f\u8ba1'+fmtAmt(totWos)+'\uff08\u5360\u6bd4'+(totAmt>0?(totWos/totAmt*100).toFixed(1):0)+'%\uff09\uff0c\u5373\u65f6\u96f6\u552e'+fmtAmt(totJl)+'\uff08\u5360\u6bd4'+(totAmt>0?(totJl/totAmt*100).toFixed(1):0)+'%\uff09\u3002',function(){const wpct=totAmt>0?totWos/totAmt*100:0,jpct=totAmt>0?totJl/totAmt*100:0;const lowW=STORE_ORDER.filter(s=>DATA[s].wk_amt>0&&DATA[s].wk_wos/DATA[s].wk_amt*100<3);let t='';if(wpct+jpct<10)t='WOS+即时零售合计占比不足10%，建议加大线上渠道投放力度，特别是美团/饿了么专属活动。';else if(wpct+jpct>20)t='线上渠道占比超过20%，多渠道运营成熟，保持投放节奏。';else t='线上渠道占比中等，建议适度增加即时零售平台曝光和WOS直播频次。';if(lowW.length>0)t+=' '+lowW.join('、')+'WOS占比偏低，需重点提升线上运营能力。';return t;}())+'</div>'+
    '<div class="chart-card"><h3><span class="icon">\ud83d\udcca</span> \u533a\u57df\u4ea7\u54c1\u7cfb\u5217\u7ed3\u6784\uff08WK29\uff09</h3><div class="chart-wrap"><canvas id="cmpRegSeries"></canvas></div>'+
    conc(function(){const rs=(DATA['_series']||{})['_region']||{};const ents=Object.entries(rs).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '\u6682\u65e0\u7cfb\u5217\u6570\u636e\u3002';const top10=ents.slice(0,10);const t3pct=totAmt>0?ents.slice(0,3).reduce((a,[,v])=>a+v,0)/totAmt*100:0;const botAmt=ents.reduce((a,[k,v])=>k.toUpperCase().includes('BOTANICAL')?a+v:a,0);const botPct=totAmt>0?botAmt/totAmt*100:0;return '\u533a\u57dfTOP10\u7cfb\u5217\uff1a'+top10.map(([k,v])=>k+'('+fmtAmt(v)+',\u5360\u6bd4'+(totAmt>0?(v/totAmt*100).toFixed(1):0)+'%)').join('\u3001')+'\uff1bTOP3\u5408\u8ba1\u5360\u6bd4'+t3pct.toFixed(1)+'%\u3002BOTANICALS\u690d\u7269\u7cfb\u5217\u5360\u6bd4'+botPct.toFixed(1)+'%\uff08\u6307\u68074.1%\uff09\u3002';}(),function(){const rs=(DATA['_series']||{})['_region']||{};const ents=Object.entries(rs).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '';const top3=ents.slice(0,3);const top5=ents.slice(0,5);const t3pct=totAmt>0?top3.reduce((a,[,v])=>a+v,0)/totAmt*100:0;let o='';if(t3pct>65)o+='区域TOP3系列占比过高（'+t3pct.toFixed(1)+'%），建议丰富中长尾系列陈列和推荐。';else if(t3pct<30)o+='系列分布过于分散，建议集中打造1-2个主推系列。';else o+='系列结构均衡，保持当前推荐策略。';const focusTop3=['CITY','TECHNIC','NINJAGO'];const missingTop3=focusTop3.filter(n=>!top3.some(([k])=>k.toUpperCase().includes(n)));if(missingTop3.length>0)o+='重点系列'+missingTop3.join('、')+'未进TOP3，建议加强陈列和推荐。';const inTop3=focusTop3.filter(n=>top3.some(([k])=>k.toUpperCase().includes(n)));if(inTop3.length>0)o+='重点系列'+inTop3.join('、')+'在TOP3表现良好。';if(!top5.some(([k])=>k.toUpperCase().includes('MINECRAFT')))o+='MINECRAFT未进TOP5，需关注沙盒游戏IP粉丝群体引流。';const botAmt=ents.reduce((a,[k,v])=>k.toUpperCase().includes('BOTANICAL')?a+v:a,0);const botPct=totAmt>0?botAmt/totAmt*100:0;if(botAmt>0){o+='BOTANICALS植物系列占比'+botPct.toFixed(1)+'%（指标4.1%）';if(botPct<4.1)o+='，低于指标，需加强植物系列陈列和推荐。';else o+='，达到指标。';}else{o+='BOTANICALS植物系列无销售，低于指标4.1%，建议关注植物系列铺货和推荐。';}return o;}())+'</div></div>';



  // Regional TOP10 products
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">\ud83c\udfc6</span> \u533a\u57dfTOP10\u4ea7\u54c1\u5468\u9500\u6392\u540d\uff08\u542b\u4ef6\u6570\uff09</h3><div class="chart-wrap"><canvas id="cmpRegProducts"></canvas></div>'+
    conc(function(){const rp=(DATA['_products']||{})['_region']||{};const ents=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,3);if(ents.length===0)return '\u6682\u65e0\u4ea7\u54c1\u6570\u636e\u3002';return '\u533a\u57dfTOP3\u4ea7\u54c1\uff1a'+ents.map(([n,a,q])=>n+'('+fmtAmt(a)+','+q+'\u4ef6)').join('\u3001')+'\u3002';}(),function(){const rp=(DATA['_products']||{})['_region']||{};const ents=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]);if(ents.length===0)return '';const t1=ents[0];const t1pct=totAmt>0?t1[1]/totAmt*100:0;let o='';if(t1pct>20)o+='TOP1产品占比'+t1pct.toFixed(1)+'%，依赖度高，需关注库存和缺货风险。';else o+='区域产品结构多元，无过度依赖单品的风险。';const newProds=DATA['_new_products']||[];if(newProds.length>0){const top10Names=ents.slice(0,10).map(([k])=>k);const newInTop=newProds.filter(np=>top10Names.some(tn=>tn.includes(np)||np.includes(tn)));if(newInTop.length>0)o+='当月新品「'+newInTop.join('、')+'」进入区域TOP10，表现突出，建议加大陈列推荐。';else o+='当月新品无进入区域TOP10，需关注新品陈列位置和推荐话术。';}return o;}())+'</div></div>';

  // Summary table
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">\ud83d\udccb</span> \u533a\u57df\u5468\u7ef4\u5ea6\u6c47\u603b\u8868</h3><table class="summary-table" id="cmpTable"></table></div></div>';

  // All guides ranking
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">\ud83d\udc69</span> \u5168\u533a\u57df\u5bfc\u8d2d\u5468\u9500\u552e\u6392\u540d\uff08\u5468\u73af\u6bd4\uff09</h3><div class="chart-wrap tall" style="height:500px;"><canvas id="cmpGuides"></canvas></div>'+
    conc(function(){const all=[];STORE_ORDER.forEach(s=>{(DATA['_guides']||{})[s]?.forEach(g=>all.push(g));});all.sort((a,b)=>b.wk_sales-a.wk_sales);if(all.length===0)return '\u6682\u65e0\u5bfc\u8d2d\u6570\u636e\u3002';const t=all[0];return '\u5468\u9500\u552e\u6392\u540d\u7b2c\u4e00'+t.name+'('+t.store+')'+fmtAmt(t.wk_sales)+'\uff0c\u5468\u73af\u6bd4'+(isNA(t.wow)?'N/A':(t.wow>=0?'+':'')+t.wow.toFixed(1)+'%')+'\u3002';}(),function(){const all=[];STORE_ORDER.forEach(s=>{(DATA['_guides']||{})[s]?.forEach(g=>all.push(g));});all.sort((a,b)=>b.wk_sales-a.wk_sales);if(all.length===0)return '';const regAvg=all.filter(g=>g.wk_sales>0).reduce((a,g)=>a+g.wk_sales,0)/all.filter(g=>g.wk_sales>0).length;const below=all.filter(g=>g.wk_sales>0&&g.wk_sales<regAvg*0.7);const neg=all.filter(g=>!isNA(g.wow)&&g.wow<-8);let o='';if(below.length>0)o+='⚠️ 低于区域均值¥'+Math.round(regAvg)+'的导购：'+below.map(g=>g.name+'('+g.store+',¥'+Math.round(g.wk_sales)+')').join('\u3001')+'\uff0c\u9700\u91cd\u70b9\u8fdb\u884c\u9500\u552e\u8f85\u5bfc\u548c\u590d\u76d8\u3002';if(neg.length>0)o+=neg.map(g=>g.name+'('+g.store+', '+g.wow.toFixed(1)+'%)').join('\u3001')+'\u73af\u6bd4\u660e\u663e\u4e0b\u6ed1\uff0c\u9700\u590d\u76d8\u5ba2\u6d41\u53d8\u5316\u548c\u63a8\u8d2d\u8f6c\u5316\u3002';const top1=all[0];const topPct=all.reduce((a,g)=>a+g.wk_sales,0)>0?top1.wk_sales/all.reduce((a,g)=>a+g.wk_sales,0)*100:0;if(topPct>40)o+='TOP1\u5bfc\u8d2d\u5360\u533a\u57df'+topPct.toFixed(1)+'%\uff0c\u5efa\u8bae\u5934\u90e8\u7ecf\u9a8c\u5206\u4eab\u590d\u5236\u3002';return o||'\u5bfc\u8d2d\u56e2\u961f\u6574\u4f53\u5e73\u7a33\uff08\u533a\u57df\u4eba\u5747¥'+Math.round(regAvg)+'\uff09\uff0c\u7ee7\u7eed\u5173\u6ce8\u5468\u672b\u9ad8\u5cf0\u51b2\u91cf\u3002';}())+'</div></div>';

  document.getElementById('content').innerHTML=html;
  Object.values(charts).forEach(c=>c&&c.destroy());charts={};

  // Chart 1: Weekly amount + completion rate (sorted by rate)
  const storeRates=STORE_ORDER.map(s=>{const d=DATA[s];const wt=d.daily.reduce((a,x)=>a+(x.target||0),0);return{s,v:d.wk_amt,target:wt,rate:wt>0?d.wk_amt/wt*100:0};}).sort((a,b)=>b.rate-a.rate);
  charts.cmpAmt=new Chart(document.getElementById('cmpAmt'),{
    data:{labels:storeRates.map(x=>x.s+'\n'+x.rate.toFixed(1)+'%'),datasets:[
      {type:'bar',label:'\u5468\u7d2f\u8ba1',data:storeRates.map(x=>x.v),backgroundColor:storeRates.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:storeRates.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6,yAxisID:'y'},
      {type:'line',label:'\u5b8c\u6210\u7387',data:storeRates.map(x=>x.rate),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:5,pointBackgroundColor:'#dc2626',tension:.3,yAxisID:'y1'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const x=storeRates[ctx.dataIndex];if(ctx.dataset.type==='line')return '\u5b8c\u6210\u7387: '+x.rate.toFixed(1)+'%';return '\u5468\u9500\u552e: '+fmtAmt(x.v)+' | \u5468\u6307\u6807: '+fmtAmt(x.target)+' | \u5b8c\u6210\u7387: '+x.rate.toFixed(1)+'%';}}}},
      scales:{y:{position:'left',title:{display:true,text:'\u5468\u7d2f\u8ba1(\u00a5)'},ticks:{callback:v=>fmtAmt(v)}},y1:{position:'right',title:{display:true,text:'\u5b8c\u6210\u7387(%)'},grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}}
  });

  // Chart 2: YoY
  const yoyData=STORE_ORDER.map(s=>({s,v:DATA[s].wk_yoy}));
  charts.cmpYoy=new Chart(document.getElementById('cmpYoy'),{
    type:'bar',
    data:{labels:yoyData.map(x=>x.s),datasets:[{label:'\u5468\u540c\u6bd4',data:yoyData.map(x=>x.v),backgroundColor:yoyData.map(x=>isNA(x.v)?'#e2e8f0':(x.v>=0?'#16a34acc':'#dc2626cc')),borderColor:yoyData.map(x=>isNA(x.v)?'#94a3b8':(x.v>=0?'#16a34a':'#dc2626')),borderWidth:1,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>isNA(ctx.raw)?'N/A':'\u5468\u540c\u6bd4: '+(ctx.raw>=0?'+':'')+ctx.raw.toFixed(1)+'%'}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}}
  });

  // Chart 3: ATV & JD (dual axis)
  charts.cmpAtvJd=new Chart(document.getElementById('cmpAtvJd'),{
    data:{labels:STORE_ORDER,datasets:[
      {type:'bar',label:'\u5ba2\u5355\u4ef7',data:STORE_ORDER.map(s=>getWeeklyATV(s)),backgroundColor:'#0ea5e9cc',borderColor:'#0ea5e9',borderWidth:1,borderRadius:4,yAxisID:'y'},
      {type:'line',label:'\u8fde\u5e26\u7387',data:STORE_ORDER.map(s=>getWeeklyJD(s)),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#dc2626',tension:.3,yAxisID:'y1'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label.includes('\u5ba2\u5355'))return '\u5ba2\u5355\u4ef7: \u00a5'+Math.round(ctx.raw);return '\u8fde\u5e26\u7387: '+ctx.raw.toFixed(2)}}}},
      scales:{y:{position:'left',title:{display:true,text:'\u5ba2\u5355\u4ef7(\u00a5)'},ticks:{callback:v=>'\u00a5'+v}},y1:{position:'right',title:{display:true,text:'\u8fde\u5e26\u7387'},grid:{drawOnChartArea:false},ticks:{callback:v=>v.toFixed(1)}}}}
  });

  // Chart 4: WoW
  const wowData=STORE_ORDER.map(s=>({s,v:DATA[s].wow}));
  charts.cmpWow=new Chart(document.getElementById('cmpWow'),{
    type:'bar',
    data:{labels:wowData.map(x=>x.s),datasets:[
      {label:'WK28',data:wowData.map(x=>DATA[x.s].wk28_amt||0),backgroundColor:'#94a3b880',borderColor:'#64748b',borderWidth:1,borderRadius:4},
      {label:'WK29',data:wowData.map(x=>DATA[x.s].wk_amt||0),backgroundColor:wowData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:wowData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmtAmt(ctx.raw)}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 5: WOS & JL comparison
  charts.cmpWosJl=new Chart(document.getElementById('cmpWosJl'),{
    type:'bar',
    data:{labels:STORE_ORDER,datasets:[
      {label:'WOS',data:STORE_ORDER.map(s=>DATA[s].wk_wos),backgroundColor:'#0284c7cc',borderColor:'#0284c7',borderWidth:1,borderRadius:4},
      {label:'\u5373\u65f6\u96f6\u552e',data:STORE_ORDER.map(s=>DATA[s].wk_jl),backgroundColor:'#ea580ccc',borderColor:'#ea580c',borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmtAmt(ctx.raw)}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: Combined long-tail & sample comparison
  charts.cmpLtSample=new Chart(document.getElementById('cmpLtSample'),{
    type:'bar',
    data:{labels:ltSampleData.map(x=>x.s),datasets:[
      {label:'\u957f\u5c3e\u6b3e',data:ltSampleData.map(x=>x.ltAmt),backgroundColor:'#ea580ccc',borderColor:'#ea580c',borderWidth:1,borderRadius:4},
      {label:'\u6837\u54c1',data:ltSampleData.map(x=>x.sampAmt),backgroundColor:'#9333eacc',borderColor:'#9333ea',borderWidth:1,borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const x=ltSampleData[ctx.dataIndex];if(ctx.datasetIndex===0)return '\u957f\u5c3e\u6b3e: '+fmtAmt(x.ltAmt)+' ('+x.ltQty+'\u4ef6, \u5360\u6bd4'+x.ltPct.toFixed(1)+'%)';return '\u6837\u54c1: '+fmtAmt(x.sampAmt)+' ('+x.sampQty+'\u4ef6)';}}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 6: Regional series doughnut
  const rs=(DATA['_series']||{})['_region']||{};
  const rse=Object.entries(rs).sort((a,b)=>b[1]-a[1]);
  const rtop10=rse.slice(0,10),roth=rse.slice(10).reduce((a,[,v])=>a+v,0);
  charts.cmpRegSeries=new Chart(document.getElementById('cmpRegSeries'),{
    type:'doughnut',
    data:{labels:rtop10.map(([k])=>k).concat(roth>0?['\u5176\u4ed6']:[]),datasets:[{data:rtop10.map(([,v])=>v).concat(roth>0?[roth]:[]),backgroundColor:SC.slice(0,rtop10.length+(roth>0?1:0)),borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},padding:6}},tooltip:{callbacks:{label:ctx=>{const pct=totAmt>0?(ctx.raw/totAmt*100).toFixed(1):0;return ctx.label+': '+fmtAmt(ctx.raw)+' ('+pct+'%)'}}}},cutout:'55%'}
  });

  // Chart 7: Regional TOP10 products
  const rp=(DATA['_products']||{})['_region']||{};
  const rpEnts=Object.entries(rp).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  charts.cmpRegProducts=new Chart(document.getElementById('cmpRegProducts'),{
    type:'bar',
    data:{labels:rpEnts.map(([k])=>k),datasets:[{label:'\u5468\u9500\u552e',data:rpEnts.map(([k,a,q])=>a),backgroundColor:SC.slice(0,rpEnts.length).map(c=>c+'cc'),borderColor:SC.slice(0,rpEnts.length),borderWidth:1,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const idx=ctx.dataIndex;const [n,a,q]=rpEnts[idx];const pct=totAmt>0?(a/totAmt*100).toFixed(1):0;return n+': '+fmtAmt(a)+' ('+pct+'%, '+q+'\u4ef6)'}}}},
      scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: 门店WK29周成交率
  charts.cmpConv=new Chart(document.getElementById('cmpConv'),{
    type:'bar',
    data:{labels:convData.map(x=>x.s),datasets:[{label:'周成交率',data:convData.map(x=>x.v),backgroundColor:convData.map(x=>STORE_COLORS[x.s]+'cc'),borderColor:convData.map(x=>STORE_COLORS[x.s]),borderWidth:1,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const x=convData[ctx.dataIndex];return '周成交率: '+(x.v===null?'缺客流':x.v.toFixed(2)+'%')+' | 笔数'+x.cnt+' / 客流'+x.flow;}}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}}
  });

  // Summary table
  let tHtml='<thead><tr><th>\u95e8\u5e97</th><th>\u5468\u7d2f\u8ba1</th><th>\u5468\u73af\u6bd4</th><th>\u5468\u540c\u6bd4</th><th>\u7b14\u6570</th><th>\u5ba2\u5355\u4ef7</th><th>\u8fde\u5e26\u7387</th><th>周成交率</th><th>WOS</th><th>\u5373\u65f6\u96f6\u552e</th><th>\u957f\u5c3e\u6b3e</th><th>\u957f\u5c3e\u4ef6\u6570</th><th>\u6837\u54c1</th><th>\u6837\u54c1\u4ef6\u6570</th><th>\u957f\u5c3e\u5360\u6bd4</th><th>\u65b0\u54c1\u5468\u7d2f\u8ba1</th><th>\u65b0\u54c1\u5360\u6bd4</th></tr></thead><tbody>';
  // 区域均值（区域汇总表低于均值标红）
  const regWkAmtAvg=totAmt/STORE_ORDER.length;
  const regWkCntAvg=totCnt/STORE_ORDER.length;
  const regWkATV=getRegWkATV();
  const regWkJD=getRegWkJD();
  const regWkLtPctAvg=STORE_ORDER.reduce((a,s)=>a+(DATA[s].wk_amt>0?(DATA[s].wk_lt/DATA[s].wk_amt*100):0),0)/STORE_ORDER.length;
  STORE_ORDER.forEach(s=>{
    const d=DATA[s];const atv=getWeeklyATV(s),jd=getWeeklyJD(s),cnt=getWeeklyCnt(s);
    const lpct=d.wk_amt>0?(d.wk_lt/d.wk_amt*100):0;
    const ltQ=d.daily.reduce((a,x)=>a+(x.lt_qty||0),0);
    const sampD=(DATA['_sample']||{})[s]||{};
    const npAmt=d.wk_np_amt||0;const npPct=d.wk_amt>0?npAmt/d.wk_amt*100:0;
    const wb=isNA(d.wow)?'<span class="badge gray">N/A</span>':(d.wow>=0?'<span class="badge green">+'+d.wow.toFixed(1)+'%</span>':'<span class="badge red">'+d.wow.toFixed(1)+'%</span>');
    const yb=isNA(d.wk_yoy)?'<span class="badge gray">N/A</span>':(d.wk_yoy>=0?'<span class="badge green">+'+d.wk_yoy.toFixed(1)+'%</span>':'<span class="badge red">'+d.wk_yoy.toFixed(1)+'%</span>');
    tHtml+='<tr><td class="store-name" style="color:'+STORE_COLORS[s]+'">'+s+'</td><td style="font-weight:700'+(d.wk_amt<regWkAmtAvg?';color:#dc2626':'')+'">'+fmtAmt(d.wk_amt)+'</td><td>'+wb+'</td><td>'+yb+'</td><td'+(cnt<regWkCntAvg?' style="color:#dc2626;font-weight:700"':'')+'>'+cnt+'</td><td'+(atv<regWkATV?' style="color:#dc2626;font-weight:700"':'')+'>¥'+Math.round(atv)+'</td><td'+(jd<regWkJD?' style="color:#dc2626;font-weight:700"':'')+'>'+jd.toFixed(2)+'</td><td style="color:'+((wcvo[s]&&wcvo[s].flow>0&&wcvo[s].rate>=rConv.rate)?'#16a34a':'#dc2626')+';font-weight:700">'+((wcvo[s]&&wcvo[s].flow>0)?wcvo[s].rate.toFixed(2)+'%':'缺客流')+'</td><td>'+fmtAmt(d.wk_wos)+'</td><td>'+fmtAmt(d.wk_jl)+'</td><td>'+fmtAmt(d.wk_lt)+'</td><td>'+ltQ+'\u4ef6</td><td'+(sampD.wk_amt>0?'':' style="color:#dc2626;font-weight:700"')+'>'+fmtAmt(sampD.wk_amt||0)+(sampD.wk_amt>0?'':'｜未开单')+'</td><td>'+(sampD.wk_qty||0)+'\u4ef6</td><td style="color:'+(lpct<regWkLtPctAvg?'#dc2626':'#16a34a')+';font-weight:700">'+lpct.toFixed(1)+'%</td><td>'+fmtAmt(npAmt)+'</td><td style="color:'+(npPct<20?'#dc2626':'#16a34a')+';font-weight:700">'+npPct.toFixed(1)+'%</td></tr>';
  });
  // Totals row
  tHtml+='<tr style="background:#7c3aed10;font-weight:700"><td class="store-name">\u533a\u57df\u5408\u8ba1</td><td>'+fmtAmt(totAmt)+'</td><td>'+(isNA(r.wow)?'<span class="badge gray">N/A</span>':(r.wow>=0?'<span class="badge green">+'+r.wow.toFixed(1)+'%</span>':'<span class="badge red">'+r.wow.toFixed(1)+'%</span>'))+'</td><td>'+(isNA(regYoy)?'<span class="badge gray">N/A</span>':(regYoy>=0?'<span class="badge green">+'+regYoy.toFixed(1)+'%</span>':'<span class="badge red">'+regYoy.toFixed(1)+'%</span>'))+'</td><td>'+totCnt+'</td><td>\u00a5'+Math.round(getRegWkATV())+'</td><td>'+getRegWkJD().toFixed(2)+'</td><td style="color:#16a34a;font-weight:700">'+(rConv.rate===null?'N/A':rConv.rate.toFixed(2)+'%')+'%</td><td>'+fmtAmt(totWos)+'</td><td>'+fmtAmt(totJl)+'</td><td>'+fmtAmt(totLt)+'</td><td>'+totLtQty+'\u4ef6</td><td>'+fmtAmt(sReg.wk_amt||0)+'</td><td>'+(sReg.wk_qty||0)+'\u4ef6</td><td style="color:'+(totAmt>0&&(totLt/totAmt*100)>=10?'#16a34a':'#dc2626')+'">'+(totAmt>0?(totLt/totAmt*100).toFixed(1):0)+'%</td><td>'+fmtAmt(npRegAmt)+'</td><td style="color:'+(totAmt>0&&(npRegAmt/totAmt*100)>=20?'#16a34a':'#dc2626')+';font-weight:700">'+(totAmt>0?(npRegAmt/totAmt*100).toFixed(1):0)+'%</td></tr>';
  tHtml+='</tbody>';
  document.getElementById('cmpTable').innerHTML=tHtml;
  // 区域均值标红说明
  let cmpNote=document.getElementById('cmpTableNote');
  if(cmpNote) cmpNote.remove();
  cmpNote=document.createElement('div');
  cmpNote.id='cmpTableNote';
  cmpNote.style='margin-top:10px;padding:8px 12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;font-size:13px;color:#dc2626;font-weight:700';
  cmpNote.textContent='🔴 周累计/笔数/客单价/连带率/长尾占比 低于区域均值已标红（周累计¥'+Math.round(regWkAmtAvg)+'、笔数'+Math.round(regWkCntAvg)+'、客单价¥'+Math.round(regWkATV)+'、连带率'+regWkJD.toFixed(2)+'、长尾占比'+regWkLtPctAvg.toFixed(1)+'%；新品占比<20%标红；样品未开单标红';
  document.getElementById('cmpTable').parentNode.appendChild(cmpNote);
  let _cmpConc=document.createElement('div');
  _cmpConc.style.cssText='margin-top:10px';
  const _below=STORE_ORDER.filter(s=>DATA[s].wk_amt<regWkAmtAvg);
  const _dual=STORE_ORDER.filter(s=>!isNA(DATA[s].wk_yoy)&&DATA[s].wk_yoy<0&&!isNA(DATA[s].wow)&&DATA[s].wow<0);
  const _topRate=STORE_ORDER.map(s=>{const d=DATA[s];const wt=d.daily.reduce((a,x)=>a+(x.target||0),0);return{s,rate:wt>0?d.wk_amt/wt*100:0};}).sort((a,b)=>b.rate-a.rate)[0];
  _cmpConc.innerHTML=conc('区域7店中'+_below.length+'家周累计低于区域均值（¥'+Math.round(regWkAmtAvg)+'）：'+(_below.length?_below.map(s=>s).join('、')+'。':'均超均值。')+_topRate.s+'完成率领先（'+_topRate.rate.toFixed(1)+'%）。',(_dual.length>0?(_dual.map(s=>s).join('、')+'同环比双降，是区域主要拖累，建议集中资源帮扶并复盘品类与客流。'):(_below.length===0?'各门店周累计与同环比表现均衡，区域整体健康。':'')));
  document.getElementById('cmpTable').parentNode.appendChild(_cmpConc);

  // All guides ranking chart (bar=周销售 + line=周环比)
  const allGuides=[];
  STORE_ORDER.forEach(s=>{((DATA['_guides']||{})[s]||[]).forEach(g=>allGuides.push(g));});
  allGuides.sort((a,b)=>b.wk_sales-a.wk_sales);
  if(allGuides.length>0){
    charts.cmpGuides=new Chart(document.getElementById('cmpGuides'),{
      data:{labels:allGuides.map(g=>g.name+'('+g.store+')'),datasets:[
        {type:'bar',label:'\u5468\u9500\u552e',data:allGuides.map(g=>g.wk_sales),backgroundColor:allGuides.map(g=>STORE_COLORS[g.store]+'cc'),borderColor:allGuides.map(g=>STORE_COLORS[g.store]),borderWidth:1,borderRadius:4,xAxisID:'x'},
        {type:'line',label:'\u5468\u73af\u6bd4',data:allGuides.map(g=>isNA(g.wow)?null:g.wow),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:5,pointBackgroundColor:allGuides.map(g=>isNA(g.wow)?'#94a3b8':(g.wow>=0?'#16a34a':'#dc2626')),tension:.3,xAxisID:'x1'}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const g=allGuides[ctx.dataIndex];if(ctx.dataset.label.includes('\u5468\u9500'))return '\u5468\u9500\u552e: '+fmtAmt(g.wk_sales);return '\u5468\u73af\u6bd4: '+(isNA(g.wow)?'N/A':(g.wow>=0?'+':'')+g.wow.toFixed(1)+'%');}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'\u5468\u9500\u552e(\u00a5)'},ticks:{callback:v=>fmtAmt(v)}},x1:{position:'top',title:{display:true,text:'\u5468\u73af\u6bd4(%)'},grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}}
    });
  }
}

function selectStore(store){
  currentStore=store;
  document.querySelectorAll('.store-tab').forEach(t=>t.classList.toggle('active',t.dataset.store===store));
  if(store==='\u533a\u57df\u5bf9\u6bd4'){renderComparison();}
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

# 替换WK标签为动态编号（WK28/wk28_amt大小写区分，安全替换）
_html_out = HTML.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}').replace('ZK29','WK29')

with open('os.path.join(OUTPUTS_DIR, '')', 'a', encoding='utf-8') as f:
    f.write(_html_out)

import os
sz = os.path.getsize('os.path.join(OUTPUTS_DIR, '')')
print(f"Part 4 appended. File size: {sz:,} bytes")
