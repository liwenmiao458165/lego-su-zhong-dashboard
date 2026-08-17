#!/usr/bin/env python3
"""Append renderStore function to weekly HTML"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import OUTPUTS_DIR
import os
import sys
import glob
import pandas as pd

OUTPUT = 'os.path.join(OUTPUTS_DIR, '')'

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
print(f"  p3: WK{WK_NUM_INT} (prev WK{WK_PREV_NUM})")

JS_RENDER_STORE = r"""
function renderStore(store){
  const d=DATA[store],region=DATA['_region'],color=STORE_COLORS[store],daily=d.daily;
  const wcvo=(DATA['_weekly_conversion']||{});
  const cvoRate=(wcvo[store]&&wcvo[store].flow>0?wcvo[store].rate:null);
  const cvoRegionRate=(wcvo._region&&wcvo._region.flow>0?wcvo._region.rate:null);
  const sampD=(DATA['_sample']||{})[store]||{};
  const sampAmt=sampD.wk_amt||0,sampQty=sampD.wk_qty||0;
  const npAmt=d.wk_np_amt||0,npPct=d.wk_amt>0?npAmt/d.wk_amt*100:0;
  let html='<div class="kpi-row">'+
    '<div class="kpi-card"><div class="label">WK29周累计</div><div class="value" style="color:'+color+'">'+fmtAmt(d.wk_amt)+'</div><div class="sub">周完成率 '+(function(){const wt=daily.reduce((a,x)=>a+(x.target||0),0);const r=wt>0?d.wk_amt/wt*100:0;return '<span style="color:'+(r>=50?'#16a34a':'#dc2626')+';font-weight:700">'+r.toFixed(1)+'%</span>';})()+'</div></div>'+
    '<div class="kpi-card"><div class="label">周环比</div><div class="value '+(isNA(d.wow)?'':colorPct(d.wow))+'">'+(isNA(d.wow)?'N/A':(d.wow>=0?'+':'')+d.wow.toFixed(1)+'%')+'</div><div class="sub">'+(isNA(d.wow)?'新店无上周':'vs WK28 '+fmtAmt(d.wk28_amt))+'</div></div>'+
    '<div class="kpi-card"><div class="label">周同比</div><div class="value '+(isNA(d.wk_yoy)?'':colorPct(d.wk_yoy))+'">'+(isNA(d.wk_yoy)?'N/A':fmtPct(d.wk_yoy))+'</div><div class="sub">'+(isNA(d.wk_yoy)?'新店无同期':'vs去年同期')+'</div></div>'+
    '<div class="kpi-card"><div class="label">周连带率</div><div class="value blue">'+getWeeklyJD(store).toFixed(2)+'</div><div class="sub">区域'+getRegWkJD().toFixed(2)+'</div></div>'+
    '<div class="kpi-card"><div class="label">周客单价</div><div class="value blue">¥'+Math.round(getWeeklyATV(store))+'</div><div class="sub">区域¥'+Math.round(region.wk_atv||getRegWkATV())+'</div></div>'+
    '<div class="kpi-card"><div class="label">WOS周累计</div><div class="value orange">'+fmtAmt(d.wk_wos)+'</div><div class="sub">占比'+(d.wk_amt>0?(d.wk_wos/d.wk_amt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">长尾款周累计</div><div class="value purple">'+fmtAmt(d.wk_lt)+'</div><div class="sub">占比'+(d.wk_amt>0?(d.wk_lt/d.wk_amt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">WK29周成交率</div><div class="value '+(cvoRate===null?'gray':(cvoRate>=cvoRegionRate?'green':'red'))+'">'+(cvoRate===null?'缺客流':cvoRate.toFixed(2)+'%')+'</div><div class="sub">区域'+(cvoRegionRate===null?'N/A':cvoRegionRate.toFixed(2)+'%')+'% | '+(wcvo[store]?wcvo[store].cnt:'0')+'笔 / '+(wcvo[store]?wcvo[store].flow:'0')+'客流</div></div>'+
    '<div class="kpi-card"><div class="label">样品周累计</div><div class="value" style="color:'+(sampAmt>0?'#9333ea':'#dc2626')+'">'+fmtAmt(sampAmt)+'</div><div class="sub">'+(sampQty||0)+'件'+(sampAmt>0?'（独立口径，不计入门店销售额）':'｜未开单')+'</div></div>'+
    '<div class="kpi-card"><div class="label">新品周累计</div><div class="value" style="color:'+(d.wk_amt>0&&npPct<20?'#dc2626':'#db2777')+'">'+fmtAmt(npAmt)+'</div><div class="sub">新品占比'+(d.wk_amt>0?npPct:0).toFixed(1)+'%'+(d.wk_amt>0&&npPct<20?'｜偏低':'')+'</div></div>'+
    '</div>';

  html+=concKPI(store);

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📊</span> 产品系列结构（WK29）</h3><div class="chart-wrap"><canvas id="chartSeriesPie"></canvas></div>'+concSeries(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🎯</span> 门店周维度能力雷达图（vs区域均值）</h3><div class="chart-wrap"><canvas id="chartRadar"></canvas></div>'+concRadar(store)+'</div></div>';

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📊</span> WK28 vs WK29 周销售环比</h3><div class="chart-wrap"><canvas id="chartWoW"></canvas></div>'+concWow(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🏆</span> TOP10产品周销排名（含件数）</h3><div class="chart-wrap"><canvas id="chartTopProducts"></canvas></div>'+concTop(store)+'</div></div>';

  const totWos=daily.reduce((a,x)=>a+x.wos,0),totJl=daily.reduce((a,x)=>a+x.jl,0);
  const wosPct=d.wk_amt>0?(totWos/d.wk_amt*100):0,jlPct=d.wk_amt>0?(totJl/d.wk_amt*100):0,bothPct=d.wk_amt>0?((totWos+totJl)/d.wk_amt*100):0;
  const rTotWos=regVal(x=>x.wk_wos),rTotJl=regVal(x=>x.wk_jl),rTot=regVal(x=>x.wk_amt);
  const rBothPct=rTot>0?((rTotWos+rTotJl)/rTot*100):0;
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">⚡</span> WOS & 即时零售 周累计对比（vs区域）</h3><div class="chart-wrap"><canvas id="chartWosJl"></canvas></div>'+
    conc(store+'本周WOS '+fmtAmt(totWos)+'（占比'+wosPct.toFixed(1)+'%），即时零售 '+fmtAmt(totJl)+'（占比'+jlPct.toFixed(1)+'%），两渠道合计占比'+bothPct.toFixed(1)+'%。',bothPct<rBothPct?'WOS+即时零售合计占比低于区域（'+rBothPct.toFixed(1)+'%），建议加大线上渠道活动投放，强化美团/饿了么专属优惠引流。':bothPct>rBothPct+3?'线上渠道占比高于区域，多渠道运营成熟，建议保持投放策略并关注转化效率。':'线上渠道占比与区域持平，适度增加即时零售平台曝光。')+'</div></div>';

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">👩</span> 导购周销售+WOS+周环比</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideSales"></canvas></div>'+concGuideSales(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🔄</span> 导购连带率+客单价</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideJdKdj"></canvas></div>'+concGuideJdKdj(store)+'</div></div>';

  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">📦</span> 导购长尾款分析（金额/件数）</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideLt"></canvas></div>'+concGuideLt(store)+'</div></div>';
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">📋</span> 导购周维度明细</h3><table class="summary-table" id="guideTable"></table></div></div>';

  document.getElementById('content').innerHTML=html;
  Object.values(charts).forEach(c=>c&&c.destroy());charts={};

  // Chart: Series Pie
  const sd=(DATA['_series']||{})[store]||{};
  const se=Object.entries(sd).sort((a,b)=>b[1]-a[1]);
  const top10=se.slice(0,10),oth=se.slice(10).reduce((a,[k,v])=>a+v,0);
  charts.seriesPie=new Chart(document.getElementById('chartSeriesPie'),{
    type:'doughnut',
    data:{labels:top10.map(([k])=>k).concat(oth>0?['其他']:[]),datasets:[{data:top10.map(([k,v])=>v).concat(oth>0?[oth]:[]),backgroundColor:SC.slice(0,top10.length+(oth>0?1:0)),borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},padding:6}},tooltip:{callbacks:{label:ctx=>{const pct=d.wk_amt>0?(ctx.raw/d.wk_amt*100).toFixed(1):0;return ctx.label+': '+fmtAmt(ctx.raw)+' ('+pct+'%)'}}}},cutout:'55%'}
  });

  // Chart 3: Radar
  const atv=getWeeklyATV(store),jd=getWeeklyJD(store),cnt=getWeeklyCnt(store);
  const wpct=d.wk_amt>0?d.wk_wos/d.wk_amt*100:0,lpct=d.wk_amt>0?d.wk_lt/d.wk_amt*100:0;
  const ratv=getRegWkATV(),rjd=getRegWkJD(),rcnt=regVal(getWeeklyCnt)/7;
  const rwpct=regVal(x=>x.wk_wos)/(regVal(x=>x.wk_amt)||1)*100,rlpct=regVal(x=>x.wk_lt)/(regVal(x=>x.wk_amt)||1)*100;
  const rIdx=(sv,rv)=>rv>0?sv/rv*100:0;
  const rYoyIdx=(sy,ry)=>{if(sy===null)return 0;if(ry===null)return 100;return 100+(sy-ry)*3};
  const rL=['周同比','周客单价','周连带率','周笔数','WOS占比','长尾款占比'];
  const rRS=[d.wk_yoy,atv,jd,cnt,wpct,lpct],rRR=[region.wk_yoy,ratv,rjd,rcnt,rwpct,rlpct];
  const rF=['%','¥','','笔','%','%'];
  const sR=[rYoyIdx(d.wk_yoy,region.wk_yoy),rIdx(atv,ratv),rIdx(jd,rjd),rIdx(cnt,rcnt),rIdx(wpct,rwpct),rIdx(lpct,rlpct)];
  charts.radar=new Chart(document.getElementById('chartRadar'),{
    type:'radar',
    data:{labels:rL,datasets:[
      {label:store,data:sR,borderColor:color,backgroundColor:color+'30',borderWidth:2,pointRadius:4},
      {label:'区域均值(=100)',data:[100,100,100,100,100,100],borderColor:'#64748b',backgroundColor:'#64748b15',borderWidth:2,pointRadius:3,borderDash:[4,4]}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const i=ctx.dataIndex,u=rF[i],sv=rRS[i],rv=rRR[i];const sStr=sv===null?'N/A':(u==='¥'?'¥'+Math.round(sv):(u==='%'?sv.toFixed(1)+'%':(u==='笔'?sv+'笔':sv.toFixed(2))));const rStr=rv===null?'N/A':(u==='¥'?'¥'+Math.round(rv):(u==='%'?rv.toFixed(1)+'%':(u==='笔'?Math.round(rv)+'笔':rv.toFixed(2))));return ctx.dataset.label+': 指数'+ctx.raw.toFixed(0)+' | 门店'+sStr+' | 区域'+rStr;}}}},
      scales:{r:{beginAtZero:true,suggestedMin:0,suggestedMax:200,ticks:{stepSize:50}}}}
  });

  // Chart 5: WoW comparison
  const wk28=d.wk28_amt||0,wk29=d.wk_amt||0;
  charts.wow=new Chart(document.getElementById('chartWoW'),{
    type:'bar',
    data:{labels:['WK28(上周)','WK29(本周)'],datasets:[{label:'周销售',data:[wk28,wk29],backgroundColor:[wk28>=wk29?'#16a34a80':'#94a3b880',wk29>=wk28?color+'cc':'#dc262680'],borderColor:['#64748b',color],borderWidth:2,borderRadius:8,barPercentage:.5}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>'周销售: '+fmtAmt(ctx.raw),afterLabel:ctx=>{if(ctx.dataIndex===1&&wk28>0){const diff=((wk29-wk28)/wk28*100);return '环比: '+(diff>=0?'+':'')+diff.toFixed(1)+'%'}return ''}}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 6: TOP10 Products (with qty)
  const pd=(DATA['_products']||{})[store]||{};
  const pEnts=Object.entries(pd).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  charts.topProducts=new Chart(document.getElementById('chartTopProducts'),{
    type:'bar',
    data:{labels:pEnts.map(([k])=>k),datasets:[{label:'周销售',data:pEnts.map(([k,a,q])=>a),backgroundColor:SC.slice(0,pEnts.length).map(c=>c+'cc'),borderColor:SC.slice(0,pEnts.length),borderWidth:1,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const idx=ctx.dataIndex;const [n,a,q]=pEnts[idx];const pct=d.wk_amt>0?(a/d.wk_amt*100).toFixed(1):0;return n+': '+fmtAmt(a)+' ('+pct+'%, '+q+'件)'}}}},
      scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: WOS & 即时零售 weekly comparison (store vs region)
  charts.wosJl=new Chart(document.getElementById('chartWosJl'),{
    type:'bar',
    data:{labels:['WOS','即时零售'],datasets:[
      {label:store,data:[totWos,totJl],backgroundColor:[color+'cc',color+'cc'],borderColor:[color,color],borderWidth:1,borderRadius:6},
      {label:'区域均值',data:[rTotWos/7,rTotJl/7],backgroundColor:['#94a3b8cc','#94a3b8cc'],borderColor:['#64748b','#64748b'],borderWidth:1,borderRadius:6}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmtAmt(ctx.raw)}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Guide charts
  const guides=(DATA['_guides']||{})[store]||[];
  if(guides.length>0){
    const gLabels=guides.map(g=>g.name);

    // Chart 7: Guide sales + WOS + WoW (merged)
    charts.guideSales=new Chart(document.getElementById('chartGuideSales'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'周销售',data:guides.map(g=>g.wk_sales),backgroundColor:guides.map(g=>g.wk_sales>0?color+'cc':'#e2e8f0'),borderColor:color,borderWidth:1,borderRadius:4,xAxisID:'x',order:3},
        {type:'bar',label:'WOS',data:guides.map(g=>g.wk_wos),backgroundColor:'#0284c7cc',borderColor:'#0284c7',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'周环比(%)',data:guides.map(g=>g.wow),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:5,pointBackgroundColor:'#dc2626',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label==='周销售')return '周销售: '+fmtAmt(ctx.raw);if(ctx.dataset.label==='WOS')return 'WOS: '+fmtAmt(ctx.raw);return '周环比: '+(isNA(ctx.raw)?'N/A':(ctx.raw>=0?'+':'')+ctx.raw.toFixed(1)+'%')}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'金额(¥)'},ticks:{callback:v=>fmtAmt(v)}},x1:{position:'top',title:{display:true,text:'周环比(%)'},ticks:{callback:v=>v+'%'},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Chart 8: Guide JD + ATV (merged)
    charts.guideJdKdj=new Chart(document.getElementById('chartGuideJdKdj'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'连带率',data:guides.map(g=>g.wk_jd),backgroundColor:'#059669cc',borderColor:'#059669',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'客单价(¥)',data:guides.map(g=>g.wk_atv),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#dc2626',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label==='连带率')return '连带率: '+ctx.raw.toFixed(2);return '客单价: ¥'+Math.round(ctx.raw)}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'连带率'},ticks:{callback:v=>v.toFixed(1)}},x1:{position:'top',title:{display:true,text:'客单价(¥)'},ticks:{callback:v=>'¥'+v},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Chart 9: Guide long-tail (amount + qty)
    charts.guideLt=new Chart(document.getElementById('chartGuideLt'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'长尾款金额',data:guides.map(g=>g.wk_lt),backgroundColor:'#7c3aedcc',borderColor:'#7c3aed',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'长尾款件数',data:guides.map(g=>g.wk_lt_qty),borderColor:'#ea580c',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#ea580c',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label.includes('金额'))return '长尾款: '+fmtAmt(ctx.raw);return '件数: '+ctx.raw+'件'}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'金额(¥)'},ticks:{callback:v=>fmtAmt(v)}},x1:{position:'top',title:{display:true,text:'件数'},ticks:{callback:v=>v+'件'},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Guide table
    const _wkJdArr=guides.filter(g=>g.wk_cnt>0).map(g=>g.wk_jd);
    const wkJdAvg=_wkJdArr.length?_wkJdArr.reduce((a,b)=>a+b,0)/_wkJdArr.length:null;
    let gHtml='<thead><tr><th>导购</th><th>职务</th><th>周销售</th><th>周环比</th><th>笔数</th><th>件数</th><th>客单价</th><th>连带率</th><th>WOS</th><th>长尾款</th><th>长尾件数</th><th style="color:#db2777">新品</th></tr></thead><tbody>';
    guides.forEach(g=>{
      const wb=isNA(g.wow)?'<span class="badge gray">N/A</span>':(g.wow>=0?'<span class="badge green">+'+g.wow.toFixed(1)+'%</span>':'<span class="badge red">'+g.wow.toFixed(1)+'%</span>');
      gHtml+='<tr><td class="store-name" style="color:'+color+'">'+g.name+'</td><td style="font-size:11px;color:#64748b">'+g.title+'</td><td style="font-weight:700">'+fmtAmt(g.wk_sales)+'</td><td>'+wb+'</td><td>'+g.wk_cnt+'</td><td>'+g.wk_qty+'件</td><td>¥'+Math.round(g.wk_atv)+'</td><td'+(wkJdAvg!==null&&g.wk_cnt>0&&g.wk_jd<wkJdAvg?' style="color:#dc2626;font-weight:700"':'')+'>'+g.wk_jd.toFixed(2)+'</td><td>'+(g.wk_wos>0?fmtAmt(g.wk_wos):'<span style="color:#94a3b8">-</span>')+'</td><td>'+(g.wk_lt>0?fmtAmt(g.wk_lt):'<span style="color:#94a3b8">-</span>')+'</td><td>'+g.wk_lt_qty+'件</td><td style="color:#db2777">'+(g.wk_np>0?fmtAmt(g.wk_np):'<span style="color:#94a3b8">-</span>')+'</td></tr>';
    });
    const tS=guides.reduce((a,g)=>a+g.wk_sales,0),tC=guides.reduce((a,g)=>a+g.wk_cnt,0),tQ=guides.reduce((a,g)=>a+g.wk_qty,0),tW=guides.reduce((a,g)=>a+g.wk_wos,0),tL=guides.reduce((a,g)=>a+g.wk_lt,0),tLQ=guides.reduce((a,g)=>a+g.wk_lt_qty,0),tNP=guides.reduce((a,g)=>a+(g.wk_np||0),0);
    const tW28=guides.reduce((a,g)=>a+(g.wk28_sales||0),0);
    const tWow=tW28>0?((tS-tW28)/tW28*100):null;
    const tWb=isNA(tWow)?'<span class="badge gray">N/A</span>':(tWow>=0?'<span class="badge green">+'+tWow.toFixed(1)+'%</span>':'<span class="badge red">'+tWow.toFixed(1)+'%</span>');
    gHtml+='<tr style="background:#05966915;font-weight:700"><td class="store-name">合计</td><td>-</td><td>'+fmtAmt(tS)+'</td><td>'+tWb+'</td><td>'+tC+'</td><td>'+tQ+'件</td><td>¥'+Math.round(tS/Math.max(tC,1))+'</td><td>'+(tQ/Math.max(tC,1)).toFixed(2)+'</td><td>'+(tW>0?fmtAmt(tW):'-')+'</td><td>'+(tL>0?fmtAmt(tL):'-')+'</td><td>'+tLQ+'件</td><td style="color:#db2777">'+(tNP>0?fmtAmt(tNP):'-')+'</td></tr>';
    gHtml+='</tbody>';
    document.getElementById('guideTable').innerHTML=gHtml;
    if(wkJdAvg!==null){const _note=document.createElement('div');_note.style.cssText='font-size:11px;color:#dc2626;margin-top:4px';_note.textContent='🔴 本周连带率低于本店导购均值（'+wkJdAvg.toFixed(2)+'）已标红';document.getElementById('guideTable').parentNode.appendChild(_note);}
    const _gAvg=tS/(guides.length||1);
    const _gBelow=guides.filter(g=>g.wk_sales>0 && g.wk_sales<_gAvg*0.7);
    const _gTop=guides.slice().sort((a,b)=>b.wk_sales-a.wk_sales)[0];
    let _gConcNote=document.createElement('div');
    _gConcNote.style.cssText='margin-top:10px';
    _gConcNote.innerHTML=conc('本店'+guides.length+'名导购，周销售合计'+fmtAmt(tS)+'；TOP导购'+(_gTop?_gTop.name+'（'+fmtAmt(_gTop.wk_sales)+'）':'-')+'领先。'+(wkJdAvg!==null?('其中'+guides.filter(g=>g.wk_cnt>0&&g.wk_jd<wkJdAvg).length+'人连带率低于店均'+wkJdAvg.toFixed(2)+'，已标红。'):''),_gBelow.length>0?('低于店均周销售（¥'+Math.round(_gAvg)+'）的导购：'+_gBelow.map(g=>g.name+'（¥'+Math.round(g.wk_sales)+'）').join('、')+'，需加强一对一销售辅导与排班优化。'):'导购团队周销售分布均衡，保持当前带教节奏。');
    document.getElementById('guideTable').parentNode.appendChild(_gConcNote);
  }
}
"""

_js_out = JS_RENDER_STORE.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}').replace('ZK29','WK29')

with open(OUTPUT, 'a', encoding='utf-8') as f:
    f.write(_js_out)

import os
print(f"Part 3 (renderStore) appended. Size: {os.path.getsize(OUTPUT)} bytes")
