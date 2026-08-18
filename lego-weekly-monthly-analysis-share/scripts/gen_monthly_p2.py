#!/usr/bin/env python3
"""Part 2: renderStore() function with all charts and tables"""

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
_month_title = f'{REPORT_MONTH}月'
_prev_range = f'{_prev_month}月1-{_prev_end_day}日'
_curr_range = f'{REPORT_MONTH}月1-{REPORT_DAY}日'
print(f"  monthly p2: {_curr_range} vs {_prev_range}")

HTML = r"""
function renderStore(store){
  const d=DATA[store],region=DATA['_region'],color=STORE_COLORS[store],daily=d.daily;
  const cvo=(DATA['_conversion']||{});
  const cvoRate=(cvo[store]?cvo[store].rate:null);
  const cvoRegionRate=(cvo._region?cvo._region.rate:null);
  let html='<div class="kpi-row">'+
    '<div class="kpi-card"><div class="label">月累计销售</div><div class="value" style="color:'+color+'">'+fmtAmt(d.month_amt)+'</div><div class="sub">月目标'+fmtAmt(d.target)+' | 完成率'+d.month_rate.toFixed(1)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">月完成率</div><div class="value '+(d.month_rate>=50?'green':'red')+'">'+d.month_rate.toFixed(1)+'%</div><div class="sub">区域'+region.month_rate.toFixed(1)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">月环比</div><div class="value '+(isNA(d.mom)?'':colorPct(d.mom))+'">'+(isNA(d.mom)?'N/A':(d.mom>=0?'+':'')+d.mom.toFixed(1)+'%')+'</div><div class="sub">vs 6月1-18日</div></div>'+
    '<div class="kpi-card"><div class="label">月同比</div><div class="value '+(isNA(d.yoy)?'':colorPct(d.yoy))+'">'+(isNA(d.yoy)?'N/A':fmtPct(d.yoy))+'</div><div class="sub">'+(isNA(d.yoy)?'新店无同期':'vs去年同期')+'</div></div>'+
    '<div class="kpi-card"><div class="label">月连带率</div><div class="value blue">'+d.month_jd.toFixed(2)+'</div><div class="sub">区域'+region.month_jd.toFixed(2)+'</div></div>'+
    '<div class="kpi-card"><div class="label">月客单价</div><div class="value blue">¥'+Math.round(d.month_atv)+'</div><div class="sub">区域¥'+Math.round(region.month_atv)+'</div></div>'+
    '<div class="kpi-card"><div class="label">月成交率</div><div class="value '+(cvoRate===null?'gray':(cvoRate>=cvoRegionRate?'green':'red'))+'">'+(cvoRate===null?'缺客流':cvoRate.toFixed(2)+'%')+'</div><div class="sub">区域'+(cvoRegionRate===null?'N/A':cvoRegionRate.toFixed(2))+'% | '+(cvo[store]?cvo[store].cnt:'0')+'笔/'+ (cvo[store]?cvo[store].flow:'0') +'客流</div></div>'+
    '<div class="kpi-card"><div class="label">长尾款月累计</div><div class="value purple">'+fmtAmt(d.month_lt)+'</div><div class="sub">'+d.month_lt_qty+'件 | 占比'+(d.month_amt>0?(d.month_lt/d.month_amt*100).toFixed(1):0)+'%</div></div>'+
    '<div class="kpi-card"><div class="label">新品月累计</div><div class="value" style="color:#db2777">'+fmtAmt(d.new_amt||0)+'</div><div class="sub">'+(d.new_qty||0)+'件 | 占比'+(d.new_rate||0).toFixed(1)+'%</div></div>'+
    '</div>';

  html+=concKPI(store);

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📊</span> 产品系列结构（7月累计）</h3><div class="chart-wrap"><canvas id="chartSeriesPie"></canvas></div>'+concSeries(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🎯</span> 门店月维度能力雷达图（vs区域均值）</h3><div class="chart-wrap"><canvas id="chartRadar"></canvas></div>'+concRadar(store)+'</div></div>';

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">📊</span> 月环比（7月 vs 6月同期）</h3><div class="chart-wrap"><canvas id="chartMom"></canvas></div>'+concMom(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🏆</span> TOP10产品月销排名（含件数）</h3><div class="chart-wrap"><canvas id="chartTopProducts"></canvas></div>'+concTop(store)+'</div></div>';

  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">🆕</span> 新品TOP销售排名（上市日期=本月）</h3><div class="chart-wrap"><canvas id="chartNewProducts"></canvas></div>'+concNew(store)+'</div></div>';

  const totWos=daily.reduce((a,x)=>a+x.wos,0),totJl=daily.reduce((a,x)=>a+x.jl,0);
  const wosPct=d.month_amt>0?(totWos/d.month_amt*100):0,jlPct=d.month_amt>0?(totJl/d.month_amt*100):0,bothPct=d.month_amt>0?((totWos+totJl)/d.month_amt*100):0;
  const rTotWos=regVal(x=>x.month_wos),rTotJl=regVal(x=>x.month_jl),rTot=regVal(x=>x.month_amt);
  const rBothPct=rTot>0?((rTotWos+rTotJl)/rTot*100):0;
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">⚡</span> WOS & 即时零售 月累计对比（vs区域）</h3><div class="chart-wrap"><canvas id="chartWosJl"></canvas></div>'+
    conc(store+'本月WOS '+fmtAmt(totWos)+'（占比'+wosPct.toFixed(1)+'%），即时零售 '+fmtAmt(totJl)+'（占比'+jlPct.toFixed(1)+'%），两渠道合计占比'+bothPct.toFixed(1)+'%。',bothPct<rBothPct?'WOS+即时零售合计占比低于区域（'+rBothPct.toFixed(1)+'%），建议加大线上渠道活动投放，强化美团/饿了么专属优惠引流。':bothPct>rBothPct+3?'线上渠道占比高于区域，多渠道运营成熟，建议保持投放策略并关注转化效率。':'线上渠道占比与区域持平，适度增加即时零售平台曝光。')+'</div></div>';

  html+='<div class="chart-grid">'+
    '<div class="chart-card"><h3><span class="icon">👩</span> 导购月销售+WOS+月环比</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideSales"></canvas></div>'+concGuideSales(store)+'</div>'+
    '<div class="chart-card"><h3><span class="icon">🔄</span> 导购连带率+客单价</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideJdKdj"></canvas></div>'+concGuideJdKdj(store)+'</div></div>';

  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">📦</span> 导购长尾款分析（金额/件数）</h3><div class="chart-wrap tall" style="height:400px;"><canvas id="chartGuideLt"></canvas></div>'+concGuideLt(store)+'</div></div>';
  html+='<div class="chart-grid full"><div class="chart-card"><h3><span class="icon">📋</span> 导购月维度明细</h3><table class="summary-table" id="guideTable"></table></div></div>';

  document.getElementById('content').innerHTML=html;
  Object.values(charts).forEach(c=>c&&c.destroy());charts={};

  // Chart: Series Pie
  const sd=(DATA['_series']||{})[store]||{};
  const se=Object.entries(sd).sort((a,b)=>b[1]-a[1]);
  const top10=se.slice(0,10),oth=se.slice(10).reduce((a,[k,v])=>a+v,0);
  charts.seriesPie=new Chart(document.getElementById('chartSeriesPie'),{
    type:'doughnut',
    data:{labels:top10.map(([k])=>k).concat(oth>0?['其他']:[]),datasets:[{data:top10.map(([k,v])=>v).concat(oth>0?[oth]:[]),backgroundColor:SC.slice(0,top10.length+(oth>0?1:0)),borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{font:{size:11},padding:6}},tooltip:{callbacks:{label:ctx=>{const pct=d.month_amt>0?(ctx.raw/d.month_amt*100).toFixed(1):0;return ctx.label+': '+fmtAmt(ctx.raw)+' ('+pct+'%)'}}}},cutout:'55%'}
  });

  // Chart: Radar
  const atv=d.month_atv,jd=d.month_jd,cnt=d.month_cnt;
  const wpct=d.month_amt>0?d.month_wos/d.month_amt*100:0,lpct=d.month_amt>0?d.month_lt/d.month_amt*100:0;
  const ratv=region.month_atv,rjd=region.month_jd,rcnt=region.month_cnt/7;
  const rwpct=region.month_amt>0?region.month_wos/region.month_amt*100:0,rlpct=region.month_amt>0?region.month_lt/region.month_amt*100:0;
  const rIdx=(sv,rv)=>rv>0?sv/rv*100:0;
  const rYoyIdx=(sy,ry)=>{if(isNA(sy))return 0;if(isNA(ry))return 100;return 100+(sy-ry)*3};
  const rL=['月同比','月客单价','月连带率','月笔数','WOS占比','长尾款占比'];
  const rRS=[d.yoy,atv,jd,cnt,wpct,lpct],rRR=[region.yoy,ratv,rjd,rcnt,rwpct,rlpct];
  const rF=['%','¥','','笔','%','%'];
  const sR=[rYoyIdx(d.yoy,region.yoy),rIdx(atv,ratv),rIdx(jd,rjd),rIdx(cnt,rcnt),rIdx(wpct,rwpct),rIdx(lpct,rlpct)];
  charts.radar=new Chart(document.getElementById('chartRadar'),{
    type:'radar',
    data:{labels:rL,datasets:[
      {label:store,data:sR,borderColor:color,backgroundColor:color+'30',borderWidth:2,pointRadius:4},
      {label:'区域均值(=100)',data:[100,100,100,100,100,100],borderColor:'#64748b',backgroundColor:'#64748b15',borderWidth:2,pointRadius:3,borderDash:[4,4]}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>{const i=ctx.dataIndex,u=rF[i],sv=rRS[i],rv=rRR[i];const sStr=isNA(sv)?'N/A':(u==='¥'?'¥'+Math.round(sv):(u==='%'?sv.toFixed(1)+'%':(u==='笔'?sv+'笔':sv.toFixed(2))));const rStr=isNA(rv)?'N/A':(u==='¥'?'¥'+Math.round(rv):(u==='%'?rv.toFixed(1)+'%':(u==='笔'?Math.round(rv)+'笔':rv.toFixed(2))));return ctx.dataset.label+': 指数'+ctx.raw.toFixed(0)+' | 门店'+sStr+' | 区域'+rStr;}}}},
      scales:{r:{beginAtZero:true,suggestedMin:0,suggestedMax:200,ticks:{stepSize:50}}}}
  });

  // Chart 5: MoM (7月 vs 6月同期)
  const mAmt=d.month_amt,mPrev=d.mom_prev||0;
  charts.mom=new Chart(document.getElementById('chartMom'),{
    type:'bar',
    data:{labels:['6月1-18日','7月1-18日'],datasets:[{label:'月销售',data:[mPrev,mAmt],backgroundColor:[mPrev>=mAmt?'#16a34a80':'#94a3b880',mAmt>=mPrev?color+'cc':'#dc262680'],borderColor:['#64748b',color],borderWidth:2,borderRadius:8,barPercentage:.5}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>'月销售: '+fmtAmt(ctx.raw),afterLabel:ctx=>{if(ctx.dataIndex===1&&mPrev>0){const diff=((mAmt-mPrev)/mPrev*100);return '月环比: '+(diff>=0?'+':'')+diff.toFixed(1)+'%'}return ''}}}},
      scales:{y:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart 6: TOP10 Products (with qty)
  const pd=(DATA['_products']||{})[store]||{};
  const pEnts=Object.entries(pd).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  charts.topProducts=new Chart(document.getElementById('chartTopProducts'),{
    type:'bar',
    data:{labels:pEnts.map(([k])=>k),datasets:[{label:'月销售',data:pEnts.map(([k,a,q])=>a),backgroundColor:SC.slice(0,pEnts.length).map(c=>c+'cc'),borderColor:SC.slice(0,pEnts.length),borderWidth:1,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const idx=ctx.dataIndex;const [n,a,q]=pEnts[idx];const pct=d.month_amt>0?(a/d.month_amt*100).toFixed(1):0;return n+': '+fmtAmt(a)+' ('+pct+'%, '+q+'件)'}}}},
      scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
  });

  // Chart: New products TOP (上市日期=本月)
  const npDetail=(DATA['_new_products_detail']||{})[store]||[];
  const npCanvas=document.getElementById('chartNewProducts');
  if(npDetail.length>0){
    charts.newProducts=new Chart(npCanvas,{
      type:'bar',
      data:{labels:npDetail.map(x=>x.name),datasets:[{label:'新品销售',data:npDetail.map(x=>x.amt),backgroundColor:SC.slice(0,npDetail.length).map(c=>c+'cc'),borderColor:SC.slice(0,npDetail.length),borderWidth:1,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const x=npDetail[ctx.dataIndex];const pct=d.month_amt>0?(x.amt/d.month_amt*100).toFixed(1):0;return x.name+': '+fmtAmt(x.amt)+' ('+pct+'%, '+x.qty+'件)';}}}},
        scales:{x:{ticks:{callback:v=>fmtAmt(v)}}}}
    });
  }else{
    npCanvas.parentNode.innerHTML='<div style="padding:40px;text-align:center;color:#94a3b8">本月无新品销售记录</div>';
  }

  // Chart: WOS & 即时零售 monthly comparison (store vs region)
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

    // Chart 7: Guide sales + WOS + MoM (merged)
    charts.guideSales=new Chart(document.getElementById('chartGuideSales'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'月销售',data:guides.map(g=>g.m_sales),backgroundColor:guides.map(g=>g.m_sales>0?color+'cc':'#e2e8f0'),borderColor:color,borderWidth:1,borderRadius:4,xAxisID:'x',order:3},
        {type:'bar',label:'WOS',data:guides.map(g=>g.m_wos),backgroundColor:'#0284c7cc',borderColor:'#0284c7',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'月环比(%)',data:guides.map(g=>g.mom),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:5,pointBackgroundColor:'#dc2626',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label==='月销售')return '月销售: '+fmtAmt(ctx.raw);if(ctx.dataset.label==='WOS')return 'WOS: '+fmtAmt(ctx.raw);return '月环比: '+(isNA(ctx.raw)?'N/A':(ctx.raw>=0?'+':'')+ctx.raw.toFixed(1)+'%')}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'金额(¥)'},ticks:{callback:v=>fmtAmt(v)}},x1:{position:'top',title:{display:true,text:'月环比(%)'},ticks:{callback:v=>v+'%'},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Chart 8: Guide JD + ATV (merged)
    charts.guideJdKdj=new Chart(document.getElementById('chartGuideJdKdj'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'连带率',data:guides.map(g=>g.m_jd),backgroundColor:'#059669cc',borderColor:'#059669',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'客单价(¥)',data:guides.map(g=>g.m_atv),borderColor:'#dc2626',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#dc2626',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label==='连带率')return '连带率: '+ctx.raw.toFixed(2);return '客单价: ¥'+Math.round(ctx.raw)}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'连带率'},ticks:{callback:v=>v.toFixed(1)}},x1:{position:'top',title:{display:true,text:'客单价(¥)'},ticks:{callback:v=>'¥'+v},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Chart 9: Guide long-tail (amount + qty)
    charts.guideLt=new Chart(document.getElementById('chartGuideLt'),{
      data:{labels:gLabels,datasets:[
        {type:'bar',label:'长尾款金额',data:guides.map(g=>g.m_lt),backgroundColor:'#7c3aedcc',borderColor:'#7c3aed',borderWidth:1,borderRadius:4,xAxisID:'x',order:2},
        {type:'line',label:'长尾款件数',data:guides.map(g=>g.m_lt_qty),borderColor:'#ea580c',backgroundColor:'transparent',borderWidth:2,pointRadius:6,pointBackgroundColor:'#ea580c',pointBorderColor:'#fff',pointBorderWidth:2,tension:.2,xAxisID:'x1',order:1}
      ]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{position:'top',labels:{font:{size:12}}},tooltip:{callbacks:{label:ctx=>{if(ctx.dataset.label.includes('金额'))return '长尾款: '+fmtAmt(ctx.raw);return '件数: '+ctx.raw+'件'}}}},
        scales:{x:{position:'bottom',title:{display:true,text:'金额(¥)'},ticks:{callback:v=>fmtAmt(v)}},x1:{position:'top',title:{display:true,text:'件数'},ticks:{callback:v=>v+'件'},grid:{drawOnChartArea:false}},y:{ticks:{font:{size:13,weight:'600'}}}}}
    });

    // Guide table
    const _mJdArr=guides.filter(g=>g.m_cnt>0).map(g=>g.m_jd);
    const mJdAvg=_mJdArr.length?_mJdArr.reduce((a,b)=>a+b,0)/_mJdArr.length:null;
    let gHtml='<thead><tr><th>导购</th><th>职务</th><th>月销售</th><th>月目标</th><th>月完成率</th><th>月环比</th><th>笔数</th><th>件数</th><th>客单价</th><th>连带率</th><th>WOS</th><th>长尾款</th><th>长尾件数</th><th style="color:#db2777">新品</th></tr></thead><tbody>';
    guides.forEach(g=>{
      const mb=isNA(g.mom)?'<span class="badge gray">N/A</span>':(g.mom>=0?'<span class="badge green">+'+g.mom.toFixed(1)+'%</span>':'<span class="badge red">'+g.mom.toFixed(1)+'%</span>');
      gHtml+='<tr><td class="store-name" style="color:'+color+'">'+g.name+'</td><td style="font-size:11px;color:#64748b">'+g.title+'</td><td style="font-weight:700">'+fmtAmt(g.m_sales)+'</td><td>'+fmtAmt(g.target)+'</td><td style="color:'+(g.m_rate>=50?'#16a34a':'#dc2626')+';font-weight:700">'+g.m_rate.toFixed(1)+'%</td><td>'+mb+'</td><td>'+g.m_cnt+'</td><td>'+g.m_qty+'件</td><td>¥'+Math.round(g.m_atv)+'</td><td'+(mJdAvg!==null&&g.m_cnt>0&&g.m_jd<mJdAvg?' style="color:#dc2626;font-weight:700"':'')+'>'+g.m_jd.toFixed(2)+'</td><td>'+(g.m_wos>0?fmtAmt(g.m_wos):'<span style="color:#94a3b8">-</span>')+'</td><td>'+(g.m_lt>0?fmtAmt(g.m_lt):'<span style="color:#94a3b8">-</span>')+'</td><td>'+g.m_lt_qty+'件</td><td style="color:#db2777">'+(g.m_np>0?fmtAmt(g.m_np):'<span style="color:#94a3b8">-</span>')+'</td></tr>';
    });
    const tS=guides.reduce((a,g)=>a+g.m_sales,0),tT=guides.reduce((a,g)=>a+g.target,0),tC=guides.reduce((a,g)=>a+g.m_cnt,0),tQ=guides.reduce((a,g)=>a+g.m_qty,0),tW=guides.reduce((a,g)=>a+g.m_wos,0),tL=guides.reduce((a,g)=>a+g.m_lt,0),tLQ=guides.reduce((a,g)=>a+g.m_lt_qty,0),tNP=guides.reduce((a,g)=>a+(g.m_np||0),0);
    const tP=guides.reduce((a,g)=>a+(g.m_prev||0),0);
    const tMom=tP>0?((tS-tP)/tP*100):null;
    const tMb=isNA(tMom)?'<span class="badge gray">N/A</span>':(tMom>=0?'<span class="badge green">+'+tMom.toFixed(1)+'%</span>':'<span class="badge red">'+tMom.toFixed(1)+'%</span>');
    gHtml+='<tr style="background:#05966915;font-weight:700"><td class="store-name">合计</td><td>-</td><td>'+fmtAmt(tS)+'</td><td>'+fmtAmt(tT)+'</td><td style="color:'+(tT>0&&tS/tT*100>=50?'#16a34a':'#dc2626')+'">'+(tT>0?(tS/tT*100).toFixed(1):0)+'%</td><td>'+tMb+'</td><td>'+tC+'</td><td>'+tQ+'件</td><td>¥'+Math.round(tS/Math.max(tC,1))+'</td><td>'+(tQ/Math.max(tC,1)).toFixed(2)+'</td><td>'+(tW>0?fmtAmt(tW):'-')+'</td><td>'+(tL>0?fmtAmt(tL):'-')+'</td><td>'+tLQ+'件</td><td style="color:#db2777">'+(tNP>0?fmtAmt(tNP):'-')+'</td></tr>';
    gHtml+='</tbody>';
    document.getElementById('guideTable').innerHTML=gHtml;
    if(mJdAvg!==null){const _note=document.createElement('div');_note.style.cssText='font-size:11px;color:#dc2626;margin-top:4px';_note.textContent='🔴 本月连带率低于本店导购均值（'+mJdAvg.toFixed(2)+'）已标红';document.getElementById('guideTable').parentNode.appendChild(_note);}
    const _gAvg=tS/(guides.length||1);
    const _gBelow=guides.filter(g=>g.m_sales>0 && g.m_sales<_gAvg*0.7);
    const _gLow=guides.filter(g=>g.m_rate<30 && g.m_sales>0);
    const _gTop=guides.slice().sort((a,b)=>b.m_sales-a.m_sales)[0];
    let _gConcNote=document.createElement('div');
    _gConcNote.style.cssText='margin-top:10px';
    _gConcNote.innerHTML=conc('本店'+guides.length+'名导购，月销售合计'+fmtAmt(tS)+'；TOP导购'+(_gTop?_gTop.name+'（'+fmtAmt(_gTop.m_sales)+'）':'-')+'领先。'+(mJdAvg!==null?('其中'+guides.filter(g=>g.m_cnt>0&&g.m_jd<mJdAvg).length+'人连带率低于店均'+mJdAvg.toFixed(2)+'，已标红。'):''),(_gBelow.length>0?('低于店均月销售（¥'+Math.round(_gAvg)+'）的导购：'+_gBelow.map(g=>g.name+'（¥'+Math.round(g.m_sales)+'）').join('、')+'，需加强一对一销售辅导。'):'')+(_gLow.length>0?(_gLow.map(g=>g.name).join('、')+'月完成率不足30%，需重点帮扶。'):'导购团队月销售分布均衡，保持当前带教节奏。'));
    document.getElementById('guideTable').parentNode.appendChild(_gConcNote);
  }
}
"""

# 替换硬编码日期为动态值
HTML = HTML.replace('6月1-18日', _prev_range).replace('7月1-18日', _curr_range)
HTML = HTML.replace('7月累计', f'{_month_title}累计')
HTML = HTML.replace('7月 vs 6月同期', f'{_month_title} vs {_prev_month}月同期')

with open('/Users/a123/WorkBuddy/Claw/outputs/monthly_analysis.html', 'a', encoding='utf-8') as f:
    f.write(HTML)

import os
sz = os.path.getsize('/Users/a123/WorkBuddy/Claw/outputs/monthly_analysis.html')
print(f"Part 2 appended. File size: {sz:,} bytes")
