#!/usr/bin/env python3
"""Append JS constants + helpers + conclusion functions to weekly HTML"""

import os
import sys
import glob
import pandas as pd

OUTPUT = '/Users/a123/WorkBuddy/Claw/outputs/weekly_analysis.html'

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
# 生成WEEKDAYS数组
_wd_items = []
for i in range((max_date.date() - WK_START.date()).days + 1):
    _d = WK_START + pd.Timedelta(days=i)
    _wd_items.append(f"'{_d.month}/{_d.day}({_DAY_NAMES[_d.dayofweek]})'")
WEEKDAYS_JS = f"const WEEKDAYS=[{_wd_items.join(',') if hasattr(_wd_items, 'join') else ','.join(_wd_items)}];"

print(f"  max_date={max_date.strftime('%Y-%m-%d')}, WK{WK_NUM_INT} (prev WK{WK_PREV_NUM}), weekdays={len(_wd_items)}天")

JS_PART1 = r"""
const STORE_ORDER=['扬州万象汇','扬州京华城','扬州江都金鹰','泰州万象城','宿迁宝龙','淮安新亚','淮安万象城'];
""" + WEEKDAYS_JS + r"""
const STORE_COLORS={'扬州万象汇':'#7c3aed','扬州京华城':'#0ea5e9','扬州江都金鹰':'#059669','泰州万象城':'#dc2626','宿迁宝龙':'#ea580c','淮安新亚':'#0891b2','淮安万象城':'#9333ea'};
const SC=['#7c3aed','#0ea5e9','#059669','#dc2626','#ea580c','#0891b2','#9333ea','#d97706','#65a30d','#4f46e5','#e11d48','#0d9488','#c026d3','#7c2d12','#1e40af','#831843','#365314','#854d0e','#3730a3','#9f1239'];
let charts={};
let currentStore='区域对比';

function fmtAmt(v){if(v===null||v===undefined)return 'N/A';return '¥'+Math.round(v).toLocaleString()}
function fmtPct(v){if(v===null||v===undefined)return 'N/A';return v.toFixed(1)+'%'}
function colorPct(v){if(v===null||v===undefined)return '';return v>=0?'green':'red'}
function isNA(v){return v===null||v===undefined}
function getWeeklyCnt(s){if(typeof s==='string')s=DATA[s];return s.daily.reduce((a,x)=>a+x.tx_cnt,0)}
function getWeeklyATV(s){if(typeof s==='string')s=DATA[s];const c=getWeeklyCnt(s);return c>0?s.wk_amt/c:0}
function getWeeklyJD(s){if(typeof s==='string')s=DATA[s];const c=getWeeklyCnt(s);if(c===0)return 0;return s.daily.reduce((a,x)=>a+x.jd*x.tx_cnt,0)/c}
function getRegWkATV(){const t=STORE_ORDER.reduce((a,s)=>a+DATA[s].wk_amt,0);const c=STORE_ORDER.reduce((a,s)=>a+getWeeklyCnt(s),0);return c>0?t/c:0}
function getRegWkJD(){const c=STORE_ORDER.reduce((a,s)=>a+getWeeklyCnt(s),0);if(c===0)return 0;return STORE_ORDER.reduce((a,s)=>a+getWeeklyJD(s)*getWeeklyCnt(s),0)/c}
function regVal(fn){return STORE_ORDER.reduce((a,s)=>a+fn(DATA[s]),0)}
function conc(c,o){if(!o)return '<div class="conclusion"><div class="conc-title">📋 结论</div><div class="conc-body">'+c+'</div></div>';return '<div class="conclusion"><div class="conc-title">📋 结论</div><div class="conc-body">'+c+'</div><div class="conc-opportunity"><span class="conc-label">💡 需关注的机会点</span><br>'+o+'</div></div>'}

function concKPI(s){
  const d=DATA[s],r=DATA['_region'];
  const atv=getWeeklyATV(s),jd=getWeeklyJD(s),ratv=getRegWkATV(),rjd=getRegWkJD();
  const wpct=d.wk_amt>0?d.wk_wos/d.wk_amt*100:0,rwpct=regVal(x=>x.wk_wos)/(regVal(x=>x.wk_amt)||1)*100;
  const lpct=d.wk_amt>0?d.wk_lt/d.wk_amt*100:0,rlpct=regVal(x=>x.wk_lt)/(regVal(x=>x.wk_amt)||1)*100;
  const wt=d.daily.reduce((a,x)=>a+(x.target||0),0),rate=wt>0?d.wk_amt/wt*100:0;
  const rwt=regVal(x=>x.daily.reduce((a,day)=>a+(day.target||0),0)),rrate=rwt>0?regVal(x=>x.wk_amt)/rwt*100:0;
  const below=[];
  if(rate<rrate)below.push('周完成率'+rate.toFixed(1)+'%（区域'+rrate.toFixed(1)+'%，差'+(rrate-rate).toFixed(1)+'pct）');
  if(!isNA(d.wk_yoy)&&!isNA(r.wk_yoy)&&d.wk_yoy<r.wk_yoy)below.push('周同比'+d.wk_yoy.toFixed(1)+'%（区域'+r.wk_yoy.toFixed(1)+'%）');
  if(!isNA(d.wow)&&!isNA(r.wow)&&d.wow<r.wow)below.push('周环比'+d.wow.toFixed(1)+'%（区域'+r.wow.toFixed(1)+'%）');
  if(atv<ratv)below.push('客单价¥'+Math.round(atv)+'（区域¥'+Math.round(ratv)+'，差¥'+Math.round(ratv-atv)+'）');
  if(jd<rjd)below.push('连带率'+jd.toFixed(2)+'（区域'+rjd.toFixed(2)+'，差'+(rjd-jd).toFixed(2)+'）');
  if(wpct<rwpct)below.push('WOS占比'+wpct.toFixed(1)+'%（区域'+rwpct.toFixed(1)+'%）');
  if(lpct<rlpct)below.push('长尾款占比'+lpct.toFixed(1)+'%（区域'+rlpct.toFixed(1)+'%）');
  const c=s+'KPI汇总：周销售'+fmtAmt(d.wk_amt)+'，完成率'+rate.toFixed(1)+'%，连带率'+jd.toFixed(2)+'，客单价¥'+Math.round(atv)+'。';
  let o='';
  if(below.length>0)o='低于区域整体的KPI：'+below.join('；')+'。需逐项分析差距原因，制定提升计划。';
  else o='各项KPI均达到或优于区域整体水平，保持当前运营策略。';
  return conc(c,o);
}

function concSeries(s){
  const d=DATA[s];const sd=(DATA['_series']||{})[s]||{};
  const ents=Object.entries(sd).sort((a,b)=>b[1]-a[1]);
  if(ents.length===0)return conc('暂无系列数据。');
  const total=d.wk_amt||1;
  const top3=ents.slice(0,3);
  const pct=top3.reduce((a,[k,v])=>a+v,0)/total*100;
  // 每个系列的销售占比
  const topPcts=ents.slice(0,10).map(([k,v])=>k+'('+fmtAmt(v)+',占比'+(v/total*100).toFixed(1)+'%)');
  const c='TOP10系列：'+topPcts.join('、')+'；TOP3合计占比'+pct.toFixed(1)+'%。';
  let o='';
  if(pct>70)o='TOP3系列过于集中，建议丰富中长尾系列推荐，降低单一系列依赖风险。';else if(pct<35)o='系列分布较散，建议打造1-2个主推核心系列进行重点陈列。';else o='系列结构良好，保持均衡推荐同时打造核心系列。';
  // 关注重点系列排名：CITY/TECHNIC/NINJAGO 是否在TOP3，MINECRAFT 是否在TOP5
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
  const botanicalStr=botanicalAmt>0?'BOTANICALS植物系列占比'+botanicalPct.toFixed(1)+'%。':'BOTANICALS植物系列本周无销售。';
  // 结论中加入BOTANICALS占比
  const c2=c+botanicalStr;
  if(botanicalAmt>0){
    o+='BOTANICALS植物系列销售占比'+botanicalPct.toFixed(1)+'%（指标4.1%）';
    if(botanicalPct<4.1)o+='，低于指标，需加强植物系列陈列和推荐，抓住礼品/家居装饰客群。';
    else o+='，达到指标。';
  }else{
    o+='BOTANICALS植物系列本周无销售，低于指标4.1%，建议关注植物系列铺货和推荐。';
  }
  return conc(c2,o);
}
function concRadar(s){
  const d=DATA[s];const atv=getWeeklyATV(s),jd=getWeeklyJD(s),cnt=getWeeklyCnt(s);
  const ratv=getRegWkATV(),rjd=getRegWkJD(),rcnt=regVal(getWeeklyCnt)/7;
  const wpct=d.wk_amt>0?d.wk_wos/d.wk_amt*100:0,lpct=d.wk_amt>0?d.wk_lt/d.wk_amt*100:0;
  const rwpct=regVal(x=>x.wk_wos)/(regVal(x=>x.wk_amt)||1)*100,rlpct=regVal(x=>x.wk_lt)/(regVal(x=>x.wk_amt)||1)*100;
  const items=[['客单价',atv>ratv],['连带率',jd>rjd],['笔数',cnt>rcnt],['WOS占比',wpct>rwpct],['长尾款占比',lpct>rlpct]];
  const lead=items.filter(x=>x[1]).map(x=>x[0]),weak=items.filter(x=>!x[1]).map(x=>x[0]);
  const c=s+'在'+(lead.length>0?lead.join('、'):'无')+'维度领先区域'+(weak.length>0?'，'+weak.join('、')+'维度低于区域需提升':'，各维度均优于区域，综合能力突出')+'。';
  let o='';
  if(weak.length>0){
    const gaps=[];
    if(weak.includes('客单价'))gaps.push('客单价差¥'+Math.round(ratv-atv)+'（¥'+Math.round(atv)+' vs ¥'+Math.round(ratv)+'）');
    if(weak.includes('连带率'))gaps.push('连带率差'+(rjd-jd).toFixed(2)+'（'+jd.toFixed(2)+' vs '+rjd.toFixed(2)+'）');
    if(weak.includes('笔数'))gaps.push('周笔数差'+(rcnt-cnt).toFixed(0)+'（'+cnt+' vs '+rcnt.toFixed(0)+'）');
    if(weak.includes('WOS占比'))gaps.push('WOS占比差'+(rwpct-wpct).toFixed(1)+'%（'+wpct.toFixed(1)+'% vs '+rwpct.toFixed(1)+'%）');
    if(weak.includes('长尾款占比'))gaps.push('长尾款占比差'+(rlpct-lpct).toFixed(1)+'%（'+lpct.toFixed(1)+'% vs '+rlpct.toFixed(1)+'%）');
    o='低于区域均值的维度：'+gaps.join('；')+'。';
  }
  if(!o)o='各维度均优于区域水平，多渠道运营体系成熟，保持综合策略。';
  return conc(c,o);
}
function concLongTail(s){
  const d=DATA[s];const lpct=d.wk_amt>0?d.wk_lt/d.wk_amt*100:0;
  const rl=regVal(x=>x.wk_lt)/(regVal(x=>x.wk_amt)||1)*100;
  const rlt=regVal(x=>x.wk_lt)/7;
  const c=s+'本周长尾款'+fmtAmt(d.wk_lt)+'，占比'+lpct.toFixed(1)+'%；区域长尾款占比'+rl.toFixed(1)+'%。';
  let o='';if(lpct<rl-2)o='长尾款占比低于区域均值'+(rl-lpct).toFixed(1)+'个百分点，建议加强79款SKU推荐话术培训，在票据台显眼位置陈列长尾小件。';else if(lpct>rl+2)o='长尾款占比高于区域均值'+(lpct-rl).toFixed(1)+'个百分点，继续保持关联推荐策略。';else o='长尾款与区域持平，关注件单价提升和连带搭配。';
  if(d.wk_lt<rlt*0.7)o+='长尾款金额低于区域门店均值（¥'+Math.round(rlt)+'），需重点跟进长尾款推荐。';
  return conc(c,o);
}
function concWow(s){
  const d=DATA[s],r=DATA['_region'];
  const c=s+'WK29环比'+(isNA(d.wow)?'N/A':(d.wow>=0?'+':'')+d.wow.toFixed(1)+'%')+'，'+(isNA(r.wow)?'区域N/A':'区域'+(r.wow>=0?'+':'')+r.wow.toFixed(1)+'%')+'。';
  let o='';if(!isNA(d.wow)){if(d.wow<-8)o='环比下滑明显，需复盘本周客流和转化率下降原因，排查天气/活动影响。';else if(d.wow<-3)o='环比小幅下滑，关注下周恢复节奏和竞品动态。';else if(d.wow<3)o='环比基本持平，需寻找增长突破点。';else o='环比增长良好，总结本周成功经验复制到下周。';}else o='新店需建立基准周数据后再分析环比趋势。';
  return conc(c,o);
}
function concTop(s){
  const d=DATA[s];const pd=(DATA['_products']||{})[s]||{};
  const ents=Object.entries(pd).map(([k,v])=>[k,v.amt||v,v.qty||0]).sort((a,b)=>b[1]-a[1]);
  if(ents.length===0)return conc('暂无产品数据。');
  const [n,a,q]=ents[0];const pct=d.wk_amt>0?a/d.wk_amt*100:0;
  const c='TOP1产品「'+n+'」销售'+fmtAmt(a)+'（'+q+'件），占周销售'+pct.toFixed(1)+'%。';
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
function concGuideSales(s){
  const gs=(DATA['_guides']||{})[s]||[];if(gs.length===0)return conc('暂无导购数据。');
  const top=gs.reduce((a,b)=>b.wk_sales>a.wk_sales?b:a);
  const tot=gs.reduce((a,g)=>a+g.wk_sales,0);
  const c='TOP1导购'+top.name+'周销售'+fmtAmt(top.wk_sales)+'（完成率'+top.wk_rate.toFixed(1)+'%），门店导购合计'+fmtAmt(tot)+'。';
  let o='';const stAvg=gs.length>0?tot/gs.length:0;
  const allRegGuides=[];STORE_ORDER.forEach(t=>{(DATA['_guides']||{})[t]?.forEach(g=>allRegGuides.push(g));});
  const regAvg=allRegGuides.length>0?allRegGuides.reduce((a,g)=>a+g.wk_sales,0)/allRegGuides.length:0;
  const topPct=tot>0?top.wk_sales/tot*100:0;if(topPct>50)o='TOP导购销售占门店'+topPct.toFixed(0)+'%，建议头部经验分享复制，缩小导购间差距。';
  const belowSt=gs.filter(g=>g.wk_sales>0&&g.wk_sales<stAvg*0.7);
  if(belowSt.length>0)o+=belowSt.map(g=>g.name+'('+fmtAmt(g.wk_sales)+')').join('、')+'低于门店均值'+fmtAmt(stAvg)+'，需重点跟进销售技巧辅导。';
  const belowReg=gs.filter(g=>g.wk_sales>0&&g.wk_sales<regAvg*0.7);
  if(belowReg.length>0)o+=belowReg.map(g=>g.name).join('、')+'低于区域导购均值'+fmtAmt(regAvg)+'，建议安排导师带教和排班优化。';
  const low=gs.filter(g=>g.wk_rate<30&&g.wk_sales>0);if(low.length>0)o+=low.length+'位导购完成率不足30%，需一对一辅导。';
  if(!o)o='导购团队整体表现均衡（门店人均'+fmtAmt(stAvg)+'，区域人均'+fmtAmt(regAvg)+'），继续保持团队协作氛围。';
  return conc(c,o);
}
function concGuideJdKdj(s){
  const atv=getWeeklyATV(s),jd=getWeeklyJD(s),ratv=getRegWkATV(),rjd=getRegWkJD();
  const c=s+'周连带率'+jd.toFixed(2)+'，客单价¥'+Math.round(atv)+'；区域连带率'+rjd.toFixed(2)+'，客单价¥'+Math.round(ratv)+'。';
  let o='';if(jd<rjd-0.1)o='连带率低于区域均值'+(rjd-jd).toFixed(2)+'（'+jd.toFixed(2)+' vs '+rjd.toFixed(2)+'），建议强化跨品类搭配推荐话术培训。';else if(jd<rjd)o='连带率略低于区域（差'+(rjd-jd).toFixed(2)+'），关注搭配推荐微调。';else o='连带率表现良好（高于区域'+(jd-rjd).toFixed(2)+'），关注高端产品推荐提升客单价。';
  if(atv<ratv*0.85)o+='客单价低于区域¥'+Math.round(ratv-atv)+'（¥'+Math.round(atv)+' vs ¥'+Math.round(ratv)+'），建议引导客户关注高价套装和限定款。';
  else if(atv<ratv)o+='客单价略低于区域¥'+Math.round(ratv-atv)+'，适度推荐高单价产品。';
  return conc(c,o);
}
function concGuideLt(s){
  const gs=(DATA['_guides']||{})[s]||[];if(gs.length===0)return conc('暂无导购数据。');
  const top=gs.reduce((a,b)=>b.wk_lt>a.wk_lt?b:a);
  const tot=gs.reduce((a,g)=>a+g.wk_lt,0),tq=gs.reduce((a,g)=>a+g.wk_lt_qty,0);
  const c='长尾款TOP导购'+top.name+'：'+fmtAmt(top.wk_lt)+'（'+top.wk_lt_qty+'件）；门店导购长尾款合计'+fmtAmt(tot)+'（'+tq+'件）。';
  let o='';const ltAvg=gs.length>0?tot/gs.length:0;
  const lowLt=gs.filter(g=>g.wk_lt===0||(g.wk_sales>0&&g.wk_lt/g.wk_sales<0.03));
  if(lowLt.length>0)o+=lowLt.map(g=>g.name+'(¥'+(g.wk_lt||0).toFixed(0)+')').join('、')+'长尾款占比偏低（门店均值¥'+Math.round(ltAvg)+'），建议专项培训79款SKU知识和搭配推荐。';
  const zeroLt=gs.filter(g=>g.wk_lt===0&&g.wk_sales>0);
  if(zeroLt.length>0)o+=zeroLt.map(g=>g.name).join('、')+'本周零长尾款销售，需重点跟进79款SKU推荐。';
  if(!o)o='导购长尾推荐整体较好（门店人均¥'+Math.round(ltAvg)+'），保持79款SKU的日常陈列和推荐习惯。';
  return conc(c,o);
}
"""

# 替换WK标签为动态编号
_js_out = JS_PART1.replace(f'WK{WK_NUM_INT}', f'WK{WK_NUM_INT}')  # no-op for current
_js_out = _js_out.replace('WK29', f'WK{WK_NUM_INT}').replace('WK28', f'WK{WK_PREV_NUM}')

with open(OUTPUT, 'a', encoding='utf-8') as f:
    f.write(_js_out)

import os
print(f"Part 2 (constants+helpers+conclusions) appended. Size: {os.path.getsize(OUTPUT)} bytes")
