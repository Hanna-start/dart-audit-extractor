# -*- coding: utf-8 -*-
"""검증 통과 시계열 Excel → 자체 완결형 인터랙티브 HTML 대시보드.

- 비율(영업이익률·부채비율·성장률 등)은 전부 이 스크립트에서 결정적으로 산출(결정성 분리).
- ECharts JS를 인라인 임베드 → 인터넷 없이 더블클릭으로 열림.
- 회사·별도/연결 토글, KPI 카드, 추이/수익성/재무구조/현금흐름 차트, 데이터 표.

사용:
  py build_dashboard.py \
     --src "무신사=output/2026-05/무신사_시계열_2026-05-31.xlsx" \
     --src "와이즐리컴퍼니=output/2026-05/와이즐리컴퍼니_시계열_2026-05-17.xlsx" \
     --out output/2026-05/재무대시보드_2026-05-31.html
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd
import requests

ECHARTS_URL = "https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"
EOK = 1e8  # 억원


def _pick(long_df, gubun, sj, acct):
    s = long_df[(long_df["구분"] == gubun) & (long_df["재무제표"] == sj) & (long_df["계정"] == acct)]
    return dict(zip(s["사업연도"].astype(int), s["값"]))


def _ratio(num, den):
    if num is None or den is None or float(den) == 0:
        return None
    return round(float(num) / float(den) * 100, 1)


def _ok(x):
    return None if x is None else round(float(x) / EOK, 1)


def build_company(long_df, company):
    out = {}
    for gubun in ("별도", "연결"):
        sub = long_df[long_df["구분"] == gubun]
        if sub.empty:
            continue
        years = [int(y) for y in sorted(sub["사업연도"].astype(int).unique())]
        BS, IS, CF = "재무상태표", "포괄손익계산서", "현금흐름표"
        acc = {}
        for sj, names in {
            BS: ["자산총계", "부채총계", "자본총계", "유동자산", "비유동자산", "유동부채", "비유동부채", "자본금"],
            IS: ["매출액", "매출원가", "매출총이익", "영업이익", "판매비와관리비", "법인세비용", "당기순이익", "총포괄이익"],
            CF: ["영업활동현금흐름", "투자활동현금흐름", "재무활동현금흐름", "기말현금"],
        }.items():
            for nm in names:
                acc[nm] = _pick(long_df, gubun, sj, nm)

        def series(nm):
            return [acc.get(nm, {}).get(y) for y in years]

        rev = series("매출액"); op = series("영업이익"); ni = series("당기순이익")
        gp = series("매출총이익"); ast = series("자산총계"); lia = series("부채총계"); eq = series("자본총계")

        opm = [_ratio(op[i], rev[i]) for i in range(len(years))]
        nim = [_ratio(ni[i], rev[i]) for i in range(len(years))]
        gpm = [_ratio(gp[i], rev[i]) for i in range(len(years))]
        der = [_ratio(lia[i], eq[i]) for i in range(len(years))]
        grow = [None] + [(_ratio(rev[i] - rev[i-1], rev[i-1]) if rev[i] is not None and rev[i-1] not in (None, 0) else None)
                         for i in range(1, len(years))]

        out[gubun] = {
            "years": years,
            "매출액": [_ok(v) for v in rev], "영업이익": [_ok(v) for v in op],
            "당기순이익": [_ok(v) for v in ni], "매출총이익": [_ok(v) for v in gp],
            "자산총계": [_ok(v) for v in ast], "부채총계": [_ok(v) for v in lia],
            "자본총계": [_ok(v) for v in eq],
            "유동자산": [_ok(v) for v in series("유동자산")], "비유동자산": [_ok(v) for v in series("비유동자산")],
            "영업활동현금흐름": [_ok(v) for v in series("영업활동현금흐름")],
            "투자활동현금흐름": [_ok(v) for v in series("투자활동현금흐름")],
            "재무활동현금흐름": [_ok(v) for v in series("재무활동현금흐름")],
            "기말현금": [_ok(v) for v in series("기말현금")],
            "영업이익률": opm, "순이익률": nim, "매출총이익률": gpm, "부채비율": der, "매출성장률": grow,
        }
    return out


HTML = """<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DART 재무 대시보드</title>
<script>__ECHARTS__</script>
<style>
:root{--bg:#0f1419;--card:#1a2230;--ink:#e6edf3;--sub:#8b98a9;--line:#2a3645;
--up:#3fb950;--down:#f85149;--blue:#58a6ff;--amber:#e3b341;--purple:#bc8cff;}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font-family:'Segoe UI','Malgun Gothic',sans-serif;padding:24px}
h1{font-size:20px;margin:0 0 2px}.sub{color:var(--sub);font-size:13px;margin-bottom:18px}
.controls{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;align-items:center}
.seg{display:inline-flex;background:var(--card);border-radius:9px;padding:3px;border:1px solid var(--line)}
.seg button{background:none;border:none;color:var(--sub);padding:7px 16px;border-radius:7px;
cursor:pointer;font-size:13px;font-weight:600}
.seg button.on{background:var(--blue);color:#0b1117}
.seg button:disabled{opacity:.35;cursor:not-allowed}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:18px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .lbl{color:var(--sub);font-size:12px;margin-bottom:6px}
.kpi .val{font-size:23px;font-weight:700}.kpi .meta{font-size:12px;color:var(--sub);margin-top:4px}
.up{color:var(--up)}.down{color:var(--down)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
.panel h3{margin:0 0 8px;font-size:14px;font-weight:600}
.chart{height:300px}
table{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px}
th,td{padding:6px 8px;text-align:right;border-bottom:1px solid var(--line)}
th:first-child,td:first-child{text-align:left;color:var(--sub)}
thead th{color:var(--sub);font-weight:600}
@media(max-width:900px){.grid{grid-template-columns:1fr}}
.foot{color:var(--sub);font-size:11px;margin-top:18px;line-height:1.6}
</style></head><body>
<h1>DART 재무 시계열 대시보드</h1>
<div class="sub">단위: 억원 · 비율 %. 검증 통과 산출물 기반(회계등식 전수 ✓). 비율은 결정적 스크립트 산출.</div>
<div class="controls">
  <div class="seg" id="coSel"></div>
  <div class="seg" id="guSel"></div>
</div>
<div class="kpis" id="kpis"></div>
<div class="grid">
  <div class="panel"><h3>매출액 · 영업이익 추이</h3><div class="chart" id="c1"></div></div>
  <div class="panel"><h3>수익성 비율 (%)</h3><div class="chart" id="c2"></div></div>
  <div class="panel"><h3>재무구조: 자산 = 부채 + 자본 &amp; 부채비율</h3><div class="chart" id="c3"></div></div>
  <div class="panel"><h3>현금흐름</h3><div class="chart" id="c4"></div></div>
</div>
<div class="panel" style="margin-top:14px"><h3>데이터 표</h3><div id="tbl"></div></div>
<div class="foot" id="foot"></div>
<script>
const DATA = __DATA__;
const fmt = v => v==null?'–':v.toLocaleString('ko-KR');
const sign = v => v==null?'':(v>=0?'+':'')+v+'%';
let co = Object.keys(DATA)[0], gu = null;
const charts = {};

function mkSeg(el, items, cur, onPick, disabledSet){
  el.innerHTML='';
  items.forEach(it=>{const b=document.createElement('button');b.textContent=it;
    if(it===cur)b.className='on';
    if(disabledSet&&disabledSet.has(it))b.disabled=true;
    b.onclick=()=>onPick(it);el.appendChild(b);});
}
function render(){
  const co_data = DATA[co];
  const gus = Object.keys(co_data);
  if(!gus.includes(gu)) gu = gus[0];
  mkSeg(document.getElementById('coSel'),Object.keys(DATA),co,c=>{co=c;render();});
  mkSeg(document.getElementById('guSel'),['별도','연결'],gu,g=>{gu=g;render();},
        new Set(['별도','연결'].filter(g=>!gus.includes(g))));
  const d = co_data[gu], Y=d.years, last=Y.length-1;
  // KPI
  const kpiDef=[
    ['매출액','매출액','매출성장률'],['영업이익','영업이익','영업이익률'],
    ['당기순이익','당기순이익',null],['자본총계','자본총계','부채비율']];
  document.getElementById('kpis').innerHTML = kpiDef.map(([lbl,key,sub])=>{
    const v=d[key][last]; let meta='';
    if(sub==='매출성장률'){const g=d['매출성장률'][last];meta=`<span class="${g>=0?'up':'down'}">${sign(g)} YoY</span>`;}
    else if(sub==='영업이익률'){meta=`영업이익률 ${d['영업이익률'][last]??'–'}%`;}
    else if(sub==='부채비율'){meta=`부채비율 ${d['부채비율'][last]??'–'}%`;}
    const cls=(v!=null&&v<0)?'down':'';
    return `<div class="kpi"><div class="lbl">${lbl} (${Y[last]})</div>
      <div class="val ${cls}">${fmt(v)}</div><div class="meta">${meta}</div></div>`;
  }).join('');
  // charts
  const axis={type:'category',data:Y,axisLine:{lineStyle:{color:'#2a3645'}},axisLabel:{color:'#8b98a9'}};
  const yval={type:'value',axisLabel:{color:'#8b98a9'},splitLine:{lineStyle:{color:'#1e2733'}}};
  const tip={trigger:'axis',backgroundColor:'#1a2230',borderColor:'#2a3645',textStyle:{color:'#e6edf3'}};
  const leg=n=>({data:n,textStyle:{color:'#8b98a9'},top:0});
  draw('c1',{tooltip:tip,legend:leg(['매출액','영업이익']),grid:{left:60,right:20,top:30,bottom:30},
    xAxis:axis,yAxis:yval,series:[
      {name:'매출액',type:'bar',data:d['매출액'],itemStyle:{color:'#58a6ff'}},
      {name:'영업이익',type:'line',data:d['영업이익'],smooth:true,lineStyle:{width:3},itemStyle:{color:'#e3b341'}}]});
  draw('c2',{tooltip:tip,legend:leg(['영업이익률','순이익률','매출총이익률']),grid:{left:50,right:20,top:30,bottom:30},
    xAxis:axis,yAxis:{...yval,axisLabel:{color:'#8b98a9',formatter:'{value}%'}},series:[
      {name:'영업이익률',type:'line',data:d['영업이익률'],smooth:true,itemStyle:{color:'#3fb950'}},
      {name:'순이익률',type:'line',data:d['순이익률'],smooth:true,itemStyle:{color:'#f85149'}},
      {name:'매출총이익률',type:'line',data:d['매출총이익률'],smooth:true,itemStyle:{color:'#bc8cff'}}]});
  draw('c3',{tooltip:tip,legend:leg(['부채총계','자본총계','부채비율']),grid:{left:60,right:55,top:30,bottom:30},
    xAxis:axis,yAxis:[yval,{...yval,axisLabel:{color:'#8b98a9',formatter:'{value}%'},splitLine:{show:false}}],series:[
      {name:'부채총계',type:'bar',stack:'t',data:d['부채총계'],itemStyle:{color:'#f85149'}},
      {name:'자본총계',type:'bar',stack:'t',data:d['자본총계'],itemStyle:{color:'#3fb950'}},
      {name:'부채비율',type:'line',yAxisIndex:1,data:d['부채비율'],smooth:true,itemStyle:{color:'#e3b341'}}]});
  draw('c4',{tooltip:tip,legend:leg(['영업활동','투자활동','재무활동','기말현금']),grid:{left:60,right:20,top:30,bottom:30},
    xAxis:axis,yAxis:yval,series:[
      {name:'영업활동',type:'bar',data:d['영업활동현금흐름'],itemStyle:{color:'#58a6ff'}},
      {name:'투자활동',type:'bar',data:d['투자활동현금흐름'],itemStyle:{color:'#bc8cff'}},
      {name:'재무활동',type:'bar',data:d['재무활동현금흐름'],itemStyle:{color:'#e3b341'}},
      {name:'기말현금',type:'line',data:d['기말현금'],smooth:true,itemStyle:{color:'#3fb950'}}]});
  // table
  const rows=['매출액','매출총이익','영업이익','당기순이익','자산총계','부채총계','자본총계',
    '영업이익률','순이익률','부채비율','매출성장률'];
  let h='<table><thead><tr><th>계정</th>'+Y.map(y=>`<th>${y}</th>`).join('')+'</tr></thead><tbody>';
  rows.forEach(r=>{const pct=['영업이익률','순이익률','부채비율','매출성장률'].includes(r);
    h+=`<tr><td>${r}</td>`+d[r].map(v=>`<td>${v==null?'–':fmt(v)+(pct?'%':'')}</td>`).join('')+'</tr>';});
  h+='</tbody></table>';document.getElementById('tbl').innerHTML=h;
  document.getElementById('foot').innerHTML=DATA.__meta__;
}
function draw(id,opt){opt.textStyle={fontFamily:'Segoe UI'};
  charts[id]=charts[id]||echarts.init(document.getElementById(id),'dark');
  charts[id].setOption(opt,true);}
window.addEventListener('resize',()=>Object.values(charts).forEach(c=>c.resize()));
render();
</script></body></html>"""


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", action="append", required=True, help="회사명=경로")
    ap.add_argument("--out", required=True)
    ap.add_argument("--meta", default="")
    args = ap.parse_args(argv[1:])

    data = {}
    notes = []
    for spec in args.src:
        company, path = spec.split("=", 1)
        long_df = pd.read_excel(path, sheet_name="long_data")
        long_df = long_df[long_df["회사명"] == company]
        if long_df.empty:  # 파일에 회사명이 다르면 전체 사용
            long_df = pd.read_excel(path, sheet_name="long_data")
            company = long_df["회사명"].iloc[0]
        data[company] = build_company(long_df, company)
        yrs = sorted(long_df["사업연도"].astype(int).unique())
        notes.append(f"{company}: {yrs[0]}~{yrs[-1]}")

    data["__meta__"] = ("출처: DART 감사보고서/사업보고서 → 검증 통과 시계열. "
                        + " · ".join(notes)
                        + ". 무신사 2024~2025는 사업보고서 정형 API(fnlttSinglAcntAll), 그 외 감사보고서 PDF. "
                        "비율은 build_dashboard.py 결정적 산출.")

    js = requests.get(ECHARTS_URL, timeout=60).text
    html = (HTML.replace("__ECHARTS__", js)
                .replace("__DATA__", json.dumps(data, ensure_ascii=False)))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(html, encoding="utf-8")
    print(f"WROTE: {args.out}  ({len(html)//1024} KB)")


if __name__ == "__main__":
    main(sys.argv)
