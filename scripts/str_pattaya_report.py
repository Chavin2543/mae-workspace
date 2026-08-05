#!/usr/bin/env python3
"""Bilingual visual report for the Pattaya STR summary (same design language as
the Bangkok one): tiles + three Jan-Jun line charts (2025 dashed vs 2026 solid)
+ H1 table. Self-contained HTML, light+dark, hover tooltips.

Usage:
  python3 scripts/str_pattaya_report.py --out output/STR_Pattaya_by_market_2025-2026_report.html
"""
import argparse, json

TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pattaya STR — H1 2025 vs 2026</title>
<style>
  :root{
    color-scheme: light dark;
    --plane:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e;
    --muted:#898781; --grid:#e1e0d9; --axis:#c3c2b7; --ring:rgba(11,11,11,.10);
    --s1:#2a78d6; --good:#006300; --bad:#d03b3b;
  }
  @media (prefers-color-scheme: dark){ :root:where(:not([data-theme="light"])){
    --plane:#0d0d0d; --surface:#1a1a19; --ink:#fff; --ink2:#c3c2b7;
    --muted:#898781; --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
    --s1:#3987e5; --good:#0ca30c; --bad:#e66767; }}
  :root[data-theme="dark"]{
    --plane:#0d0d0d; --surface:#1a1a19; --ink:#fff; --ink2:#c3c2b7;
    --muted:#898781; --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
    --s1:#3987e5; --good:#0ca30c; --bad:#e66767; }
  *{box-sizing:border-box}
  body{margin:0;background:var(--plane);color:var(--ink);
    font-family:system-ui,-apple-system,"Segoe UI",sans-serif;line-height:1.5}
  .wrap{max-width:1040px;margin:0 auto;padding:28px 20px 64px}
  header h1{font-size:1.5rem;margin:0 0 4px}
  header p{margin:2px 0;color:var(--ink2);font-size:.9rem}
  .th{color:var(--muted)}
  .toggle{float:right;border:1px solid var(--ring);background:var(--surface);
    color:var(--ink2);border-radius:8px;padding:6px 12px;cursor:pointer;font-size:.8rem}
  .tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
    gap:12px;margin:22px 0 6px}
  .tile{background:var(--surface);border:1px solid var(--ring);border-radius:12px;
    padding:14px 16px}
  .tile .seg{font-weight:700;font-size:.85rem;margin-bottom:8px;display:flex;
    align-items:center;gap:7px}
  .dot{width:11px;height:11px;border-radius:3px;flex:none}
  .tile .kv{display:flex;justify-content:space-between;font-size:.82rem;
    color:var(--ink2);padding:2px 0}
  .tile .kv b{color:var(--ink);font-variant-numeric:tabular-nums}
  b.up{color:var(--good)} b.dn{color:var(--bad)}
  .tiles-note{color:var(--muted);font-size:.75rem;margin:0 0 18px}
  .card{background:var(--surface);border:1px solid var(--ring);border-radius:14px;
    padding:18px 18px 8px;margin:20px 0}
  .card h2{font-size:1.06rem;margin:0 0 2px}
  .card .sub{color:var(--muted);font-size:.8rem;margin:0 0 6px}
  .legend{display:flex;flex-wrap:wrap;gap:14px;margin:6px 0 4px}
  .legend span{display:flex;align-items:center;gap:6px;font-size:.8rem;color:var(--ink2)}
  .lsw{width:22px;height:0;border-top:2.5px solid;flex:none}
  .chart{width:100%;overflow-x:auto}
  svg{display:block;width:100%;height:auto;font-family:inherit}
  .gl{stroke:var(--grid);stroke-width:1}
  .ax{stroke:var(--axis);stroke-width:1}
  .tk{fill:var(--muted);font-size:11px}
  .lbl{font-size:11px;font-weight:600}
  .ln{fill:none;stroke-width:2}
  .mk{stroke:var(--surface);stroke-width:1.5}
  .cross{stroke:var(--axis);stroke-width:1;stroke-dasharray:3 3;opacity:0}
  .tip{position:fixed;pointer-events:none;background:var(--surface);
    border:1px solid var(--ring);border-radius:9px;padding:9px 11px;font-size:.78rem;
    box-shadow:0 6px 22px rgba(0,0,0,.16);opacity:0;transition:opacity .08s;z-index:9;
    color:var(--ink);min-width:170px}
  .tip .tt{font-weight:700;margin-bottom:5px}
  .tip .row{display:flex;justify-content:space-between;gap:14px;padding:1px 0}
  .tip .row span{display:flex;align-items:center;gap:6px;color:var(--ink2)}
  .tip .row b{font-variant-numeric:tabular-nums}
  footer{color:var(--muted);font-size:.75rem;margin-top:30px;
    border-top:1px solid var(--grid);padding-top:12px}
  table{border-collapse:collapse;width:100%;font-size:.8rem;margin-top:8px}
  th,td{border:1px solid var(--grid);padding:5px 8px;text-align:center;
    font-variant-numeric:tabular-nums}
  th{background:var(--surface);color:var(--ink2)}
  td.seg{text-align:left;font-weight:600}
  details{margin-top:6px}summary{cursor:pointer;color:var(--ink2);font-size:.82rem}
</style>
</head>
<body>
<div class="wrap">
  <button class="toggle" id="tg">◐ theme</button>
  <header>
    <h1>Pattaya STR — H1 2025 vs H1 2026</h1>
    <p>Occupancy, ADR and RevPAR · Jan–Jun monthly, Pattaya Area overall &amp; Upscale + Upper Mid class</p>
    <p class="th">เปรียบเทียบพัทยา ครึ่งปีแรก 2025 กับ 2026 (ภาพรวมตลาด และกลุ่ม Upscale &amp; Upper Mid)</p>
  </header>

  <div class="tiles" id="tiles"></div>
  <p class="tiles-note">YTD 2026 = Jan–Jun, calculated from monthly data · arrows compare with H1 2025 (calculated) · ตัวเลขคำนวณเองจากข้อมูลรายเดือน</p>

  <div class="card">
    <h2>Occupancy</h2>
    <p class="sub">อัตราการเข้าพัก (%) · Jan–Jun</p>
    <div class="legend" id="lg"></div>
    <div class="chart" id="c_occ"></div>
  </div>
  <div class="card">
    <h2>ADR — Average Daily Rate</h2>
    <p class="sub">ราคาห้องเฉลี่ยต่อคืน (THB) · Jan–Jun</p>
    <div class="chart" id="c_adr"></div>
  </div>
  <div class="card">
    <h2>RevPAR — Revenue per Available Room</h2>
    <p class="sub">รายได้ต่อห้องที่มีทั้งหมด (THB) · Jan–Jun</p>
    <div class="chart" id="c_rev"></div>
  </div>

  <div class="card">
    <h2>H1 (Jan–Jun) summary</h2>
    <p class="sub">ค่าเฉลี่ยครึ่งปีแรก และ % เปลี่ยนแปลง</p>
    <div id="ytables"></div>
  </div>

  <footer>
    Source: STR Monthly Performance Data (CoStar), Pattaya Area submarket &amp;
    Pattaya Area Upscale + Upper Mid class, Jan–Jun (This Year = 2026, Last Year
    = 2025). All H1/YTD averages are calculated from the monthly data
    (Occ/RevPAR day-weighted, ADR room-night-weighted); STR's own YTD rows are
    not used (Mae's decision, 2026-08-03). Generated 2026-08-03.
  </footer>
</div>
<div class="tip" id="tip"></div>

<script>
const D = __PAYLOAD__;
const SERIES = D.series;   // ["Pattaya overall","Upscale & Upper Mid"]
const COL = {"Pattaya overall":"--muted","Upscale & Upper Mid":"--s1"};
const YEARS=["2025","2026"];
const cssv = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const fmt = (n,d=0)=>n.toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d});

function tiles(){
  const el=document.getElementById("tiles"); el.innerHTML="";
  SERIES.forEach(s=>{
    const t=D.h1[s]["2026"], p=D.h1[s]["2025"];
    const g=[0,1,2].map(i=>(t[i]/p[i]-1)*100);
    const arrow=v=>`<b class="${v>=0?"up":"dn"}" style="font-size:.72rem">${v>=0?"▲":"▼"}${fmt(Math.abs(v),1)}%</b>`;
    const d=document.createElement("div"); d.className="tile";
    d.innerHTML=`<div class="seg"><span class="dot" style="background:var(${COL[s]})"></span>${s}</div>
      <div class="kv"><span>Occ YTD</span><span><b>${fmt(t[0],1)}%</b> ${arrow(g[0])}</span></div>
      <div class="kv"><span>ADR</span><span><b>${fmt(t[1])}</b> ${arrow(g[1])}</span></div>
      <div class="kv"><span>RevPAR</span><span><b>${fmt(t[2])}</b> ${arrow(g[2])}</span></div>`;
    el.appendChild(d);
  });
}

function legend(){
  const el=document.getElementById("lg"); el.innerHTML="";
  SERIES.forEach(s=>{
    YEARS.forEach(y=>{
      const sp=document.createElement("span");
      const dash=y==="2025"?"border-top-style:dashed;":"";
      sp.innerHTML=`<span class="lsw" style="border-top-color:var(${COL[s]});${dash}"></span>${s} ${y}`;
      el.appendChild(sp);
    });
  });
}

function chart(hostId, key, decimals, unit){
  const host=document.getElementById(hostId);
  const W=980,H=330,mL=58,mR=140,mT=14,mB=32;
  const iw=W-mL-mR, ih=H-mT-mB;
  const series=[];
  SERIES.forEach(s=>YEARS.forEach(y=>series.push(
    {name:s,year:y,vals:D.data[s][y][key],dashed:y==="2025"})));
  let lo=Infinity,hi=-Infinity;
  series.forEach(se=>se.vals.forEach(v=>{lo=Math.min(lo,v);hi=Math.max(hi,v);}));
  const pad=(hi-lo)*0.10||1; lo-=pad; hi+=pad;
  const n=D.months.length;
  const X=i=>mL+iw*i/(n-1);
  const Y=v=>mT+ih-(v-lo)/(hi-lo)*ih;
  const NS="http://www.w3.org/2000/svg";
  const svg=document.createElementNS(NS,"svg");
  svg.setAttribute("viewBox",`0 0 ${W} ${H}`);
  const add=(t,a)=>{const e=document.createElementNS(NS,t);for(const k in a)e.setAttribute(k,a[k]);svg.appendChild(e);return e;};
  for(let i=0;i<=4;i++){
    const v=lo+(hi-lo)*i/4, y=Y(v);
    add("line",{class:"gl",x1:mL,x2:mL+iw,y1:y,y2:y});
    const tk=add("text",{class:"tk",x:mL-8,y:y+4,"text-anchor":"end"});
    tk.textContent=unit==="%"?fmt(v,0)+"%":fmt(v,0);
  }
  add("line",{class:"ax",x1:mL,x2:mL+iw,y1:mT+ih,y2:mT+ih});
  D.months.forEach((m,i)=>{
    const tk=add("text",{class:"tk",x:X(i),y:mT+ih+17,"text-anchor":"middle"});
    tk.textContent=m;
  });
  const cross=add("line",{class:"cross",y1:mT,y2:mT+ih,x1:mL,x2:mL});
  const colOf=se=>cssv(COL[se.name]);
  series.forEach(se=>{
    let d="";
    se.vals.forEach((v,i)=>{ d+=(i?"L":"M")+X(i).toFixed(1)+" "+Y(v).toFixed(1)+" "; });
    const attrs={class:"ln",d,stroke:colOf(se)};
    if(se.dashed) attrs["stroke-dasharray"]="6 4";
    add("path",attrs);
  });
  const li=n-1;
  let ends=series.map(se=>({t:`${se.name==="Pattaya overall"?"Overall":"Up+UpMid"} ${se.year}`,
                            col:colOf(se),y:Y(se.vals[li])})).sort((a,b)=>a.y-b.y);
  const GAP=13;
  for(let i=1;i<ends.length;i++) if(ends[i].y-ends[i-1].y<GAP) ends[i].y=ends[i-1].y+GAP;
  const over=ends[ends.length-1].y-(mT+ih); if(over>0) ends.forEach(e=>e.y-=over);
  ends.forEach(e=>{
    const lab=add("text",{class:"lbl",x:X(li)+8,y:e.y+4,fill:e.col});
    lab.textContent=e.t;
  });
  const tip=document.getElementById("tip");
  const dots=[];
  series.forEach(se=>{
    se.vals.forEach((v,i)=>{
      const c=add("circle",{cx:X(i),cy:Y(v),r:0,fill:colOf(se),class:"mk"});
      dots.push({i,c});
    });
  });
  const hit=add("rect",{x:mL,y:mT,width:iw,height:ih,fill:"transparent"});
  hit.style.cursor="crosshair";
  function show(i,px,py){
    cross.setAttribute("x1",X(i));cross.setAttribute("x2",X(i));cross.style.opacity=1;
    dots.forEach(d=>d.c.setAttribute("r",d.i===i?4.5:0));
    let rows=`<div class="tt">${D.months[i]}</div>`;
    series.forEach(se=>{
      rows+=`<div class="row"><span><span class="dot" style="background:${colOf(se)}"></span>${se.name} ${se.year}</span><b>${fmt(se.vals[i],decimals)}${unit==="%"?"%":""}</b></div>`;
    });
    tip.innerHTML=rows; tip.style.opacity=1;
    let tx=px+16; if(tx+200>window.innerWidth) tx=px-210;
    tip.style.left=tx+"px"; tip.style.top=(py-10)+"px";
  }
  hit.addEventListener("mousemove",ev=>{
    const r=svg.getBoundingClientRect();
    const sx=(ev.clientX-r.left)/r.width*W;
    let i=Math.round((sx-mL)/(iw/(n-1)));
    i=Math.max(0,Math.min(n-1,i));
    show(i,ev.clientX,ev.clientY);
  });
  hit.addEventListener("mouseleave",()=>{cross.style.opacity=0;dots.forEach(d=>d.c.setAttribute("r",0));tip.style.opacity=0;});
  host.innerHTML=""; host.appendChild(svg);
}

function ytables(){
  const host=document.getElementById("ytables");
  const metrics=[["Occupancy (%)",0,1],["ADR (THB)",1,0],["RevPAR (THB)",2,0]];
  let html="";
  metrics.forEach(([name,mi,dec])=>{
    html+=`<details ${mi===0?"open":""}><summary>${name}</summary><table>
      <tr><th>Market</th><th>H1 2025</th><th>H1 2026</th><th>26 vs 25</th></tr>`;
    SERIES.forEach(s=>{
      const v25=D.h1[s]["2025"][mi], v26=D.h1[s]["2026"][mi];
      const pct=(v26/v25-1)*100, up=pct>=0;
      html+=`<tr><td class="seg"><span class="dot" style="display:inline-block;background:var(${COL[s]})"></span> ${s}</td>
        <td>${fmt(v25,dec)}</td><td>${fmt(v26,dec)}</td>
        <td><b class="${up?"up":"dn"}">${up?"▲":"▼"} ${fmt(Math.abs(pct),1)}%</b></td></tr>`;
    });
    html+="</table></details>";
  });
  host.innerHTML=html;
}

function draw(){ tiles(); legend(); ytables();
  chart("c_occ","occ",1,"%"); chart("c_adr","adr",0,"THB"); chart("c_rev","revpar",0,"THB"); }
draw();
document.getElementById("tg").addEventListener("click",()=>{
  const cur=document.documentElement.getAttribute("data-theme");
  const next=cur==="dark"?"light":(cur==="light"?"dark":
    (matchMedia("(prefers-color-scheme: dark)").matches?"light":"dark"));
  document.documentElement.setAttribute("data-theme",next);
  draw();
});
matchMedia("(prefers-color-scheme: dark)").addEventListener("change",draw);
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", default="")
    ap.add_argument("--out", default="output/STR_Pattaya_by_market_2025-2026_report.html")
    a = ap.parse_args()
    if a.json:
        payload = json.load(open(a.json))
    else:
        from str_pattaya_summary import load_all, make_payload
        payload = make_payload(load_all())
    html = TEMPLATE.replace("__PAYLOAD__", json.dumps(payload))
    with open(a.out, "w") as f:
        f.write(html)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
