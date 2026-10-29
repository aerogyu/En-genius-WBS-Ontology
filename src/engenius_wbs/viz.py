"""WBS 보기 HTML 한 장 — 외부 JS·CDN 없음 (폐쇄망에서 파일을 그대로 연다).

화면: ① 분야 간 연결 (Lv4 × Lv4 건수, 구성품·개발 단계로 거르기, 칸을 누르면 업무 목록)
      ② 업무 중심 선후 (검색한 업무의 앞뒤 N 단계만 그림, 누르면 그 업무로 옮김)
      ③ 검사 (문제 코드별 건수와 목록)
"""
from __future__ import annotations

import html
import json
from pathlib import Path

from .analyze import Analysis
from .parse import Wbs
from .report import all_issues


def payload(w: Wbs, a: Analysis) -> dict:
    crit = set(a.critical_chain)
    tasks = []
    for t in sorted(w.tasks.values(), key=lambda x: x.excel_row):
        tasks.append([t.no, t.wbs, t.level, t.name, t.module, t.phase, t.discipline, t.duration_days,
                      t.start.isoformat() if t.start else "", t.finish.isoformat() if t.finish else "",
                      a.up_count.get(t.no, 0), a.down_count.get(t.no, 0), round(a.total_float.get(t.no, 0), 1),
                      1 if t.no in crit else 0, t.parent, "; ".join(t.deliverables), 1 if t.is_summary else 0])
    edges = [[e.pred, e.succ, ("" if e.type == "FS" else e.type) + e.lag_text] for e in a.edges]
    issues = [[i.code, i.level, i.title, i.no, i.wbs, i.detail] for i in all_issues(w, a)]
    return {"source": Path(w.source).name, "tasks": tasks, "edges": edges, "issues": issues,
            "project_days": round(a.project_days, 1), "cycles": len(a.cycles)}


def write_html(w: Wbs, a: Analysis, path: Path, title: str = "WBS 선후 관계") -> Path:
    data = json.dumps(payload(w, a), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = TEMPLATE.replace("__TITLE__", html.escape(title)).replace("__DATA__", data)
    path.write_text(page, encoding="utf-8", newline="\n")
    return path


TEMPLATE = r"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title>
<style>
:root{--bg:#fafafa;--fg:#1f2328;--mut:#6b7280;--line:#d0d7de;--card:#fff;--acc:#2563eb;--hot:#dc2626;--sel:#fde68a}
@media (prefers-color-scheme:dark){:root{--bg:#16181d;--fg:#e6e6e6;--mut:#9aa0a6;--line:#363b44;--card:#1f2228;--acc:#60a5fa;--hot:#f87171;--sel:#7c5e10}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 "Malgun Gothic","Apple SD Gothic Neo",sans-serif}
header{padding:12px 16px;border-bottom:1px solid var(--line);display:flex;gap:16px;align-items:baseline;flex-wrap:wrap}
header h1{font-size:17px;margin:0}header .mut{color:var(--mut);font-size:12px}
nav{display:flex;gap:4px;padding:8px 16px}nav button{border:1px solid var(--line);background:var(--card);color:var(--fg);padding:6px 12px;border-radius:6px;cursor:pointer}
nav button.on{background:var(--acc);color:#fff;border-color:var(--acc)}
section{display:none;padding:8px 16px 40px}section.on{display:block}
.bar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:6px 0 10px}
select,input{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:6px;padding:5px 8px;font:inherit}
table{border-collapse:collapse}td,th{border:1px solid var(--line);padding:3px 6px;text-align:right;font-variant-numeric:tabular-nums}
th{background:var(--card);position:sticky;top:0}td.l,th.l{text-align:left}
.mx td.c{cursor:pointer;min-width:42px}.mx td.diag{outline:1px dashed var(--mut);outline-offset:-3px}
.scroll{overflow:auto;max-width:100%}.list{max-height:420px;overflow:auto;margin-top:10px}
.list table td{text-align:left}.lnk{color:var(--acc);cursor:pointer;text-decoration:underline}
#graph{border:1px solid var(--line);background:var(--card);border-radius:8px;overflow:auto;max-height:72vh}
#info{margin-top:8px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card)}
#hits div{cursor:pointer;padding:2px 4px}#hits div:hover{background:var(--sel)}
.pill{display:inline-block;padding:0 6px;border-radius:9px;background:var(--line);font-size:12px;margin-right:4px}
.err{color:var(--hot);font-weight:bold}
</style></head><body>
<header><h1>__TITLE__</h1><span class="mut" id="meta"></span></header>
<nav><button data-t="mx" class="on">① 분야 간 연결</button><button data-t="gr">② 업무 중심 선후</button><button data-t="is">③ 검사</button></nav>
<section id="mx" class="on">
 <div class="bar">구성품 <select id="fMod"></select> 개발 단계 <select id="fPh"></select>
  <label><input type="checkbox" id="fSame"> 같은 분야끼리도 보기</label>
  <span class="mut">행 = 선행 분야, 열 = 후속 분야 · 칸 = 선행 관계 수 (후속 업무 기준으로 거름) · 칸을 누르면 목록</span></div>
 <div class="scroll"><table class="mx" id="mxT"></table></div>
 <div class="list" id="mxL"></div>
</section>
<section id="gr">
 <div class="bar">업무 찾기 <input id="q" size="36" placeholder="번호 · WBS · 이름 일부">
  앞쪽 <select id="dU"></select> 단계 · 뒤쪽 <select id="dD"></select> 단계
  <label><input type="checkbox" id="onlyLeaf" checked> 상위 묶음 숨김</label></div>
 <div id="hits"></div>
 <div id="graph"></div>
 <div id="info" class="mut">업무를 찾거나 ① 의 목록에서 번호를 누르세요.</div>
</section>
<section id="is">
 <div class="bar">문제 코드 <select id="fIs"></select> <span class="mut" id="isN"></span></div>
 <div class="list" style="max-height:70vh"><table id="isT"></table></div>
</section>
<script>
const D=__DATA__;
const T=new Map(),PRE=new Map(),SUC=new Map();
for(const r of D.tasks){T.set(r[0],{no:r[0],wbs:r[1],lv:r[2],name:r[3],mod:r[4],ph:r[5],dis:r[6],dur:r[7],st:r[8],fi:r[9],up:r[10],dn:r[11],fl:r[12],cr:r[13],par:r[14],del:r[15],sum:r[16]})}
for(const e of D.edges){(PRE.get(e[1])||PRE.set(e[1],[]).get(e[1])).push([e[0],e[2]]);(SUC.get(e[0])||SUC.set(e[0],[]).get(e[0])).push([e[1],e[2]])}
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const $=id=>document.getElementById(id);
$("meta").textContent=`${D.source} · 업무 ${D.tasks.length.toLocaleString()} · 선행 관계 ${D.edges.length.toLocaleString()} · 순환 ${D.cycles} · 가장 긴 사슬 ${D.project_days} 근무일(달력 없는 근사) · 문제 ${D.issues.length.toLocaleString()}`;
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>show(b.dataset.t));
function show(t){document.querySelectorAll("nav button").forEach(b=>b.classList.toggle("on",b.dataset.t===t));document.querySelectorAll("section").forEach(s=>s.classList.toggle("on",s.id===t))}
// 색: 분야마다
const DIS=[...new Set(D.tasks.map(r=>r[6]).filter(Boolean))].sort();
const PAL=["#4e79a7","#f28e2b","#59a14f","#e15759","#76b7b2","#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac","#1f77b4","#2ca02c","#d62728","#9467bd","#8c564b","#17becf"];
const colorOf=d=>d?PAL[DIS.indexOf(d)%PAL.length]:"#999";
// ① 행렬
function opts(sel,vals){sel.innerHTML='<option value="">(전체)</option>'+vals.map(v=>`<option>${esc(v)}</option>`).join("")}
opts($("fMod"),[...new Set(D.tasks.map(r=>r[4]).filter(Boolean))].sort());
opts($("fPh"),[...new Set(D.tasks.map(r=>r[5]).filter(Boolean))]);
["fMod","fPh","fSame"].forEach(id=>$(id).onchange=drawMx);
let mxEdges=[];
function drawMx(){
 const m=$("fMod").value,p=$("fPh").value,same=$("fSame").checked,cnt=new Map();mxEdges=[];
 for(const e of D.edges){const a=T.get(e[0]),b=T.get(e[1]);if(!a.dis||!b.dis)continue;if(m&&b.mod!==m)continue;if(p&&b.ph!==p)continue;if(!same&&a.dis===b.dis)continue;
  mxEdges.push(e);const k=a.dis+"\u0000"+b.dis;cnt.set(k,(cnt.get(k)||0)+1)}
 const used=DIS.filter(d=>[...cnt.keys()].some(k=>k.split("\u0000").includes(d)));
 const mx=Math.max(1,...cnt.values());
 let h='<tr><th class="l">선행 ↓ · 후속 →</th>'+used.map(d=>`<th>${esc(d)}</th>`).join("")+"<th>합</th></tr>";
 for(const a of used){let s=0;h+=`<tr><th class="l">${esc(a)}</th>`;for(const b of used){const n=cnt.get(a+"\u0000"+b)||0;s+=n;
  const al=n?(0.12+0.78*n/mx):0;h+=`<td class="c${a===b?" diag":""}" data-a="${esc(a)}" data-b="${esc(b)}" style="background:${n?`rgba(37,99,235,${al.toFixed(2)})`:"transparent"};color:${al>0.55?"#fff":"inherit"}">${n||""}</td>`}
  h+=`<td><b>${s||""}</b></td></tr>`}
 $("mxT").innerHTML=used.length?h:'<tr><td class="l">조건에 맞는 분야 간 선행 관계가 없습니다.</td></tr>';
 $("mxT").querySelectorAll("td.c").forEach(td=>td.onclick=()=>listMx(td.dataset.a,td.dataset.b));$("mxL").innerHTML="";
}
function listMx(a,b){
 const rows=mxEdges.filter(e=>T.get(e[0]).dis===a&&T.get(e[1]).dis===b);
 let h=`<p><b>${esc(a)} → ${esc(b)}</b> ${rows.length}건</p><table><tr><th class="l">선행</th><th class="l">관계</th><th class="l">후속</th><th class="l">구성품 · 단계</th></tr>`;
 for(const e of rows.slice(0,2000)){const p=T.get(e[0]),q=T.get(e[1]);
  h+=`<tr><td><span class="lnk" data-n="${p.no}">${p.no}</span> ${esc(p.wbs)} ${esc(p.name)}</td><td>${esc(e[2]||"FS")}</td><td><span class="lnk" data-n="${q.no}">${q.no}</span> ${esc(q.wbs)} ${esc(q.name)}</td><td>${esc(q.mod)} · ${esc(q.ph)}</td></tr>`}
 if(rows.length>2000)h+=`<tr><td colspan="4">… ${rows.length-2000}건 더 (보고서 xlsx '분야 간 연결' 시트)</td></tr>`;
 $("mxL").innerHTML=h+"</table>";bindLinks($("mxL"));
}
function bindLinks(el){el.querySelectorAll(".lnk").forEach(x=>x.onclick=()=>{show("gr");center(+x.dataset.n)})}
drawMx();
// ② 업무 중심
for(const id of ["dU","dD"]){$(id).innerHTML=[1,2,3,4,5,6].map(n=>`<option${n===2?" selected":""}>${n}</option>`).join("");$(id).onchange=()=>cur&&center(cur)}
$("onlyLeaf").onchange=()=>cur&&center(cur);
$("q").oninput=()=>{const q=$("q").value.trim().toLowerCase();if(!q){$("hits").innerHTML="";return}
 const out=[];for(const t of T.values()){if(String(t.no)===q||t.wbs.toLowerCase().startsWith(q)||t.name.toLowerCase().includes(q)){out.push(t);if(out.length>=30)break}}
 $("hits").innerHTML=out.map(t=>`<div data-n="${t.no}"><span class="pill">Lv${t.lv}</span>${t.no} · ${esc(t.wbs)} · ${esc(t.name)} <span class="mut">${esc(t.mod)} / ${esc(t.ph)} / ${esc(t.dis)}</span></div>`).join("")||'<div class="mut">없음</div>';
 $("hits").querySelectorAll("div[data-n]").forEach(d=>d.onclick=()=>{$("hits").innerHTML="";center(+d.dataset.n)})};
let cur=null;
function walk(start,adj,depth,skipSum){ // 최단 거리 BFS. 상위 묶음은 숨기되 지나서 간다
 const dist=new Map([[start,0]]);let fr=[start];
 for(let d=1;d<=depth&&fr.length;d++){const nx=[];for(const v of fr)for(const [w] of (adj.get(v)||[])){if(dist.has(w))continue;dist.set(w,d);nx.push(w)}fr=nx}
 return dist}
function center(no){
 cur=no;const t=T.get(no);if(!t)return;
 const up=walk(no,PRE,+$("dU").value),dn=walk(no,SUC,+$("dD").value),hide=$("onlyLeaf").checked;
 const col=new Map();for(const [v,d] of up)col.set(v,-d);for(const [v,d] of dn)if(!col.has(v)||v===no)col.set(v,d);
 const cols=new Map();for(const [v,c] of col){if(hide&&T.get(v).sum&&v!==no)continue;(cols.get(c)||cols.set(c,[]).get(c)).push(v)}
 const keys=[...cols.keys()].sort((a,b)=>a-b),CAP=40,W=250,H=38,GX=70,GY=8,pos=new Map();let more=[];
 keys.forEach((c,i)=>{const vs=cols.get(c).sort((a,b)=>(T.get(a).dis||"").localeCompare(T.get(b).dis||"")||a-b);
  vs.slice(0,CAP).forEach((v,j)=>pos.set(v,[20+i*(W+GX),40+j*(H+GY)]));if(vs.length>CAP)more.push([i,vs.length-CAP])});
 const maxRows=Math.max(1,...keys.map(c=>Math.min(CAP,cols.get(c).length)));
 const width=40+keys.length*(W+GX),height=80+maxRows*(H+GY);
 let s=`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" font-size="12"><defs><marker id="ar" viewBox="0 0 8 8" refX="8" refY="4" markerWidth="7" markerHeight="7" orient="auto"><path d="M0,0L8,4L0,8z" fill="#888"/></marker></defs>`;
 keys.forEach((c,i)=>{s+=`<text x="${20+i*(W+GX)}" y="22" fill="currentColor" opacity=".6">${c<0?`앞 ${-c}단계`:c>0?`뒤 ${c}단계`:"선택한 업무"}</text>`});
 for(const [i,n] of more)s+=`<text x="${20+i*(W+GX)}" y="${height-12}" fill="currentColor" opacity=".6">+${n}개 더</text>`;
 for(const e of D.edges){const a=pos.get(e[0]),b=pos.get(e[1]);if(!a||!b)continue;
  const x1=a[0]+W,y1=a[1]+H/2,x2=b[0],y2=b[1]+H/2,mx=(x1+x2)/2,back=x2<=x1;
  const d=back?`M${a[0]+W/2},${a[1]+H} C${a[0]+W/2},${a[1]+H+40} ${b[0]+W/2},${b[1]+H+40} ${b[0]+W/2},${b[1]+H}`:`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`;
  s+=`<path d="${d}" fill="none" stroke="${back?"#dc2626":"#888"}" stroke-width="1.2" marker-end="url(#ar)"/>`;
  if(e[2])s+=`<text x="${(x1+x2)/2}" y="${(y1+y2)/2-3}" fill="currentColor" font-size="10" text-anchor="middle">${esc(e[2])}</text>`}
 for(const [v,[x,y]] of pos){const u=T.get(v),c=colorOf(u.dis),me=v===no;
  s+=`<g data-n="${v}" style="cursor:pointer"><rect x="${x}" y="${y}" width="${W}" height="${H}" rx="6" fill="${me?"var(--sel)":"var(--card)"}" stroke="${c}" stroke-width="${me?3:1.5}"/>`;
  s+=`<rect x="${x}" y="${y}" width="6" height="${H}" rx="3" fill="${c}"/>`;
  s+=`<text x="${x+12}" y="${y+15}" fill="currentColor">${esc((u.no+" "+u.name).slice(0,30))}</text>`;
  s+=`<text x="${x+12}" y="${y+30}" fill="currentColor" opacity=".65" font-size="11">${esc((u.dis||"")+" · "+(u.dur??"")+"일"+(u.cr?" · ★사슬":""))}</text><title>${esc(u.wbs+" "+u.name+"\n"+u.mod+" / "+u.ph+" / "+u.dis+"\n"+u.st+" ~ "+u.fi)}</title></g>`}
 s+="</svg>";$("graph").innerHTML=s;
 $("graph").querySelectorAll("g[data-n]").forEach(g=>g.onclick=()=>center(+g.dataset.n));
 const path=[];let p=t.par;while(p!=null){const x=T.get(p);path.unshift(`${x.wbs} ${x.name}`);p=x.par}
 const pre=(PRE.get(no)||[]).map(([v,k])=>`<span class="lnk" data-n="${v}">${v}</span>${esc(k)}`).join(", ")||"없음";
 const suc=(SUC.get(no)||[]).map(([v,k])=>`<span class="lnk" data-n="${v}">${v}</span>${esc(k)}`).join(", ")||"없음";
 const iss=D.issues.filter(i=>i[3]===no).map(i=>`<span class="${i[1]==="오류"?"err":""}">${i[0]} ${esc(i[2])}${i[5]?" — "+esc(i[5]):""}</span>`).join("<br>");
 $("info").className="";$("info").innerHTML=`<b>${t.no} · ${esc(t.wbs)} · ${esc(t.name)}</b> <span class="pill">Lv${t.lv}</span>${t.cr?'<span class="pill">가장 긴 사슬</span>':""}<br>
 <span class="mut">${esc(path.join("  ›  "))}</span><br>기간 ${t.dur??"-"}일 · ${t.st||"-"} ~ ${t.fi||"-"} · 여유 ${t.fl} 근무일 · 앞쪽 전체 ${t.up} · 뒤쪽 전체 ${t.dn}<br>
 선행: ${pre}<br>후속: ${suc}${t.del?`<br>산출물: ${esc(t.del)}`:""}${iss?`<br>${iss}`:""}`;
 bindLinks($("info"));
}
// ③ 검사
const codes=[...new Set(D.issues.map(i=>i[0]))];
$("fIs").innerHTML='<option value="">(전체)</option>'+codes.map(c=>{const i=D.issues.find(x=>x[0]===c);return `<option value="${c}">${c} ${esc(i[2])} (${D.issues.filter(x=>x[0]===c).length})</option>`}).join("");
$("fIs").onchange=drawIs;
function drawIs(){const c=$("fIs").value,rows=D.issues.filter(i=>!c||i[0]===c);$("isN").textContent=`${rows.length}건${rows.length>3000?" (앞 3,000건만 — 전체는 보고서 xlsx)":""}`;
 let h='<tr><th class="l">수준</th><th class="l">코드</th><th class="l">이름</th><th>번호</th><th class="l">WBS</th><th class="l">내용</th></tr>';
 for(const i of rows.slice(0,3000))h+=`<tr><td class="l ${i[1]==="오류"?"err":""}">${i[1]}</td><td class="l">${i[0]}</td><td class="l">${esc(i[2])}</td><td>${i[3]!=null&&T.has(i[3])?`<span class="lnk" data-n="${i[3]}">${i[3]}</span>`:(i[3]??"")}</td><td class="l">${esc(i[4])}</td><td class="l">${esc(i[5])}</td></tr>`;
 $("isT").innerHTML=h;bindLinks($("isT"))}
drawIs();
</script></body></html>
"""
