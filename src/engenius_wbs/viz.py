"""WBS 보기 HTML 한 장 (v2) — 외부 JS·CDN 없음 (폐쇄망에서 파일을 그대로 연다).

조작 셋: 범위(왼쪽 트리에서 고른 WBS 요소, 또는 전체) × 알갱이(범위 아래 몇 레벨로 뭉쳐 그릴지) × 배치(시간표 | 흐름도).
  · 시간표: 가로 = 날짜(E·F열), 줄 = 범위 바로 아래 묶음, 막대 = 알갱이 요소, 화살표 = 선행 (줄을 넘는 것 = 넘겨받기, 주황)
  · 흐름도: 선행 순서로 층을 나눠 왼쪽→오른쪽 (뭉친 뒤 거꾸로 가는 간선은 빨간 점선)
  범위 밖에서 오는/나가는 선행은 점선 토막(범위와 같은 레벨의 요소로 뭉침). 이름은 자르지 않는다 (줄바꿈·막대 옆).
  한 번 누르면 앞뒤 사슬 강조 + 오른쪽 패널, 두 번 누르면 그 안으로 내려감, 빵부스러기로 올라옴. 검사 탭은 그대로.
규모 참고 (사내 2026-10-29, 숫자만): Lv1 10 · Lv2 19 · Lv3 147 · Lv4 629 · Lv5 2,249 · Lv6 9,771 → Lv3 범위 + Lv5 알갱이 ≈ 15 상자.
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
        tasks.append([t.no, t.wbs, t.level, t.name, t.duration_days, t.start.isoformat() if t.start else "",
                      t.finish.isoformat() if t.finish else "", a.up_count.get(t.no, 0), a.down_count.get(t.no, 0),
                      round(a.total_float.get(t.no, 0), 1), 1 if t.no in crit else 0, t.parent, "; ".join(t.deliverables),
                      t.man_hours, t.staff])
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
:root{--bg:#fafafa;--fg:#1f2328;--mut:#6b7280;--line:#d0d7de;--card:#fff;--acc:#2563eb;--hot:#dc2626;--hand:#d97706;--sel:#fde68a;--up:#7c3aed;--dn:#059669;--band:#b0b4ba}
@media (prefers-color-scheme:dark){:root{--bg:#16181d;--fg:#e6e6e6;--mut:#9aa0a6;--line:#363b44;--card:#1f2228;--acc:#60a5fa;--hot:#f87171;--hand:#fbbf24;--sel:#7c5e10;--up:#a78bfa;--dn:#34d399;--band:#c9ccd1}}
*{box-sizing:border-box}html,body{height:100%}body{margin:0;background:var(--bg);color:var(--fg);font:13px/1.45 "Malgun Gothic","Apple SD Gothic Neo",sans-serif;display:flex;flex-direction:column}
header{padding:8px 14px;border-bottom:1px solid var(--line);display:flex;gap:14px;align-items:baseline;flex-wrap:wrap}
header h1{font-size:16px;margin:0}.mut{color:var(--mut);font-size:12px}
nav{display:flex;gap:4px;padding:6px 14px;border-bottom:1px solid var(--line)}nav button{border:1px solid var(--line);background:var(--card);color:var(--fg);padding:4px 10px;border-radius:6px;cursor:pointer}
nav button.on{background:var(--acc);color:#fff;border-color:var(--acc)}
section{display:none;flex:1;min-height:0}section.on{display:flex}
#gr.on{flex-direction:row}
#tree{width:270px;min-width:200px;border-right:1px solid var(--line);display:flex;flex-direction:column;background:var(--card)}
#tree .bar{padding:6px 8px;border-bottom:1px solid var(--line)}#tree input{width:100%}
#treeBody{flex:1;overflow:auto;padding:4px 0}
.tn{display:flex;align-items:center;gap:4px;padding:2px 6px;cursor:pointer;white-space:nowrap}.tn:hover{background:var(--bg)}.tn.cur{background:var(--sel)}
.tn .tg{width:14px;color:var(--mut);text-align:center;flex:none}.tn .lv{color:var(--mut);font-size:10.5px;flex:none;width:26px}.tn .nm{overflow:hidden;text-overflow:ellipsis}.tn .ct{color:var(--mut);font-size:11px;margin-left:auto;padding-left:6px}
#hits div{padding:3px 8px;cursor:pointer;border-bottom:1px solid var(--line)}#hits div:hover{background:var(--bg)}#hits .p{color:var(--mut);font-size:11px;display:block}
#main{flex:1;min-width:0;display:flex;flex-direction:column}
#crumb{padding:6px 12px;border-bottom:1px solid var(--line);display:flex;gap:6px;flex-wrap:wrap;align-items:center}
#crumb span.c{cursor:pointer;color:var(--acc)}#crumb span.c:hover{text-decoration:underline}
#ctl{padding:6px 12px;display:flex;gap:12px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--line)}
select,input{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:6px;padding:4px 6px;font:inherit}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:6px;overflow:hidden}.seg button{border:0;background:var(--card);color:var(--fg);padding:4px 10px;cursor:pointer}.seg button.on{background:var(--acc);color:#fff}
#canvasWrap{flex:1;position:relative;overflow:hidden;background:var(--card)}
#canvas{width:100%;height:100%;cursor:grab;display:block}#canvas:active{cursor:grabbing}
#legend{position:absolute;left:10px;bottom:8px;font-size:11px;color:var(--mut);background:var(--card);padding:2px 6px;border-radius:6px;border:1px solid var(--line)}
#zoom{position:absolute;right:10px;top:8px;display:flex;gap:4px}#zoom button{border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:6px;width:28px;height:28px;cursor:pointer}
#side{width:330px;min-width:260px;border-left:1px solid var(--line);overflow:auto;padding:10px 12px;background:var(--card)}
#side h3{margin:8px 0 4px;font-size:13px}#side table{border-collapse:collapse;width:100%}#side td{padding:2px 4px;border-bottom:1px solid var(--line);vertical-align:top}#side td:first-child{color:var(--mut);white-space:nowrap}
.lnk{color:var(--acc);cursor:pointer;text-decoration:underline}.pill{display:inline-block;padding:0 6px;border-radius:9px;background:var(--line);font-size:11px;margin-right:4px}.err{color:var(--hot);font-weight:bold}
#is.on{flex-direction:column;padding:8px 14px;overflow:auto}#is table{border-collapse:collapse}#is td,#is th{border:1px solid var(--line);padding:3px 6px;text-align:left}#is th{background:var(--card);position:sticky;top:0}
svg text{user-select:none}.nd{cursor:pointer}.nd rect{stroke-width:1.5}.nd.dim{opacity:.18}.ed.dim{opacity:.08}.nd.sel rect.bx{stroke-width:3}
</style></head><body>
<header><h1>__TITLE__</h1><span class="mut" id="meta"></span></header>
<nav><button data-t="gr" class="on">① 선후 관계 보기</button><button data-t="is">② 검사</button></nav>
<section id="gr" class="on">
 <div id="tree"><div class="bar"><input id="q" placeholder="찾기: 번호 · WBS · 이름 일부"></div><div id="hits"></div><div id="treeBody"></div></div>
 <div id="main">
  <div id="crumb"></div>
  <div id="ctl">
   <span>알갱이 <select id="grain"></select></span>
   <span class="seg" id="layout"><button data-v="time" class="on">시간표</button><button data-v="flow">흐름도</button></span>
   <span title="범위 안에서 어느 레벨의 같은 이름만 남긴다 — 예: Lv4 '성능' 을 고르면 모든 구성품·단계의 성능 분야만">같은 이름만 <select id="fLv"><option value="">(끄기)</option></select> <select id="fName" style="max-width:180px"></select></span>
   <label><input type="checkbox" id="chk"> 가장 긴 사슬 강조</label>
   <label><input type="checkbox" id="ext" checked> 범위 밖 출입 표시</label>
   <span class="mut" id="stat"></span>
  </div>
  <div id="canvasWrap"><svg id="canvas" xmlns="http://www.w3.org/2000/svg"></svg>
   <div id="legend">상자 = 이름표(왼쪽 끝 = 시작일) · 상자 아래 회색 띠 = 실제 기간 · 회색 = 같은 줄 안 선행 · <span style="color:var(--hand)">주황</span> = 줄을 넘는 넘겨받기 · 점선 토막 = 범위 밖 · 누름 = 앞(보라)/뒤(초록) 강조 · 두 번 누름 = 안으로</div>
   <div id="zoom"><button id="zi" title="확대">+</button><button id="zo" title="축소">−</button><button id="zf" title="맞춤">⌂</button></div>
  </div>
 </div>
 <div id="side"><div class="mut">왼쪽 트리에서 범위를 고르세요.</div></div>
</section>
<section id="is">
 <div style="margin-bottom:6px">문제 코드 <select id="fIs"></select> <span class="mut" id="isN"></span></div>
 <table id="isT"></table>
</section>
<script>
const D=__DATA__;
const T=new Map(),KIDS=new Map(),PRE=new Map(),SUC=new Map(),ROOTS=[];
for(const r of D.tasks){T.set(r[0],{no:r[0],wbs:r[1],lv:r[2],name:r[3],dur:r[4],st:r[5],fi:r[6],up:r[7],dn:r[8],fl:r[9],cr:r[10],par:r[11],del:r[12],mh:r[13],staff:r[14],kids:[]})}
for(const t of T.values()){if(t.par!=null&&T.has(t.par))T.get(t.par).kids.push(t.no);else ROOTS.push(t.no)}
for(const e of D.edges){(PRE.get(e[1])||PRE.set(e[1],[]).get(e[1])).push(e);(SUC.get(e[0])||SUC.set(e[0],[]).get(e[0])).push(e)}
// 하위 잎 수·MH 합·사슬 포함 여부를 아래에서 위로
const LEAF=new Map(),MH=new Map(),CR=new Map();
(function roll(list){for(const no of list){const t=T.get(no);if(!t.kids.length){LEAF.set(no,1);MH.set(no,t.mh||0);CR.set(no,t.cr);continue}roll(t.kids);let l=0,m=0,c=0;for(const k of t.kids){l+=LEAF.get(k);m+=MH.get(k);c=c||CR.get(k)}LEAF.set(no,l);MH.set(no,m);CR.set(no,c)}})(ROOTS);
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const $=id=>document.getElementById(id);
const LVNAME={1:"구성품",2:"최상위 업무",3:"개발 단계",4:"기술 분야",5:"상위 기술 업무",6:"하위 기술 업무"};
const PAL=["#4e79a7","#f28e2b","#59a14f","#e15759","#76b7b2","#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac","#1f77b4","#2ca02c","#d62728","#9467bd","#8c564b","#17becf"];
$("meta").textContent=`${D.source} · 업무 ${D.tasks.length.toLocaleString()} · 선행 관계 ${D.edges.length.toLocaleString()} · 순환 ${D.cycles} · 가장 긴 사슬 ${D.project_days} 근무일(달력 없는 근사) · 문제 ${D.issues.length.toLocaleString()}`;
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>{document.querySelectorAll("nav button").forEach(x=>x.classList.toggle("on",x===b));document.querySelectorAll("section").forEach(s=>s.classList.toggle("on",s.id===b.dataset.t))});
function anc(no,lv){let t=T.get(no);while(t&&t.lv>lv&&t.par!=null)t=T.get(t.par);return t?t.no:no}
function path(no){const p=[];let t=T.get(no);while(t){p.unshift(t);t=t.par!=null?T.get(t.par):null}return p}
function pathText(no){return path(no).map(t=>t.name).join(" › ")}
// ---- 트리 ----
const open=new Set();
function renderTree(){const el=$("treeBody");let h="";const rec=(list,d)=>{for(const no of list){const t=T.get(no),has=t.kids.length>0,o=open.has(no);
 h+=`<div class="tn${S.scope===no?" cur":""}" data-n="${no}" style="padding-left:${6+d*14}px"><span class="tg" data-tg="${no}">${has?(o?"▾":"▸"):"·"}</span><span class="lv">Lv${t.lv}</span><span class="nm" title="${esc(t.wbs+" "+t.name)}">${esc(t.name)}</span><span class="ct">${LEAF.get(no).toLocaleString()}</span></div>`;
 if(has&&o)rec(t.kids,d+1)}};rec(ROOTS,0);el.innerHTML=h;
 el.querySelectorAll(".tg").forEach(x=>x.onclick=ev=>{ev.stopPropagation();const n=+x.dataset.tg;open.has(n)?open.delete(n):open.add(n);renderTree()});
 el.querySelectorAll(".tn").forEach(x=>x.onclick=()=>{const n=+x.dataset.n;if(T.get(n).kids.length){setScope(n)}else{setScope(T.get(n).par,6);select(n)}})}
$("q").oninput=()=>{const q=$("q").value.trim().toLowerCase();if(!q){$("hits").innerHTML="";return}
 const out=[];for(const t of T.values()){if(String(t.no)===q||t.wbs.toLowerCase().startsWith(q)||t.name.toLowerCase().includes(q)){out.push(t);if(out.length>=40)break}}
 $("hits").innerHTML=out.map(t=>`<div data-n="${t.no}"><span class="pill">Lv${t.lv}</span>${t.no} · ${esc(t.name)}<span class="p">${esc(path(t.no).slice(0,-1).map(x=>x.name).join(" › "))}</span></div>`).join("")||'<div class="mut">없음</div>';
 $("hits").querySelectorAll("div[data-n]").forEach(d=>d.onclick=()=>{const n=+d.dataset.n;$("hits").innerHTML="";$("q").value="";if(T.get(n).kids.length)setScope(n);else{setScope(T.get(n).par,6);select(n)}})};
// ---- 상태 ----
const S={scope:0,grain:2,layout:"time",sel:null,fLv:0,fName:""};
function scopeLevel(){return S.scope?T.get(S.scope).lv:0}
function setScope(no,grain){S.scope=no||0;const L=scopeLevel();S.grain=grain||Math.min(6,L+2);S.sel=null;
 for(const t of path(S.scope||ROOTS[0]).slice(0,-1))open.add(t.no);if(S.scope)open.add(S.scope);
 const g=$("grain");g.innerHTML="";for(let l=L+1;l<=6;l++){g.innerHTML+=`<option value="${l}"${l===S.grain?" selected":""}>Lv${l} ${LVNAME[l]}</option>`}
 if(S.fLv&&S.fLv<=L){S.fLv=0;S.fName=""}
 fillFilter();renderTree();draw();const cur=$("treeBody").querySelector(".tn.cur");if(cur)cur.scrollIntoView({block:"nearest"})}
// 같은 이름만 거르기 — 범위 안 그 레벨 요소의 이름 목록
function namesAt(lv){const L=scopeLevel(),out=new Set();for(const t of T.values())if(t.lv===lv&&(!S.scope||anc(t.no,L)===S.scope))out.add(t.name);return [...out].sort()}
function fillFilter(){const L=scopeLevel(),fl=$("fLv");fl.innerHTML='<option value="">(끄기)</option>';for(let l=L+1;l<=5;l++)fl.innerHTML+=`<option value="${l}"${l===S.fLv?" selected":""}>Lv${l} ${LVNAME[l]}</option>`;
 const fn=$("fName");fn.innerHTML="";fn.style.display=S.fLv?"":"none";if(!S.fLv)return;const names=namesAt(S.fLv);if(!names.includes(S.fName))S.fName=names[0]||"";
 fn.innerHTML=names.map(n=>`<option${n===S.fName?" selected":""}>${esc(n)}</option>`).join("")}
$("fLv").onchange=()=>{S.fLv=+$("fLv").value||0;S.fName="";S.sel=null;fillFilter();draw()};
$("fName").onchange=()=>{S.fName=$("fName").value;S.sel=null;draw()};
$("grain").onchange=()=>{S.grain=+$("grain").value;S.sel=null;draw()};
$("layout").querySelectorAll("button").forEach(b=>b.onclick=()=>{S.layout=b.dataset.v;$("layout").querySelectorAll("button").forEach(x=>x.classList.toggle("on",x===b));draw()});
$("chk").onchange=draw;$("ext").onchange=draw;
// ---- 뭉치기 ----
let G=null;
function build(){const L=scopeLevel(),g=S.grain,F=S.fLv&&S.fName?S.fLv:0;
 const pass=no=>!F||T.get(anc(no,F)).name===S.fName;             // 같은 이름만: 그 레벨 조상의 이름이 같은 잎만
 const inScope=no=>(!S.scope||anc(no,L)===S.scope)&&pass(no);
 const leaves=[];const walk=list=>{for(const no of list){const t=T.get(no);if(t.kids.length)walk(t.kids);else if(pass(no))leaves.push(no)}};walk(S.scope?[S.scope]:ROOTS);
 const nodes=new Map(),nodeOf=new Map();
 for(const lf of leaves){const n=anc(lf,g);nodeOf.set(lf,n);if(!nodes.has(n)){const t=T.get(n);nodes.set(n,{id:n,t,ext:0,lane:g>L+1?anc(n,L+1):0,leaves:0,mh:0,cr:CR.get(n)})}const nd=nodes.get(n);nd.leaves++;nd.mh+=T.get(lf).mh||0}
 const lanes=[];const laneIdx=new Map();if(g>L+1){const kids=S.scope?T.get(S.scope).kids:ROOTS;kids.forEach((k,i)=>{laneIdx.set(k,i);lanes.push(k)})}
 const edges=new Map(),ext=new Map(),handoff=new Map();let internal=0,inN=0,outN=0;
 const extKey=no=>anc(no,F&&(!S.scope||anc(no,L)===S.scope)?F:Math.max(L,1));   // 거른 같은 범위 안이면 그 레벨(예: 다른 분야)로 뭉친다
 const extNode=(k,dir)=>{const id=`${dir}:${k}`;if(!ext.has(id)){const pth=path(k),sp=S.scope?path(S.scope):[];let i=0;while(i<pth.length&&i<sp.length&&pth[i].no===sp[i].no)i++;ext.set(id,{id,ext:dir,key:k,t:T.get(k),label:pth.slice(Math.max(0,i)).map(x=>x.name).join(" › ")||T.get(k).name,count:0})}return ext.get(id)};
 for(const e of D.edges){const pi=inScope(e[0]),si=inScope(e[1]);if(!pi&&!si)continue;
  if(pi&&si){const a=nodeOf.get(e[0])??anc(e[0],g),b=nodeOf.get(e[1])??anc(e[1],g);if(a===b){internal++;continue}const k=a+">"+b;let x=edges.get(k);if(!x){x=edges.set(k,{a,b,n:0,ex:[],cross:nodes.get(a)?.lane!==nodes.get(b)?.lane}).get(k)}x.n++;if(x.ex.length<8)x.ex.push(e);
   if(x.cross&&lanes.length){const hk=nodes.get(a).lane+">"+nodes.get(b).lane;handoff.set(hk,(handoff.get(hk)||0)+1)}}
  else if(si){inN++;if(!$("ext").checked)continue;const b=nodeOf.get(e[1])??anc(e[1],g),xn=extNode(extKey(e[0]),"in");xn.count++;const k=xn.id+">"+b;let x=edges.get(k);if(!x)x=edges.set(k,{a:xn.id,b,n:0,ex:[],cross:false,ext:true}).get(k);x.n++;if(x.ex.length<8)x.ex.push(e)}
  else{outN++;if(!$("ext").checked)continue;const a=nodeOf.get(e[0])??anc(e[0],g),xn=extNode(extKey(e[1]),"out");xn.count++;const k=a+">"+xn.id;let x=edges.get(k);if(!x)x=edges.set(k,{a,b:xn.id,n:0,ex:[],cross:false,ext:true}).get(k);x.n++;if(x.ex.length<8)x.ex.push(e)}}
 // 범위 안 가장 긴 사슬 (잎 단위, 기간 합 — 달력 없는 근사)
 const lset=new Set(leaves),dist=new Map(),from=new Map(),indeg=new Map();for(const l of leaves){indeg.set(l,0)}
 for(const l of leaves)for(const e of (SUC.get(l)||[]))if(lset.has(e[1]))indeg.set(e[1],indeg.get(e[1])+1);
 const q=leaves.filter(l=>!indeg.get(l));for(const l of q)dist.set(l,T.get(l).dur||0);let qi=0;
 while(qi<q.length){const v=q[qi++];for(const e of (SUC.get(v)||[])){const w=e[1];if(!lset.has(w))continue;const d=dist.get(v)+(T.get(w).dur||0);if(!dist.has(w)||d>dist.get(w)){dist.set(w,d);from.set(w,v)}indeg.set(w,indeg.get(w)-1);if(!indeg.get(w))q.push(w)}}
 let end=null,best=-1;for(const [v,d] of dist)if(d>best){best=d;end=v}const chain=[];for(let v=end;v!=null;v=from.get(v))chain.unshift(v);
 G={L,g,nodes,edges,ext,lanes,laneIdx,handoff,internal,inN,outN,leaves:leaves.length,chain,chainDays:best<0?0:best}}
// ---- 그리기 ----
const svg=$("canvas");let VB={x:0,y:0,w:1000,h:600},drawn={w:1000,h:600};
const ctx=document.createElement("canvas").getContext("2d");ctx.font='12px "Malgun Gothic","Apple SD Gothic Neo",sans-serif';
const tw=s=>ctx.measureText(s).width;
function wrap(s,maxW){const words=s.split(/(\s+)/),lines=[];let cur="";for(const w of words){if(tw(cur+w)>maxW&&cur.trim()){lines.push(cur.trim());cur=w.trim()?w:""}else cur+=w}if(cur.trim())lines.push(cur.trim());
 const out=[];for(const ln of lines){if(tw(ln)<=maxW){out.push(ln);continue}let b="";for(const ch of ln){if(tw(b+ch)>maxW){out.push(b);b=ch}else b+=ch}if(b)out.push(b)}return out}
function laneColor(lane){const i=G.laneIdx.get(lane);return i==null?"#888":PAL[i%PAL.length]}
function dateNum(s){return s?Date.parse(s)/864e5:NaN}
function draw(){build();const nodes=[...G.nodes.values()],exts=[...G.ext.values()];let pos=new Map(),W=0,H=0,laneBands=[],axis=null,back=new Set();
 if(S.layout==="time")({W,H,laneBands,axis}=layoutTime(nodes,exts,pos));else({W,H,back}=layoutFlow(nodes,exts,pos));
 drawn={w:W,h:H};
 const chk=$("chk").checked;let s=`<defs><marker id="ar" viewBox="0 0 8 8" refX="8" refY="4" markerWidth="7" markerHeight="7" orient="auto"><path d="M0,0L8,4L0,8z" fill="context-stroke"/></marker></defs>`;
 for(const b of laneBands)s+=`<rect x="0" y="${b.y}" width="${W}" height="${b.h}" fill="${b.i%2?"rgba(127,127,127,.06)":"transparent"}"/><text x="8" y="${b.y+16}" font-size="12" font-weight="bold" fill="${b.color}">${esc(b.name)}</text>`;
 if(laneBands.length&&laneBands.labelW)s+=`<line x1="${laneBands.labelW-2}" y1="0" x2="${laneBands.labelW-2}" y2="${H}" stroke="var(--line)"/>`;
 if(axis){for(const t of axis.ticks)s+=`<line x1="${t.x}" y1="${axis.y}" x2="${t.x}" y2="${H}" stroke="var(--line)" stroke-dasharray="2 4"/><text x="${t.x+3}" y="${axis.y-4}" font-size="11" fill="var(--mut)">${t.label}</text>`}
 const sel=S.sel,upS=new Set(),dnS=new Set();
 if(sel!=null){const adjU=new Map(),adjD=new Map();for(const e of G.edges.values()){(adjU.get(e.b)||adjU.set(e.b,[]).get(e.b)).push(e.a);(adjD.get(e.a)||adjD.set(e.a,[]).get(e.a)).push(e.b)}
  const bfs=(start,adj,out)=>{const st=[start];while(st.length){const v=st.pop();for(const w of (adj.get(v)||[]))if(!out.has(w)&&w!==start){out.add(w);st.push(w)}}};bfs(sel,adjU,upS);bfs(sel,adjD,dnS)}
 const edgesSorted=[...G.edges.values()].sort((a,b)=>(a.cross?1:0)-(b.cross?1:0));
 for(const e of edgesSorted){const A=pos.get(e.a),B=pos.get(e.b);if(!A||!B)continue;
  const isBack=back.has(e.a+">"+e.b);let col=e.ext?"var(--mut)":e.cross?"var(--hand)":"#999",wdt=Math.min(1+Math.log2(e.n),4);
  let cls="ed";if(sel!=null){const on=(e.a===sel&&dnS.has(e.b))||(e.b===sel&&upS.has(e.a))||(upS.has(e.a)&&upS.has(e.b))||(dnS.has(e.a)&&dnS.has(e.b))||(e.a===sel||e.b===sel);if(!on)cls+=" dim";else{if(upS.has(e.a)||e.b===sel&&upS.has(e.a))col="var(--up)";if(dnS.has(e.b)||e.a===sel)col="var(--dn)"}}
  if(isBack)col="var(--hot)";
  const x1=A.x+A.w,y1=A.y+A.h/2,x2=B.x,y2=B.y+B.h/2;let d;
  if(x2>=x1+6)d=`M${x1},${y1} C${(x1+x2)/2},${y1} ${(x1+x2)/2},${y2} ${x2},${y2}`;
  else{const yy=Math.max(A.y+A.h,B.y+B.h)+18;d=`M${x1},${y1} C${x1+30},${y1} ${x1+30},${yy} ${x1},${yy} L${x2},${yy} C${x2-30},${yy} ${x2-30},${y2} ${x2},${y2}`}
  s+=`<path class="${cls}" d="${d}" fill="none" stroke="${col}" stroke-width="${wdt}" ${e.ext||isBack?'stroke-dasharray="5 4"':""} marker-end="url(#ar)"><title>${esc(e.n+"건"+(e.ex.length?"\n"+e.ex.map(x=>x[0]+" → "+x[1]+(x[2]?" "+x[2]:"")).join("\n")+(e.n>e.ex.length?"\n…":""):""))}</title></path>`;
  if(e.n>1)s+=`<text class="${cls}" x="${(x1+x2)/2}" y="${(y1+y2)/2-4}" font-size="10" text-anchor="middle" fill="${col}">${e.n}</text>`}
 for(const n of nodes){const P=pos.get(n.id);if(!P)continue;const col=laneColor(n.lane),me=n.id===sel;let cls="nd";if(sel!=null&&!me&&!upS.has(n.id)&&!dnS.has(n.id))cls+=" dim";if(me)cls+=" sel";
  const crit=chk&&n.cr;s+=`<g class="${cls}" data-n="${n.id}"><rect class="bx" x="${P.x}" y="${P.y}" width="${P.w}" height="${P.h}" rx="5" fill="${me?"var(--sel)":crit?"rgba(220,38,38,.12)":"var(--card)"}" stroke="${crit?"var(--hot)":col}"/>`;
  if(G.lanes.length)s+=`<rect x="${P.x}" y="${P.y}" width="5" height="${P.h}" rx="2.5" fill="${col}"/>`;
  if(P.band)s+=`<rect x="${P.x+1}" y="${P.y+P.h-8}" width="${Math.max(2,P.band-2)}" height="7" rx="2" fill="var(--band)"/>`;
  P.lines.forEach((ln,i)=>{s+=`<text x="${P.tx}" y="${P.y+14+i*15}" font-size="12" fill="currentColor">${esc(ln)}</text>`});
  if(P.sub)s+=`<text x="${P.tx}" y="${P.y+14+P.lines.length*15}" font-size="10.5" fill="var(--mut)">${esc(P.sub)}</text>`;
  s+=`<title>${esc(n.t.wbs+" "+n.t.name+"\n"+pathText(n.id)+"\n"+(n.t.st||"-")+" ~ "+(n.t.fi||"-")+" · 하위 업무 "+n.leaves+" · MH "+Math.round(n.mh).toLocaleString())}</title></g>`}
 for(const x of exts){const P=pos.get(x.id);if(!P)continue;let cls="nd";if(sel!=null&&!upS.has(x.id)&&!dnS.has(x.id))cls+=" dim";
  s+=`<g class="${cls}" data-x="${x.key}"><rect x="${P.x}" y="${P.y}" width="${P.w}" height="${P.h}" rx="5" fill="var(--bg)" stroke="var(--mut)" stroke-dasharray="4 3"/>`;
  P.lines.forEach((ln,i)=>{s+=`<text x="${P.x+8}" y="${P.y+14+i*14}" font-size="11" fill="var(--mut)">${esc(ln)}</text>`});s+=`<title>${esc((x.ext==="in"?"범위 밖에서 옴: ":"범위 밖으로 감: ")+pathText(x.key)+"\n"+x.count+"건 — 두 번 누르면 그쪽으로")}</title></g>`}
 svg.innerHTML=s;
 svg.querySelectorAll("g[data-n]").forEach(g=>{g.onclick=ev=>{ev.stopPropagation();if(!dragged)select(+g.dataset.n)};g.ondblclick=ev=>{ev.stopPropagation();const n=+g.dataset.n;if(T.get(n).kids.length)setScope(n)}});
 svg.querySelectorAll("g[data-x]").forEach(g=>{g.ondblclick=ev=>{ev.stopPropagation();const k=+g.dataset.x;if(T.get(k).kids.length)setScope(k);else setScope(T.get(k).par,6)}});
 svg.onclick=()=>{if(dragged)return;if(S.sel!=null){S.sel=null;keepView=true;draw()}};
 renderCrumb();renderStat();if(sel!=null)renderNode(sel);else renderScope();
 if(!keepView){fit()}keepView=false}
let keepView=false;
function nodeBox(n,maxW){const lines=wrap(n.t.name,maxW-16),w=Math.max(90,Math.min(maxW,Math.max(...lines.map(tw))+16)),sub=`${n.t.dur??"-"}일${n.leaves>1?" · 하위 "+n.leaves:""}`;return{lines,w,h:14+lines.length*15+6+(sub?12:0),sub}}
function layoutTime(nodes,exts,pos){const dated=nodes.filter(n=>n.t.st&&n.t.fi);const d0=dated.length?Math.min(...dated.map(n=>dateNum(n.t.st))):0,d1=dated.length?Math.max(...dated.map(n=>dateNum(n.t.fi))):30;
 // 왼쪽부터: 줄 이름 기둥 | 들어오는 토막 열 | 막대 영역 | 나가는 토막 열 — 서로 겹치지 않게 기둥을 나눈다
 const LABELW=Math.max(60,...G.lanes.map(l=>tw(T.get(l).name)+22)),INW=exts.some(x=>x.ext==="in")?236:0,RIGHT=exts.some(x=>x.ext==="out")?240:40;
 const LEFT=LABELW+INW+12,span=Math.max(30,d1-d0),pxPerDay=Math.max(0.6,Math.min(6,1400/span)),TOP=34;
 const X=d=>LEFT+(d-d0)*pxPerDay;let y=TOP;const laneBands=[];const groups=G.lanes.length?G.lanes.map(l=>nodes.filter(n=>n.lane===l)):[nodes];
 groups.forEach((grp,gi)=>{if(!grp.length)return;const rows=[];     // 빈 줄(거르기로 비었거나 잎 없음)은 그리지 않는다const sorted=grp.slice().sort((a,b)=>(a.t.st?dateNum(a.t.st):-1e9)-(b.t.st?dateNum(b.t.st):-1e9));const y0=y;
  for(const n of sorted){const has=!!(n.t.st&&n.t.fi);const x=has?X(dateNum(n.t.st)):LEFT,durPx=has?Math.max(6,(dateNum(n.t.fi)-dateNum(n.t.st)+1)*pxPerDay):0;
   // 상자 = 이름표 (글자는 늘 안에). 실제 기간은 상자 아래 띠(durPx)로 — 기간이 짧으면 상자가 띠보다 넓다
   const sub=has?`${n.t.st} ~ ${n.t.fi} · ${n.t.dur??"-"}일${n.leaves>1?" · 하위 "+n.leaves:""}`:"날짜 없음";
   const lines=wrap(n.t.name,Math.max(durPx,220)-14),textW=Math.max(tw(sub)*10.5/12,...lines.map(tw))+16,w=Math.max(durPx,textW,60),h=14+lines.length*15+18;
   let r=0;while(rows[r]!=null&&rows[r]>x-8)r++;rows[r]=x+w;const yy=y0+22+r*(h+8);
   pos.set(n.id,{x,y:yy,w,h,lines,tx:x+10,band:durPx,sub});
   y=Math.max(y,yy+h)}
  y+=18;const lane=G.lanes[gi];laneBands.push({i:laneBands.length,y:y0,h:y-y0,name:lane?T.get(lane).name:"",color:lane?laneColor(lane):"var(--mut)"})});
 const W=LEFT+span*pxPerDay+RIGHT+40;
 // 범위 밖 토막: 들어오는 것은 줄 이름 기둥 오른쪽 열, 나가는 것은 오른쪽 끝 열 — 상대 막대 높이에 맞춤
 const place=(list,col)=>{const used=[];for(const x of list.sort((a,b)=>b.count-a.count)){const lines=wrap(`${x.ext==="in"?"← ":"→ "}${x.label} (${x.count})`,200);const h=10+lines.length*14;const tgt=[...G.edges.values()].filter(e=>x.ext==="in"?e.a===x.id:e.b===x.id).map(e=>pos.get(x.ext==="in"?e.b:e.a)).filter(Boolean);
  let yy=tgt.length?Math.min(...tgt.map(p=>p.y)):TOP;while(used.some(u=>Math.abs(u-yy)<h+6))yy+=h+6;used.push(yy);pos.set(x.id,{x:col,y:yy,w:216,h,lines})}};
 place(exts.filter(x=>x.ext==="in"),LABELW+6);place(exts.filter(x=>x.ext==="out"),W-RIGHT+16);
 laneBands.labelW=LABELW;
 const H=Math.max(y,...[...pos.values()].map(p=>p.y+p.h))+30;
 const ticks=[];const start=new Date(d0*864e5);let yr=start.getUTCFullYear();for(;;){const d=Date.UTC(yr,0,1)/864e5;if(d>d1)break;if(d>=d0)ticks.push({x:X(d),label:String(yr)});if(pxPerDay>2.5)for(let m=1;m<12;m+=3){const dm=Date.UTC(yr,m,1)/864e5;if(dm>=d0&&dm<=d1)ticks.push({x:X(dm),label:`${yr}-${String(m+1).padStart(2,"0")}`})}yr++}
 return{W,H,laneBands,axis:{y:TOP-6,ticks}}}
function layoutFlow(nodes,exts,pos){const ids=[...nodes.map(n=>n.id),...exts.map(x=>x.id)],adj=new Map(ids.map(i=>[i,[]]));
 for(const e of G.edges.values())if(adj.has(e.a)&&adj.has(e.b))adj.get(e.a).push(e.b);
 // 거꾸로 가는 간선 제거 (DFS)
 const back=new Set(),state=new Map();const dfs=v=>{state.set(v,1);for(const w of adj.get(v)){if(state.get(w)===1)back.add(v+">"+w);else if(!state.has(w))dfs(w)}state.set(v,2)};for(const v of ids)if(!state.has(v))dfs(v);
 const fwd=new Map(ids.map(i=>[i,adj.get(i).filter(w=>!back.has(i+">"+w))]));
 const indeg=new Map(ids.map(i=>[i,0]));for(const v of ids)for(const w of fwd.get(v))indeg.set(w,indeg.get(w)+1);
 const layer=new Map(),q=ids.filter(i=>!indeg.get(i));for(const v of q)layer.set(v,0);let qi=0;while(qi<q.length){const v=q[qi++];for(const w of fwd.get(v)){layer.set(w,Math.max(layer.get(w)||0,layer.get(v)+1));indeg.set(w,indeg.get(w)-1);if(!indeg.get(w))q.push(w)}}
 for(const v of ids)if(!layer.has(v))layer.set(v,0);
 const maxL=Math.max(0,...layer.values());for(const x of exts)if(x.ext==="out")layer.set(x.id,maxL+1);
 const cols=[];for(const v of ids){const l=layer.get(v);(cols[l]||(cols[l]=[])).push(v)}
 // 순서: 같은 줄끼리 모으고, 앞 열의 평균 위치(barycenter)로
 const ord=new Map();cols.forEach((c,li)=>{c.sort((a,b)=>{const la=G.nodes.get(a)?.lane??-1,lb=G.nodes.get(b)?.lane??-1;const ia=G.laneIdx.get(la)??-1,ib=G.laneIdx.get(lb)??-1;if(ia!==ib)return ia-ib;const ba=bary(a,li),bb=bary(b,li);return ba-bb});c.forEach((v,i)=>ord.set(v,i))});
 function bary(v,li){if(!li)return 0;const ps=[...G.edges.values()].filter(e=>e.b===v&&layer.get(e.a)<li).map(e=>ord.get(e.a)??0);return ps.length?ps.reduce((a,b)=>a+b,0)/ps.length:0}
 const GX=90,colW=[];const boxes=new Map();for(const v of ids){const n=G.nodes.get(v);if(n){const b=nodeBox(n,260);boxes.set(v,b)}else{const x=G.ext.get(v);const lines=wrap(`${x.ext==="in"?"← ":"→ "}${x.label} (${x.count})`,200);boxes.set(v,{lines,w:216,h:10+lines.length*14,sub:""})}}
 cols.forEach((c,li)=>colW[li]=Math.max(...c.map(v=>boxes.get(v).w)));
 let x=30;const colX=[];cols.forEach((c,li)=>{colX[li]=x;x+=colW[li]+GX});const W=x+20;let H=0;
 cols.forEach((c,li)=>{let y=30,prevLane=null;for(const v of c){const b=boxes.get(v),n=G.nodes.get(v);if(n&&prevLane!=null&&n.lane!==prevLane)y+=16;if(n)prevLane=n.lane;pos.set(v,{x:colX[li],y,w:b.w,h:b.h,lines:b.lines,tx:colX[li]+(n&&G.lanes.length?12:8),sub:b.sub});y+=b.h+12}H=Math.max(H,y)});
 return{W,H:H+20,back}}
// ---- 확대·이동 ----
function setVB(){svg.setAttribute("viewBox",`${VB.x} ${VB.y} ${VB.w} ${VB.h}`)}
function fit(){const r=svg.getBoundingClientRect(),sc=Math.max(drawn.w/r.width,drawn.h/r.height,0.5);VB={x:0,y:0,w:r.width*sc,h:r.height*sc};if(sc<1){VB={x:0,y:0,w:r.width,h:r.height}}setVB()}
function zoom(f,cx,cy){const r=svg.getBoundingClientRect();cx=cx??r.width/2;cy=cy??r.height/2;const px=VB.x+cx/r.width*VB.w,py=VB.y+cy/r.height*VB.h;VB.w*=f;VB.h*=f;VB.x=px-cx/r.width*VB.w;VB.y=py-cy/r.height*VB.h;setVB()}
svg.addEventListener("wheel",ev=>{ev.preventDefault();const r=svg.getBoundingClientRect();zoom(ev.deltaY>0?1.15:1/1.15,ev.clientX-r.left,ev.clientY-r.top)},{passive:false});
let drag=null,dragged=false;   // 끌어 옮긴 뒤 따라오는 click 은 무시한다 (끌기 = 선택 해제·맞춤이 아니다)
svg.onmousedown=ev=>{drag={x:ev.clientX,y:ev.clientY,vx:VB.x,vy:VB.y};dragged=false};
window.onmousemove=ev=>{if(!drag)return;if(Math.abs(ev.clientX-drag.x)+Math.abs(ev.clientY-drag.y)>3)dragged=true;const r=svg.getBoundingClientRect();VB.x=drag.vx-(ev.clientX-drag.x)*VB.w/r.width;VB.y=drag.vy-(ev.clientY-drag.y)*VB.h/r.height;setVB()};
window.onmouseup=()=>{drag=null;setTimeout(()=>dragged=false,0)};
$("zi").onclick=()=>zoom(1/1.3);$("zo").onclick=()=>zoom(1.3);$("zf").onclick=fit;window.onresize=fit;
function select(n){S.sel=n;keepView=true;draw()}
// ---- 빵부스러기·패널 ----
function renderCrumb(){let h=`<span class="c" data-n="0">전체</span>`;for(const t of (S.scope?path(S.scope):[]))h+=` › <span class="c" data-n="${t.no}">${esc(t.name)}</span>`;h+=` <span class="mut">(Lv${S.grain} ${LVNAME[S.grain]} 알갱이${S.fLv&&S.fName?` · Lv${S.fLv} '${esc(S.fName)}' 만`:""})</span>`;$("crumb").innerHTML=h;$("crumb").querySelectorAll(".c").forEach(c=>c.onclick=()=>setScope(+c.dataset.n))}
function renderStat(){$("stat").textContent=`상자 ${G.nodes.size} · 선 ${[...G.edges.values()].filter(e=>!e.ext).length} (뭉쳐진 안쪽 ${G.internal}) · 밖에서 옴 ${G.inN} · 밖으로 감 ${G.outN}`}
function fmtMH(v){return Math.round(v).toLocaleString()}
function renderScope(){const sc=S.scope?T.get(S.scope):null;const name=sc?sc.name:"전체";const ds=[...G.nodes.values()].map(n=>n.t.st).filter(Boolean).sort(),de=[...G.nodes.values()].map(n=>n.t.fi).filter(Boolean).sort();
 const hand=[...G.handoff.entries()].sort((a,b)=>b[1]-a[1]).slice(0,8).map(([k,v])=>{const [a,b]=k.split(">").map(Number);return `<tr><td>${esc(T.get(a).name)} → ${esc(T.get(b).name)}</td><td>${v}</td></tr>`}).join("");
 const chain=G.chain.map(no=>`<span class="lnk" data-n="${no}">${esc(T.get(no).name)}</span>`).join(" → ");
 let h=`<h3>${esc(name)}${sc?` <span class="pill">Lv${sc.lv} ${LVNAME[sc.lv]}</span>`:""}${S.fLv&&S.fName?` <span class="pill">Lv${S.fLv} '${esc(S.fName)}' 만</span>`:""}</h3>${sc?`<div class="mut">${esc(pathText(sc.no))}</div>`:""}
 <table><tr><td>하위 업무</td><td>${G.leaves.toLocaleString()}</td></tr><tr><td>MH 합(잎)</td><td>${fmtMH([...G.nodes.values()].reduce((a,n)=>a+n.mh,0))}</td></tr>
 <tr><td>기간</td><td>${ds[0]||"-"} ~ ${de[de.length-1]||"-"}</td></tr><tr><td>넘겨받기</td><td>${[...G.handoff.values()].reduce((a,b)=>a+b,0)}건 (줄을 넘는 선행)</td></tr>
 <tr><td>범위 밖</td><td>들어옴 ${G.inN} · 나감 ${G.outN}</td></tr></table>`;
 if(hand)h+=`<h3>넘겨받기 (많은 순)</h3><table>${hand}</table>`;
 h+=`<h3>범위 안 가장 긴 사슬 <span class="mut">${G.chainDays} 일 (기간 합, 달력 없음)</span></h3><div style="font-size:12px">${chain||"-"}</div>`;
 const kids=sc?sc.kids:ROOTS;if(kids.length)h+=`<h3>바로 아래 (${kids.length})</h3><table>${kids.map(k=>`<tr><td><span class="lnk" data-s="${k}">${esc(T.get(k).name)}</span></td><td>${LEAF.get(k).toLocaleString()}</td></tr>`).join("")}</table>`;
 $("side").innerHTML=h;bindSide()}
function renderNode(no){const t=T.get(no),n=G.nodes.get(no);const under=x=>anc(x,t.lv)===no;   // 뭉친 상자면 그 아래 잎으로 드나드는 원본 행
 const preE=D.edges.filter(e=>under(e[1])&&!under(e[0])),sucE=D.edges.filter(e=>under(e[0])&&!under(e[1]));
 const row=(e,o,i)=>`<span class="lnk" data-n="${e[o]}">${e[o]}</span>${esc(e[2])} ${esc(T.get(e[o]).name)}${t.kids.length?` <span class="mut">(${esc(T.get(e[i]).name)} 쪽)</span>`:""}`;
 const pre=preE.slice(0,40).map(e=>row(e,0,1)).join("<br>")+(preE.length>40?`<br>… ${preE.length-40}건 더`:"")||"없음";
 const suc=sucE.slice(0,40).map(e=>row(e,1,0)).join("<br>")+(sucE.length>40?`<br>… ${sucE.length-40}건 더`:"")||"없음";
 const agg=[...G.edges.values()].filter(e=>e.a===no||e.b===no).map(e=>{const other=e.a===no?e.b:e.a,on=G.nodes.get(other)?G.nodes.get(other).t.name:G.ext.get(other)?.label;return `<tr><td>${e.a===no?"→":"←"}</td><td>${esc(on)} <span class="mut">${e.n}건</span></td></tr>`}).join("");
 const iss=D.issues.filter(i=>i[3]===no).map(i=>`<div class="${i[1]==="오류"?"err":""}">${i[0]} ${esc(i[2])}${i[5]?" — "+esc(i[5]):""}</div>`).join("");
 let h=`<h3>${t.no} · ${esc(t.name)} <span class="pill">Lv${t.lv}</span>${t.cr?'<span class="pill">가장 긴 사슬</span>':""}</h3><div class="mut">${esc(t.wbs)} · ${esc(path(no).slice(0,-1).map(x=>x.name).join(" › "))}</div>
 <table><tr><td>기간</td><td>${t.dur??"-"}일 · ${t.st||"-"} ~ ${t.fi||"-"}</td></tr><tr><td>여유</td><td>${t.fl} 근무일 (근사)</td></tr><tr><td>앞쪽·뒤쪽 전체</td><td>${t.up} · ${t.dn}</td></tr>
 ${n?`<tr><td>하위 업무</td><td>${n.leaves} · MH ${fmtMH(n.mh)}</td></tr>`:""}${t.staff!=null?`<tr><td>투입인력</td><td>${t.staff}</td></tr>`:""}${t.del?`<tr><td>산출물</td><td>${esc(t.del)}</td></tr>`:""}</table>`;
 if(t.kids.length)h+=`<div style="margin:6px 0"><span class="lnk" data-s="${no}">이 안으로 들어가기 (${t.kids.length}개)</span></div>`;
 if(agg)h+=`<h3>이 상자의 연결 (뭉친 것)</h3><table>${agg}</table>`;
 h+=`<h3>선행 (원본 행 ${preE.length})</h3><div style="font-size:12px">${pre}</div><h3>후속 (원본 행 ${sucE.length})</h3><div style="font-size:12px">${suc}</div>`;if(iss)h+=`<h3>검사</h3>${iss}`;
 $("side").innerHTML=h;bindSide()}
function bindSide(){$("side").querySelectorAll(".lnk[data-n]").forEach(x=>x.onclick=()=>goTask(+x.dataset.n));$("side").querySelectorAll(".lnk[data-s]").forEach(x=>x.onclick=()=>setScope(+x.dataset.s))}
function goTask(no){const t=T.get(no);const L=scopeLevel();const inside=!S.scope||anc(no,L)===S.scope;const node=anc(no,S.grain);if(inside&&G.nodes.has(node)){select(node)}else{setScope(t.kids.length?no:t.par,6);select(no)}}
// ---- 검사 ----
const codes=[...new Set(D.issues.map(i=>i[0]))];
$("fIs").innerHTML='<option value="">(전체)</option>'+codes.map(c=>{const i=D.issues.find(x=>x[0]===c);return `<option value="${c}">${c} ${esc(i[2])} (${D.issues.filter(x=>x[0]===c).length})</option>`}).join("");
$("fIs").onchange=drawIs;
function drawIs(){const c=$("fIs").value,rows=D.issues.filter(i=>!c||i[0]===c);$("isN").textContent=`${rows.length}건${rows.length>3000?" (앞 3,000건만 — 전체는 보고서 xlsx)":""}`;
 let h='<tr><th>수준</th><th>코드</th><th>이름</th><th>번호</th><th>WBS</th><th>내용</th></tr>';
 for(const i of rows.slice(0,3000))h+=`<tr><td class="${i[1]==="오류"?"err":""}">${i[1]}</td><td>${i[0]}</td><td>${esc(i[2])}</td><td>${i[3]!=null&&T.has(i[3])?`<span class="lnk" data-n="${i[3]}">${i[3]}</span>`:(i[3]??"")}</td><td>${esc(i[4])}</td><td>${esc(i[5])}</td></tr>`;
 $("isT").innerHTML=h;$("isT").querySelectorAll(".lnk").forEach(x=>x.onclick=()=>{document.querySelector('nav button[data-t="gr"]').click();goTask(+x.dataset.n)})}
drawIs();
for(const r of ROOTS)open.add(r);
setScope(0);
</script></body></html>
"""
