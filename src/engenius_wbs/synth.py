"""가짜 WBS 엑셀 만들기 — 실제 WBS 와 같은 칸·형식, 일반 명칭만 (구성품·업무·단계·분야는 교과서 수준 이름).

모양: A 번호 · B WBS(1.2.1.1.2.1) · C 작업 이름(앞 띄어쓰기 3칸/레벨) · D 기간('12 일') · E·F 날짜 글자('27-01-01 (금)') ·
G 선행(쉼표, FS 기본·SS+5일·FF 섞음) · H 후속(G 를 뒤집은 것) · I 작성자(자리표시 '가나다') · J 투입인력 · K Man Hours · L 산출물(주로 Lv5).
선행은 늘 앞 번호를 가리키게 만들어 순환이 없고, 날짜는 선행으로 계산한다 (주말 건너뜀, 휴일 없음).
그 위에 검사가 찾아야 할 문제를 일부러 심고 `injected` 로 돌려준다 (테스트 골든).
"""
from __future__ import annotations

import random
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

MODULES = ["Fan", "LPC", "HPC", "연소기", "HPT", "LPT", "체계종합"]
TOPS = ["사업관리", "설계통합", "설계해석", "시험평가"]
PHASES = ["사전기본설계", "기본설계", "상세설계"]
DISCIPLINES = ["성능", "공력", "구조", "이차유로", "열전달", "재료", "소음"]
WP_NAMES = ["요구도 분석", "설계 조건 정의", "형상 설계", "해석", "결과 검토", "설계 문서화"]
TASK_NAMES = ["입력 자료 수집", "조건 정리", "모델 준비", "계산 수행", "결과 정리", "검토 회의", "수정 반영", "자료 배포", "점검"]
HEADERS = ["번호", "WBS", "작업 이름", "기간", "시작", "완료", "선행 작업", "후속 작업", "작성자", "투입인력(명)", "Man Hours", "산출물"]
START = date(2027, 1, 1)
WD = "월화수목금토일"


@dataclass
class Node:
    code: str
    name: str
    level: int
    children: list = field(default_factory=list)
    no: int = 0
    deps: list = field(default_factory=list)     # [(번호, 표기 접미사)]
    dur: int = 0
    es: int = 0                                  # 근무일 0 = START
    staff: float | None = None
    mh: float | None = None
    deliv: str = ""
    path: tuple = ()                             # (구성품, 최상위, 단계, 분야) 색인


def workday(n: int) -> date:
    """근무일 n (0 = START, 주말 건너뜀) 의 날짜."""
    d = START
    while d.weekday() >= 5:
        d += timedelta(days=1)
    weeks, rest = divmod(n, 5)
    d += timedelta(days=7 * weeks)
    while rest:
        d += timedelta(days=1)
        if d.weekday() < 5:
            rest -= 1
    return d


def date_text(d: date, compact: bool = False) -> str:
    return f"{d.year % 100:02d}-{d.month:02d}-{d.day:02d}{'' if compact else ' '}({WD[d.weekday()]})"


def build_tree(rng: random.Random, rows: int):
    roots = []
    wps = []                                     # (Node, path)
    for mi, m in enumerate(MODULES):
        n1 = Node(f"{mi + 1}", m, 1)
        roots.append(n1)
        for ti, top in enumerate(TOPS):
            n2 = Node(f"{n1.code}.{ti + 1}", top, 2)
            n1.children.append(n2)
            for pi, ph in enumerate(PHASES):
                n3 = Node(f"{n2.code}.{pi + 1}", ph, 3)
                n2.children.append(n3)
                for di, dis in enumerate(DISCIPLINES):
                    n4 = Node(f"{n3.code}.{di + 1}", dis, 4)
                    n3.children.append(n4)
                    for wi in range(rng.randint(2, 4)):
                        n5 = Node(f"{n4.code}.{wi + 1}", f"{dis} {WP_NAMES[wi % len(WP_NAMES)]}", 5, path=(mi, ti, pi, di))
                        n4.children.append(n5)
                        wps.append(n5)
    def count():
        return sum(1 for _ in walk(roots))
    for wp in wps:                               # 하위 업무 최소 2개
        for _ in range(2):
            add_task(wp, rng)
    total = count()
    while total < rows:
        add_task(rng.choice(wps), rng)
        total += 1
    while total > rows:                          # 너무 많으면 하위 업무가 3개 이상인 묶음에서 뺀다
        wp = rng.choice(wps)
        if len(wp.children) > 2:
            wp.children.pop()
            total -= 1
    return roots, wps


def add_task(wp: Node, rng: random.Random) -> None:
    k = len(wp.children)
    wp.children.append(Node(f"{wp.code}.{k + 1}", f"{TASK_NAMES[k % len(TASK_NAMES)]}" + (f" {k // len(TASK_NAMES) + 1}" if k >= len(TASK_NAMES) else ""), 6))


def walk(nodes):
    for n in nodes:
        yield n
        yield from walk(n.children)


def generate(path: Path, rows: int = 12827, seed: int = 7) -> dict:
    rng = random.Random(seed)
    roots, wps = build_tree(rng, rows)
    order = list(walk(roots))
    for i, n in enumerate(order, start=1):
        n.no = i
    # 선행: 묶음 안 사슬 + 다음 묶음 + 다음 단계 같은 분야 + (가끔) 앞 분야에서 온 것
    last_of = {}                                 # path → 그 상위 업무의 마지막 하위 업무
    by_path = {}
    for wp in wps:
        by_path.setdefault(wp.path[:4], []).append(wp)
    for wp in wps:
        kids = wp.children
        join = {}                                # 나란한 두 업무 (j-1, j) 를 j+1 이 같이 기다림 — 짧은 쪽에 여유가 생긴다 (준주경로)
        for j, t in enumerate(kids):
            t.dur = rng.randint(3, 40)
            t.staff = float(rng.randint(1, 4))
            t.mh = t.staff * t.dur * 8
            if j >= 2 and j + 1 < len(kids) and j - 1 not in join and rng.random() < 0.2:
                t.deps.append((kids[j - 2].no, ""))
                t.dur = max(3, kids[j - 1].dur - rng.randint(0, 12))
                t.mh = t.staff * t.dur * 8
                join[j] = kids[j - 1].no
            elif j:
                r = rng.random()
                t.deps.append((kids[j - 1].no, "" if r < 0.85 else ("SS+5일" if r < 0.93 else "FF")))
            if j - 1 in join:
                t.deps.append((join[j - 1], ""))
        if len(kids) >= 3 and rng.random() < 0.12:              # 마일스톤 (기간 0) — 묶음의 마지막 업무, MH 없음
            ms = kids[-1]
            ms.dur, ms.staff, ms.mh, ms.name = 0, None, None, ms.name + " 완료"
        mi, ti, pi, di = wp.path
        siblings = by_path[wp.path]
        idx = siblings.index(wp)
        first = kids[0]
        if idx:
            first.deps.append((siblings[idx - 1].children[-1].no, ""))
        elif pi:
            prev = by_path.get((mi, ti, pi - 1, di))
            if prev:
                first.deps.append((prev[-1].children[-1].no, "" if rng.random() < 0.9 else "FS+10일"))
        if di and rng.random() < 0.35:
            other = by_path.get((mi, ti, pi, rng.randrange(di)))
            if other:
                src = rng.choice(other).children[-1]
                if src.no < first.no:
                    first.deps.append((src.no, ""))
        if mi and rng.random() < 0.04:
            other = by_path.get((rng.randrange(mi), ti, pi, di))
            if other:
                first.deps.append((other[-1].children[-1].no, ""))
        wp.deliv = f"{DISCIPLINES[di]} {PHASES[pi]} {wp.name.split(' ', 1)[1]} 보고서"
        if rng.random() < 0.15:
            wp.deliv += f", {DISCIPLINES[di]} 검토 자료"
    # 날짜: 번호 순으로 앞에서부터 (선행은 늘 앞 번호)
    by_no = {n.no: n for n in order}
    for n in order:
        if n.level != 6:
            continue
        es = 0
        for p, suf in n.deps:
            q = by_no[p]
            lag = 5 if "+5" in suf else 10 if "+10" in suf else 0
            if suf.startswith("SS"):
                es = max(es, q.es + lag)
            elif suf == "FF":
                es = max(es, q.es + q.dur - n.dur)
            else:
                es = max(es, q.es + q.dur + lag)
        n.es = es
    def roll(n):
        if n.level == 6:
            return n.es, n.es + n.dur, n.staff or 0, n.mh or 0
        spans = [roll(c) for c in n.children]
        n.es = min(s[0] for s in spans)
        end = max(s[1] for s in spans)
        n.dur = end - n.es
        n.staff = sum(s[2] for s in spans) if n.level == 5 else None
        n.mh = sum(s[3] for s in spans)
        return n.es, end, sum(s[2] for s in spans), n.mh
    for r in roots:
        roll(r)

    rows_out = []
    for n in order:
        if n.level == 6 and n.dur == 0:                         # 마일스톤: 선행이 끝난 날 (MS Project 표시와 같음)
            st = fi = workday(max(n.es - 1, 0))
        else:
            st, fi = workday(n.es), workday(n.es + max(n.dur, 1) - 1)
        compact = n.no % 9 == 0
        dur = f"{n.dur} 일" if n.no % 11 else f"{n.dur}일"
        if n.level == 6 and n.no % 97 == 0:
            dur += "?"
        rows_out.append({"no": n.no, "wbs": n.code, "name": "   " * (n.level - 1) + n.name, "dur": dur,
                         "start": date_text(st, compact), "finish": date_text(fi, compact),
                         "preds": ",".join(f"{p}{s}" for p, s in n.deps), "author": "가나다" if n.no % 50 == 0 else None,
                         "staff": n.staff if n.level >= 5 else None, "mh": n.mh, "deliv": n.deliv})

    injected = inject(rows_out, order, by_no, rng)
    undated = injected.pop("undated", [])
    succ = {}
    for r in rows_out:
        for tok in filter(None, str(r["preds"]).split(",")):
            head = ""
            for ch in tok:
                if not ch.isdigit():
                    break
                head += ch
            if head:
                succ.setdefault(int(head), []).append(r["no"])
    for r in rows_out:
        r["succs"] = ",".join(map(str, sorted(succ.get(r["no"], []))))
    no_h = injected["C07"][0]
    rows_out[no_h - 1]["succs"] = (rows_out[no_h - 1]["succs"] + ",99999").strip(",")

    wb = Workbook()
    ws = wb.active
    ws.title = "WBS"
    ws.append(HEADERS)
    for r in rows_out:
        ws.append([r["no"], r["wbs"], r["name"], r["dur"], r["start"], r["finish"], r["preds"] or None, r["succs"] or None,
                   r["author"], r["staff"], r["mh"], r["deliv"] or None])
    for col, wdt in zip("ABCDEFGHIJKL", [7, 16, 50, 9, 14, 14, 18, 18, 8, 9, 10, 36]):
        ws.column_dimensions[col].width = wdt
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    milestones = [n.no for n in order if n.level == 6 and n.dur == 0]
    return {"rows": len(rows_out), "injected": injected, "undated": undated, "milestones": milestones}


def inject(rows: list[dict], order: list[Node], by_no: dict, rng: random.Random) -> dict:
    """검사가 찾아야 할 문제를 심는다 → {코드: [업무 번호, …]} (C07 은 H 열을 쓴 뒤 심는다)."""
    leaves = [n for n in order if n.level == 6]
    wps = [n for n in order if n.level == 5]
    pick = iter(rng.sample(range(len(leaves) // 10, len(leaves) - 10), 40))
    def leaf():
        return leaves[next(pick)]
    def add(no, tok):
        r = rows[no - 1]
        r["preds"] = (str(r["preds"]) + "," + tok).strip(",") if r["preds"] else tok
    out = {k: [] for k in ("C01", "C02", "C03", "C04", "C05", "C07", "C08", "C10", "C11", "C12", "P07", "P09", "P11")}
    firsts = {w.children[0].no for w in wps}
    for _ in range(2):
        t = leaf()
        add(t.no, str(len(rows) + 1000 + t.no))
        out["C01"].append(t.no)
    t = leaf()
    add(t.no, str(t.no))
    out["C02"].append(t.no)
    wp = next(w for w in wps[len(wps) // 2:] if len(w.children) >= 4)        # 순환: 첫 업무가 셋째를 선행으로
    a, c = wp.children[0], wp.children[2]
    add(a.no, str(c.no))
    out["C03"].append(min(a.no, wp.children[1].no, c.no))
    out["C05"].append(a.no)                                                  # 순환 간선은 날짜도 어긋난다
    t = leaf()
    other = next(w for w in wps if w.no < t.no and w.no != by_no[t.no].no and t.no not in [k.no for k in w.children]
                 and w.es + w.dur <= t.es)                                   # 날짜는 맞게 (C05 를 따로 만들지 않음)
    add(t.no, str(other.no))
    out["C04"].append(t.no)
    wp2 = next(w for w in wps[len(wps) // 3:] if len(w.children) >= 3 and w is not wp)
    kid = wp2.children[1]
    add(kid.no, str(wp2.no))
    out["C10"].append(kid.no)
    for _ in range(2):                                                       # 날짜 충돌: FS 선행이 있는 업무의 시작을 앞당김
        t = leaf()
        while not t.deps or t.deps[0][1] or t.no in out["C05"] or t.no in firsts or t.dur == 0:   # 첫 업무를 당기면 묶음 날짜도 바뀐다 (C11)
            t = leaf()
        p = by_no[t.deps[0][0]]
        early = workday(max(p.es + p.dur - 3, 0))
        rows[t.no - 1]["start"] = date_text(early)
        out["C05"].append(t.no)
    t = leaf()
    out["C07"].append(t.no)
    wp3 = wps[-5]                                                            # 고립 하위 업무: 사슬 밖에 하나 더 — 번호를 새로 줄 수 없으니 마지막 업무의 선행을 지움
    lone = wp3.children[-1]
    rows[lone.no - 1]["preds"] = ""
    for r in rows:                                                           # 그 업무를 선행으로 둔 것도 지움
        if r["preds"]:
            toks = [x for x in str(r["preds"]).split(",") if not (x == str(lone.no) or (x.startswith(str(lone.no)) and not x[len(str(lone.no))].isdigit()))]
            r["preds"] = ",".join(toks)
    out["C08"].append(lone.no)
    t = leaf()
    rows[t.no - 1]["name"] = rows[t.no - 1]["name"].lstrip()                 # 들여쓰기 없앰
    out["P07"].append(t.no)
    t = leaf()
    s = rows[t.no - 1]["start"]
    wrong = WD[(WD.index(s[-2]) + 1) % 7]
    rows[t.no - 1]["start"] = s[:-2] + wrong + ")"
    out["P09"].append(t.no)
    t = leaf()
    add(t.no, "약 12")
    out["P11"].append(t.no)
    used = {no for v in out.values() for no in v}
    ph = [n for n in order if n.level == 3]                                  # 요약 날짜 손댐: 개발 단계 하나의 완료를 1주 뒤로
    n3 = ph[len(ph) // 2]
    fi = workday(n3.es + max(n3.dur, 1) - 1) + timedelta(days=7)
    rows[n3.no - 1]["finish"] = date_text(fi)
    out["C11"].append(n3.no)
    n4 = next(n for n in order if n.level == 4 and n.no > n3.no + 200)      # 요약 MH 손댐
    rows[n4.no - 1]["mh"] = (rows[n4.no - 1]["mh"] or 0) + 120
    out["C12"].append(n4.no)
    for wp in wps[len(wps) * 2 // 3:]:                                       # 날짜 없는 업무: 묶음 가운데 업무 (가장자리는 묶음 날짜를 바꾼다)
        mid = wp.children[1:-1]
        cand = [k for k in mid if k.no not in used and k.dur and not k.deps[1:]]
        if cand:
            rows[cand[0].no - 1]["start"] = rows[cand[0].no - 1]["finish"] = None
            out["undated"] = [cand[0].no]
            break
    return out
