"""그래프 검사·계산 — 선행 관계 확인(C01~C10), 앞쪽·뒤쪽 전체 수, 달력 없는 CPM(주경로), 분야 간 연결.

선행 관계는 업무 번호(A열) 사이. 상위 묶음(요약 작업)이 선행에 걸린 것도 그대로 둔다 (MS Project 가 허용한다).
CPM 은 업무마다 '시작·완료' 사건 두 개를 두고 근무일로 계산한다 — 휴일·달력은 모른다 (엑셀 날짜와 다를 수 있다, 근사).
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta

from .parse import Issue, Task, Wbs


@dataclass
class Edge:
    pred: int
    succ: int
    type: str
    lag_days: float | None
    lag_text: str


@dataclass
class Analysis:
    edges: list[Edge]                       # 두 끝이 모두 있는 선행 관계 (자기 자신 제외)
    issues: list[Issue]                     # C 코드
    cycles: list[list[int]]                 # 순환 묶음 (업무 번호)
    up_count: dict[int, int]                # 앞쪽 전체 (순환 안 업무는 빠짐)
    down_count: dict[int, int]
    preds_of: dict[int, list[int]]
    succs_of: dict[int, list[int]]
    es: dict[int, float] = field(default_factory=dict)     # CPM 이른 시작 (근무일, 프로젝트 시작 = 0)
    ef: dict[int, float] = field(default_factory=dict)
    total_float: dict[int, float] = field(default_factory=dict)
    project_days: float = 0.0
    critical_chain: list[int] = field(default_factory=list)  # 기간 합이 가장 긴 하위 업무 사슬 (순서대로)
    cpm_dropped_edges: int = 0              # 순환 때문에 CPM 에서 뺀 사건 간선 수
    discipline_flows: Counter = field(default_factory=Counter)   # (선행 분야, 후속 분야) → 건수


def weekdays_between(a: date, b: date) -> int:
    """a~b (양끝 포함) 의 월~금 날 수."""
    if b < a:
        return 0
    days = (b - a).days + 1
    weeks, rest = divmod(days, 7)
    n = weeks * 5
    for i in range(rest):
        if (a + timedelta(days=weeks * 7 + i)).weekday() < 5:
            n += 1
    return n


def tarjan(nodes, adj) -> list[list]:
    """강연결 요소 (반복형). 크기 2 이상 또는 자기 고리만 순환으로 쓴다."""
    index, low, on, stack, out = {}, {}, set(), [], []
    counter = 0
    for root in nodes:
        if root in index:
            continue
        work = [(root, iter(adj.get(root, ())))]
        index[root] = low[root] = counter
        counter += 1
        stack.append(root)
        on.add(root)
        while work:
            v, it = work[-1]
            advanced = False
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter
                    counter += 1
                    stack.append(w)
                    on.add(w)
                    work.append((w, iter(adj.get(w, ()))))
                    advanced = True
                    break
                if w in on:
                    low[v] = min(low[v], index[w])
            if advanced:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[v])
            if low[v] == index[v]:
                comp = []
                while True:
                    w = stack.pop()
                    on.discard(w)
                    comp.append(w)
                    if w == v:
                        break
                out.append(comp)
    return out


def topo(nodes, adj) -> list:
    indeg = Counter()
    for v in nodes:
        for w in adj.get(v, ()):
            indeg[w] += 1
    order = [v for v in nodes if indeg[v] == 0]
    i = 0
    while i < len(order):
        for w in adj.get(order[i], ()):
            indeg[w] -= 1
            if indeg[w] == 0:
                order.append(w)
        i += 1
    return order


def analyze(w: Wbs) -> Analysis:
    tasks = w.tasks
    issues: list[Issue] = []
    edges: list[Edge] = []

    def iss(code, t: Task, detail=""):
        issues.append(Issue(code, t.no, t.excel_row, t.wbs, detail))

    def ancestors(no):
        p = tasks[no].parent
        while p is not None:
            yield p
            p = tasks[p].parent

    nested = set()                          # 상위 묶음 ↔ 그 하위 업무 사이 선행 (날짜 비교가 뜻이 없다)
    for t in tasks.values():
        for d in t.deps:
            if d.pred == t.no:
                iss("C02", t)
                continue
            if d.pred not in tasks:
                iss("C01", t, f"선행 {d.pred}")
                continue
            p = tasks[d.pred]
            if d.pred in set(ancestors(t.no)) or t.no in set(ancestors(d.pred)):
                iss("C10", t, f"선행 {d.pred} ({p.wbs})")
                nested.add((d.pred, t.no))
            elif p.is_summary or t.is_summary:
                which = "선행이 상위 묶음" if p.is_summary else "이 업무가 상위 묶음"
                iss("C04", t, f"선행 {d.pred} ({p.wbs}) — {which}")
            edges.append(Edge(d.pred, t.no, d.type, d.lag_days, d.lag_text))

    preds_of, succs_of = defaultdict(list), defaultdict(list)
    for e in edges:
        preds_of[e.succ].append(e.pred)
        succs_of[e.pred].append(e.succ)

    # 순환
    nodes = list(tasks)
    comps = tarjan(nodes, succs_of)
    cycles = [sorted(c) for c in comps if len(c) > 1]
    in_cycle = {n for c in cycles for n in c}
    for c in cycles:
        first = tasks[c[0]]
        shown = ", ".join(map(str, c[:12])) + (f" 외 {len(c) - 12}" if len(c) > 12 else "")
        iss("C03", first, f"{len(c)}개 업무: {shown}")

    # 날짜
    for e in edges:
        p, s = tasks[e.pred], tasks[e.succ]
        if (e.lag_days is not None and e.lag_days < 0) or (e.pred, e.succ) in nested:
            continue
        bad = None
        if e.type == "FS" and p.finish and s.start and s.start < p.finish:
            bad = f"FS: 시작 {s.start} < 선행 {p.no} 완료 {p.finish}"
        elif e.type == "SS" and p.start and s.start and s.start < p.start:
            bad = f"SS: 시작 {s.start} < 선행 {p.no} 시작 {p.start}"
        elif e.type == "FF" and p.finish and s.finish and s.finish < p.finish:
            bad = f"FF: 완료 {s.finish} < 선행 {p.no} 완료 {p.finish}"
        elif e.type == "SF" and p.start and s.finish and s.finish < p.start:
            bad = f"SF: 완료 {s.finish} < 선행 {p.no} 시작 {p.start}"
        if bad:
            iss("C05", s, bad)
    for t in tasks.values():
        if t.start and t.finish:
            if t.start > t.finish:
                iss("C09", t, f"{t.start} > {t.finish}")
            elif t.duration_days is not None and not t.is_summary and t.duration_days >= 1:
                wd = weekdays_between(t.start, t.finish)
                if wd < t.duration_days or wd > t.duration_days * 1.5 + 5:
                    iss("C06", t, f"기간 {t.duration_text} · 평일 {wd}일")

    # 후속 열 검산
    derived_succs = defaultdict(set)
    for t in tasks.values():
        for d in t.deps:
            derived_succs[d.pred].add(t.no)
    if any(t.succs_declared is not None for t in tasks.values()):
        for t in tasks.values():
            if t.succs_declared is None:
                continue
            decl, calc = set(t.succs_declared), derived_succs.get(t.no, set())
            if decl != calc:
                more, less = sorted(decl - calc), sorted(calc - decl)
                parts = []
                if more:
                    parts.append("H열에만 " + ", ".join(map(str, more[:10])))
                if less:
                    parts.append("G열 역방향에만 " + ", ".join(map(str, less[:10])))
                iss("C07", t, " · ".join(parts))

    # 고립 하위 업무
    for t in tasks.values():
        if not t.is_summary and t.level >= 6 and not preds_of.get(t.no) and not succs_of.get(t.no):
            iss("C08", t)

    # 앞쪽·뒤쪽 전체 (순환 안 업무는 순환 간선을 빼고 셈 — 근사)
    dag = {v: [w2 for w2 in succs_of.get(v, ()) if not (v in in_cycle and w2 in in_cycle)] for v in nodes}
    order = topo(nodes, dag)
    bit = {v: 1 << i for i, v in enumerate(nodes)}
    up = {v: 0 for v in nodes}
    for v in order:
        for w2 in dag[v]:
            up[w2] |= up[v] | bit[v]
    down = {v: 0 for v in nodes}
    for v in reversed(order):
        for w2 in dag[v]:
            down[v] |= down[w2] | bit[w2]
    up_count = {v: up[v].bit_count() for v in nodes}
    down_count = {v: down[v].bit_count() for v in nodes}

    a = Analysis(edges, issues, cycles, up_count, down_count, dict(preds_of), dict(succs_of))
    cpm(tasks, edges, a)

    for e in edges:
        pd, sd = tasks[e.pred].discipline, tasks[e.succ].discipline
        if pd and sd:
            a.discipline_flows[(pd, sd)] += 1
    issues.sort(key=lambda i: (i.code, i.no or 0))
    return a


def cpm(tasks: dict[int, Task], edges: list[Edge], a: Analysis) -> None:
    """사건 그래프 CPM: 업무 t → s(t)·f(t). 하위 업무 s→f 가중치 = 기간, 상위 묶음은 s(X)→s(자식)·f(자식)→f(X) 로 감싼다."""
    adj = defaultdict(list)            # (사건) → [(사건, 가중치)]
    for t in tasks.values():
        s, f = ("s", t.no), ("f", t.no)
        adj[s].append((f, 0.0 if t.is_summary else float(t.duration_days or 0)))
        for c in t.children:
            adj[s].append((("s", c), 0.0))
            adj[("f", c)].append((f, 0.0))
    for e in edges:
        lag = e.lag_days or 0.0
        src = ("f", e.pred) if e.type in ("FS", "FF") else ("s", e.pred)
        dst = ("s", e.succ) if e.type in ("FS", "SS") else ("f", e.succ)
        adj[src].append((dst, lag))
    events = list(dict.fromkeys([ev for t in tasks.values() for ev in (("s", t.no), ("f", t.no))]))
    plain = {v: [w for w, _ in adj.get(v, ())] for v in events}
    comps = tarjan(events, plain)
    comp_of = {}
    for i, c in enumerate(comps):
        for v in c:
            comp_of[v] = i
    dropped = 0
    dag = defaultdict(list)
    for v in events:
        for w, wt in adj.get(v, ()):
            if comp_of[v] == comp_of[w]:
                dropped += 1
                continue
            dag[v].append((w, wt))
    order = topo(events, {v: [w for w, _ in dag.get(v, ())] for v in events})
    early = {v: 0.0 for v in events}
    best_from = {}
    for v in order:
        for w, wt in dag.get(v, ()):
            if early[v] + wt > early[w]:
                early[w] = early[v] + wt
                best_from[w] = v
    project = max(early.values(), default=0.0)
    late = {v: project for v in events}
    for v in reversed(order):
        for w, wt in dag.get(v, ()):
            late[v] = min(late[v], late[w] - wt)
    for t in tasks.values():
        a.es[t.no] = early[("s", t.no)]
        a.ef[t.no] = early[("f", t.no)]
        a.total_float[t.no] = round(late[("s", t.no)] - early[("s", t.no)], 6)
    a.project_days = project
    a.cpm_dropped_edges = dropped
    # 가장 늦게 끝나는 사건에서 거슬러 하위 업무만
    if events:
        end = max(events, key=lambda v: early[v])
        chain, v = [], end
        while v is not None:
            kind, no = v
            if kind == "f" and not tasks[no].is_summary and (not chain or chain[-1] != no):
                chain.append(no)
            v = best_from.get(v)
        a.critical_chain = list(reversed(chain))
