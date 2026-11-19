"""날짜 기준 일정 계산 — 선행 간격·총 여유·주경로 (v0.3).

엑셀에 적힌 날짜를 그대로 놓고 잰다 (MS Project 화면과 같은 기준). 시간 축 = 평일 번호 — 월~금, 휴일·사내 달력은 모른다 (근사).
  시작 사건 = 그 날 아침  widx(시작)
  완료 사건 = 그 날 저녁  widx(완료 + 1일)       → 완료 − 시작 = 시작~완료 평일 수
  마일스톤(기간 0) = 그 날 저녁 하나 (시작 = 완료)
선행 간격 = 후속 쪽 사건 − 선행 쪽 사건 − 지연.  FS: 완료(선행) → 시작(후속) · SS: 시작 → 시작 · FF: 완료 → 완료 · SF: 시작 → 완료.
  0 = 빡빡(구동) · 양수 = 여유 · 음수 = 날짜가 선행 조건을 어김 (겹침).
총 여유 = 늦은 시작 − 실제 시작. 늦은 값은 사업 마지막 완료에서 거꾸로 (기간·지연만 쓴다). 요약 작업 = 그 아래 하위 업무 여유의 최소.
주경로 = 총 여유 ≤ critical_float_days 인 하위 업무 · 준주경로 = 그보다 크고 ≤ near_critical_days.
날짜 없는 하위 업무와 시작 > 완료(C09)는 계산에서 빠진다 (excluded).
보기 HTML 의 JS 가 같은 규칙으로 다시 계산해 이 값과 대조한다 (template '일치 ✓').
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, timedelta

EPOCH = date(2000, 1, 3)        # 월요일
ONE = timedelta(days=1)
EPS = 1e-6


def widx(d: date) -> int:
    """d 앞에 있는 평일 수 (EPOCH 기준). 주말 날짜는 다음 월요일과 같은 값."""
    w, r = divmod((d - EPOCH).days, 7)
    return 5 * w + min(r, 5)


def wdate(n: int) -> date:
    """평일 번호 n 의 날짜."""
    w, r = divmod(int(n), 5)
    return EPOCH + timedelta(days=7 * w + r)


def start_date(v: float) -> date:
    return wdate(math.floor(v + EPS))


def finish_date(v: float) -> date:
    return wdate(math.ceil(v - EPS) - 1)


def ends(typ: str) -> tuple[int, int]:
    """관계 종류 → (선행 쪽 사건, 후속 쪽 사건) — 0 = 시작, 1 = 완료."""
    return {"FS": (1, 0), "SS": (0, 0), "FF": (1, 1), "SF": (0, 1)}.get(typ, (1, 0))


@dataclass
class Schedule:
    val: dict[int, float]                   # 사건 (2·번호 = 시작, 2·번호+1 = 완료) → 평일 값
    late: dict[int, float]
    float_: dict[int, float]                # 업무 번호 → 총 여유 (평일), 계산된 것만
    gaps: list[float | None]                # a.edges 와 같은 순서
    start: float = 0.0
    end: float = 0.0
    critical: list[int] = field(default_factory=list)   # 주경로 하위 업무 (시작 순)
    near: list[int] = field(default_factory=list)       # 준주경로 하위 업무 (시작 순)
    excluded: list[int] = field(default_factory=list)   # 날짜가 없거나 뒤집힌 하위 업무
    dropped: int = 0                        # 순환 때문에 뺀 간선
    next_float: float | None = None         # 주경로 밖 하위 업무 중 가장 작은 여유 (= 이만큼 넘게 줄이면 다른 경로가 주경로)
    crit_days: float = 0.0
    near_days: float = 5.0

    @property
    def start_date(self) -> date | None:
        return start_date(self.start) if self.val else None

    @property
    def end_date(self) -> date | None:
        return finish_date(self.end) if self.val else None

    def task_dates(self, no: int) -> tuple[date, date] | None:
        s, f = self.val.get(2 * no), self.val.get(2 * no + 1)
        if s is None or f is None:
            return None
        return (start_date(s) if f > s else finish_date(f)), finish_date(f)


def event_values(tasks) -> tuple[dict[int, float], list[int]]:
    val, excluded = {}, []
    for t in tasks.values():
        if t.children:
            continue
        if not (t.start and t.finish) or t.start > t.finish:
            excluded.append(t.no)
        elif t.is_milestone:
            val[2 * t.no] = val[2 * t.no + 1] = widx(t.finish + ONE)
        else:
            val[2 * t.no], val[2 * t.no + 1] = widx(t.start), widx(t.finish + ONE)
    for t in sorted(tasks.values(), key=lambda x: -x.level):
        if not t.children:
            continue
        ss = [val[2 * c] for c in t.children if 2 * c in val]
        if ss:
            val[2 * t.no] = min(ss)
            val[2 * t.no + 1] = max(val[2 * c + 1] for c in t.children if 2 * c + 1 in val)
        elif t.start and t.finish and t.start <= t.finish:
            val[2 * t.no], val[2 * t.no + 1] = widx(t.start), widx(t.finish + ONE)
    return val, sorted(excluded)


def date_schedule(tasks, edges, crit_days: float = 0.0, near_days: float = 5.0, val: dict[int, float] | None = None) -> Schedule:
    """val 을 주면 그 사건 값(가정 편집으로 다시 짠 것)으로, 아니면 엑셀 날짜로."""
    from .analyze import tarjan, topo

    if val is None:
        val, excluded = event_values(tasks)
    else:
        excluded = sorted(t.no for t in tasks.values() if not t.children and 2 * t.no not in val)
    arcs: dict[int, list[tuple[int, float]]] = {}

    def arc(a, b, w):
        if a in val and b in val:
            arcs.setdefault(a, []).append((b, w))

    for t in tasks.values():
        s, f = 2 * t.no, 2 * t.no + 1
        if t.children:
            for c in t.children:
                arc(s, 2 * c, 0.0)
                arc(2 * c + 1, f, 0.0)
        elif s in val:
            arc(s, f, val[f] - val[s])
    gaps = []
    for e in edges:
        a, b = ends(e.type)
        src, dst = 2 * e.pred + a, 2 * e.succ + b
        lag = e.lag_days or 0.0
        arc(src, dst, lag)
        gaps.append(round(val[dst] - val[src] - lag, 6) if src in val and dst in val else None)
    events = sorted(val)
    comps = tarjan(events, {v: [w for w, _ in arcs.get(v, ())] for v in events})
    comp_of = {v: i for i, c in enumerate(comps) for v in c}
    dag, dropped = {}, 0
    for v in events:
        for w, wt in arcs.get(v, ()):
            if comp_of[v] == comp_of[w]:
                dropped += 1
            else:
                dag.setdefault(v, []).append((w, wt))
    order = topo(events, {v: [w for w, _ in dag.get(v, ())] for v in events})
    end = max(val.values(), default=0.0)
    late = {v: end for v in events}
    for v in reversed(order):
        for w, wt in dag.get(v, ()):
            if late[w] - wt < late[v]:
                late[v] = late[w] - wt
    fl: dict[int, float] = {}
    for t in tasks.values():
        if not t.children and 2 * t.no in val:
            fl[t.no] = round(late[2 * t.no] - val[2 * t.no], 6)
    for t in sorted(tasks.values(), key=lambda x: -x.level):
        if t.children:
            sub = [fl[c] for c in t.children if c in fl]
            if sub:
                fl[t.no] = min(sub)
    sch = Schedule(val, late, fl, gaps, min(val.values(), default=0.0), end, excluded=excluded, dropped=dropped,
                   crit_days=crit_days, near_days=near_days)
    leaves = [t for t in tasks.values() if not t.children and t.no in fl]
    leaves.sort(key=lambda t: (val[2 * t.no], val[2 * t.no + 1], t.excel_row))
    sch.critical = [t.no for t in leaves if fl[t.no] <= crit_days + EPS]
    sch.near = [t.no for t in leaves if crit_days + EPS < fl[t.no] <= near_days + EPS]
    rest = [fl[t.no] for t in leaves if fl[t.no] > crit_days + EPS]
    sch.next_float = round(min(rest), 3) if rest else None
    return sch


# --- v0.4 가정 편집: 다시 짜기 ('간격 유지') ------------------------------------------------------------
# 고치지 않으면 날짜가 하나도 안 바뀐다. 단위 = 하위 업무 L · 요약 시작 제약 C · 요약 실제 시작 R · 요약 완료 F.
#   선행 간선: 출발 = 하위 업무의 시작/완료 또는 요약의 R/F, 도착 = 하위 업무(완료 쪽이면 기간을 빼 시작 요구로) 또는 요약 C.
#   요약으로 들어오는 FF·SF 는 쓰지 않는다 (근사). 위 요약의 C 는 아래 C·L 에 0 간선.
#   L·C: 원래 간격이 음수인 간선은 그만큼 느슨하게. 구동된(원래 간격 ≤ 0 인 간선이 있는) 것은 요구의 최대,
#   아니면 max(원래, 요구). 간선 없는 L 은 원래 그대로, C 는 '제약 없음'. 시작을 고친 L 은 max(고친 날, 요구).
#   새로 더한 선행(orig 에 없는 것)은 원래 간격이 없다 — 미는 쪽으로만 (구동으로 치지 않고, 느슨하게 하지 않음).
#   선행을 지워도 저절로 당겨지지는 않는다 (당기려면 시작을 고친다).

def reschedule(tasks, edges, val0: dict[int, float], dur: dict[int, float] | None = None,
               start: dict[int, float] | None = None, orig: set | None = None) -> dict[int, float]:
    """val0 = event_values(원래 날짜). dur = {하위 업무: 새 기간(평일)}, start = {하위 업무: '이 날 이후' 평일 번호}.
    orig = 원래 선행 {(선행, 후속, 종류, 지연)} — 없으면 edges 전부를 원래 것으로 본다. → 새 사건 값."""
    from .analyze import tarjan, topo

    dur, start = dur or {}, start or {}

    def leaf(no):
        return not tasks[no].children

    def d_of(no):
        return dur.get(no, val0[2 * no + 1] - val0[2 * no])

    # 단위 키: ("L", no) · ("C", no) · ("R", no) · ("F", no)
    ins: dict[tuple, list] = {}             # 도착 단위 → [(출발 단위, 출발 쪽 0/1, 가중치, 도착이 완료 쪽인가, 원래 간격)]
    outs: dict[tuple, list] = {}

    def link(src, a, dst, w, to_f, gap):
        ins.setdefault(dst, []).append((src, a, w, to_f, gap))
        outs.setdefault(src, []).append(dst)

    def orig_src(no, a):
        return val0[2 * no + a]

    for e in edges:
        p, s = e.pred, e.succ
        a, b = ends(e.type)
        if 2 * p + a not in val0 or 2 * s + b not in val0:
            continue
        lag = e.lag_days or 0.0
        gap = val0[2 * s + b] - val0[2 * p + a] - lag if orig is None or (p, s, e.type, lag) in orig else math.inf
        src = ("L", p) if leaf(p) else (("R", p) if a == 0 else ("F", p))
        if leaf(s):
            link(src, a, ("L", s), lag, b == 1, gap)
        elif b == 0:
            link(src, a, ("C", s), lag, False, gap)
    for t in tasks.values():
        if not t.children or 2 * t.no not in val0:
            continue
        for c in t.children:
            if 2 * c not in val0:
                continue
            link(("C", t.no), 0, ("L", c) if leaf(c) else ("C", c), 0.0, False, val0[2 * c] - val0[2 * t.no])
            link(("L", c) if leaf(c) else ("R", c), 0, ("R", t.no), 0.0, False, 0.0)
            link(("L", c) if leaf(c) else ("F", c), 1, ("F", t.no), 0.0, False, 0.0)
    units = sorted(set(ins) | set(outs) | {("L", t.no) for t in tasks.values() if not t.children and 2 * t.no in val0})
    comps = tarjan(units, outs)
    comp_of = {u: i for i, c in enumerate(comps) for u in c}
    order = topo(units, {u: [v for v in outs.get(u, ()) if comp_of[u] != comp_of[v]] for u in units})
    nv: dict[tuple, float | None] = {}

    def src_val(u, a):
        kind, no = u
        if kind == "L":
            return nv[u] + (d_of(no) if a == 1 else 0.0)
        return nv.get(u)

    for u in order:
        kind, no = u
        if kind in ("R", "F"):
            kids = [src_val(src, a) for src, a, _, _, _ in ins.get(u, ()) if comp_of[src] != comp_of[u]]
            kids = [k for k in kids if k is not None]
            nv[u] = (min(kids) if kind == "R" else max(kids)) if kids else val0[2 * no + (kind == "F")]
            continue
        reqs, driven = [], False
        for src, a, w, to_f, gap in ins.get(u, ()):
            if comp_of[src] == comp_of[u]:
                continue
            v = src_val(src, a)
            if v is None:
                continue
            r = v + w + min(gap, 0.0) - (d_of(no) if to_f else 0.0)
            reqs.append(r)
            driven = driven or gap <= EPS
        o = val0[2 * no]
        if kind == "L" and no in start:
            nv[u] = max([start[no]] + reqs)
        elif not reqs:
            nv[u] = o if kind == "L" else None
        elif driven:
            nv[u] = max(reqs)
        else:
            nv[u] = max([o] + reqs)
    val = {}
    for t in tasks.values():
        if not t.children and ("L", t.no) in nv:
            s = nv[("L", t.no)]
            val[2 * t.no], val[2 * t.no + 1] = s, s + d_of(t.no)
    for t in sorted(tasks.values(), key=lambda x: -x.level):
        if not t.children:
            continue
        ss = [val[2 * c] for c in t.children if 2 * c in val]
        if ss:
            val[2 * t.no] = min(ss)
            val[2 * t.no + 1] = max(val[2 * c + 1] for c in t.children if 2 * c + 1 in val)
        elif 2 * t.no in val0:
            val[2 * t.no], val[2 * t.no + 1] = val0[2 * t.no], val0[2 * t.no + 1]
    return val


# --- v0.5 저장소: 앞으로 짜기 (정본 = 도구 DB) -----------------------------------------------------------
# reschedule 은 '원래 날짜'를 기준으로 간격을 유지한다 (엑셀이 정본일 때). 정본이 DB 가 되면 기준을 값으로 굳혀 둔다:
#   하위 업무 시작 = max(시작 하한 floor, 선행·위 묶음이 요구하는 날짜) · 완료 = 시작 + 일하는 평일 수 work · 요약 = 하위 범위.
#   선행 요구 = 선행 쪽 사건 + 지연 − 느슨함 relax (들여올 때 이미 어겨 있던 만큼 — 고치지 않으면 안 움직이게).
#   하한도 요구도 없으면 저장된 날짜 그대로. 단위(L·C·R·F)·순환 간선 빼기·요약으로 들어오는 FF/SF 무시는 reschedule 과 같다.
#   요약의 C 에는 하한이 없다 (reschedule 은 요약의 원래 시작을 하한처럼 썼다) — 당겨지는 경우에만 결과가 다를 수 있다.

def forward(tasks, edges, val0: dict[int, float], floor: dict[int, float], work: dict[int, float],
            relax: list[float] | None = None) -> tuple[dict[int, float], dict[int, float]]:
    """val0 = 저장된 날짜의 사건 값 (계산에 낄 업무·기본값). floor = {하위 업무: 시작 하한}, work = {하위 업무: 일하는 평일 수},
    relax = edges 와 같은 순서의 느슨함. → (새 사건 값, {하위 업무: 선행·위 묶음 요구의 최대})."""
    from .analyze import tarjan, topo

    def leaf(no):
        return not tasks[no].children

    def d_of(no):
        return work.get(no, val0[2 * no + 1] - val0[2 * no])

    ins: dict[tuple, list] = {}
    outs: dict[tuple, list] = {}

    def link(src, a, dst, w, to_f):
        ins.setdefault(dst, []).append((src, a, w, to_f))
        outs.setdefault(src, []).append(dst)

    for i, e in enumerate(edges):
        p, s = e.pred, e.succ
        a, b = ends(e.type)
        if 2 * p + a not in val0 or 2 * s + b not in val0:
            continue
        w = (e.lag_days or 0.0) - (relax[i] if relax else 0.0)
        src = ("L", p) if leaf(p) else (("R", p) if a == 0 else ("F", p))
        if leaf(s):
            link(src, a, ("L", s), w, b == 1)
        elif b == 0:
            link(src, a, ("C", s), w, False)
    for t in tasks.values():
        if not t.children or 2 * t.no not in val0:
            continue
        for c in t.children:
            if 2 * c not in val0:
                continue
            link(("C", t.no), 0, ("L", c) if leaf(c) else ("C", c), 0.0, False)
            link(("L", c) if leaf(c) else ("R", c), 0, ("R", t.no), 0.0, False)
            link(("L", c) if leaf(c) else ("F", c), 1, ("F", t.no), 0.0, False)
    units = sorted(set(ins) | set(outs) | {("L", t.no) for t in tasks.values() if not t.children and 2 * t.no in val0})
    comps = tarjan(units, outs)
    comp_of = {u: i for i, c in enumerate(comps) for u in c}
    order = topo(units, {u: [v for v in outs.get(u, ()) if comp_of[u] != comp_of[v]] for u in units})
    nv: dict[tuple, float | None] = {}
    need: dict[int, float] = {}

    def src_val(u, a):
        if u[0] == "L":
            return nv[u] + (d_of(u[1]) if a == 1 else 0.0)
        return nv.get(u)

    for u in order:
        kind, no = u
        vals = [(src_val(src, a), w, to_f) for src, a, w, to_f in ins.get(u, ()) if comp_of[src] != comp_of[u]]
        vals = [x for x in vals if x[0] is not None]
        if kind in ("R", "F"):
            ks = [v for v, _, _ in vals]
            nv[u] = (min(ks) if kind == "R" else max(ks)) if ks else val0[2 * no + (kind == "F")]
            continue
        reqs = [v + w - (d_of(no) if to_f else 0.0) for v, w, to_f in vals]
        if kind == "C":
            nv[u] = max(reqs) if reqs else None
            continue
        if reqs:
            need[no] = max(reqs)
        fl = floor.get(no)
        cand = reqs + ([fl] if fl is not None else [])
        nv[u] = max(cand) if cand else val0[2 * no]
    val = {}
    for t in tasks.values():
        if not t.children and ("L", t.no) in nv:
            s = nv[("L", t.no)]
            val[2 * t.no], val[2 * t.no + 1] = s, s + d_of(t.no)
    for t in sorted(tasks.values(), key=lambda x: -x.level):
        if not t.children:
            continue
        ss = [val[2 * c] for c in t.children if 2 * c in val]
        if ss:
            val[2 * t.no] = min(ss)
            val[2 * t.no + 1] = max(val[2 * c + 1] for c in t.children if 2 * c + 1 in val)
        elif 2 * t.no in val0:
            val[2 * t.no], val[2 * t.no + 1] = val0[2 * t.no], val0[2 * t.no + 1]
    return val, need


# --- v0.6 일정 옮기기 (쉬프트) ---------------------------------------------------------------------------
# 범위(하위 업무 묶음)를 n 평일 통째로 옮긴다. 날짜를 직접 고치지 않고 '시작 하한' 만 바꾼다 — 그 뒤는 forward 가 짠다.
#   늦추기 (n > 0): 범위 안 하위 업무는 모두 정확히 n 만큼 늦어진다. 범위 안 선행이 밀어 주는 업무는 하한을 두지 않고(계속 선행을 따름),
#                   그렇지 않은 업무(하한에 걸려 있었거나 범위 밖 선행에 붙어 있던 것)만 새 시작을 하한으로 둔다.
#   당기기 (n < 0): 하한을 n 만큼 당긴다. 범위 밖 선행이 붙잡고 있는 업무는 그 선행이 허락하는 데까지만 (held).
#   범위 밖 후속은 여느 고침과 같이 선행을 따라 밀리거나(하한이 없으면) 당겨진다.
# 화면 미리 보기(JS shiftPlan)가 같은 규칙으로 계산한다.

def shift_floors(tasks, edges, val0: dict[int, float], floor: dict[int, float], work: dict[int, float],
                 relax: list[float] | None, scope, n: float) -> tuple[dict[int, float], dict[int, float], list[int]]:
    """scope = 옮길 하위 업무 번호들. → (새 하한 전체, 새 사건 값, 목표만큼 못 움직인 하위 업무)."""
    leaves = [no for no in scope if 2 * no in val0 and not tasks[no].children]
    target = {no: val0[2 * no] + n for no in leaves}
    new = dict(floor)
    if n > 0:
        trial = dict(floor)
        trial.update(target)
        val, need = forward(tasks, edges, val0, trial, work, relax)
        for no in leaves:
            if need.get(no, -math.inf) >= target[no] - EPS:      # 범위 안 선행이 그만큼 밀어 줌 — 하한은 그대로 n 만큼만
                if no in floor:
                    new[no] = floor[no] + n
            else:
                new[no] = target[no]
    else:
        _, need0 = forward(tasks, edges, val0, floor, work, relax)
        for no in leaves:
            if no in floor:
                new[no] = floor[no] + n
            elif no not in need0:                                # 하한도 요구도 없어 저장된 날짜에 머물던 업무
                new[no] = target[no]
        val, _ = forward(tasks, edges, val0, new, work, relax)
    held = sorted(no for no in leaves if (target[no] - val[2 * no] if n > 0 else val[2 * no] - target[no]) > EPS)   # 목표만큼 못 간 업무
    return new, val, held
