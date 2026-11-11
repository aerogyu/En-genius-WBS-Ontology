"""월별 공수 — Man Hours 를 업무의 시작~완료 평일에 고르게 나눠 달마다 더한다 (휴일·사내 달력 모름 — 근사, 총합은 같다).

어느 행의 MH 를 쓰나 ('가장 아래에 적힌 값'): 아래(하위)에 MH 가 적힌 업무가 하나라도 있으면 아래 것만, 없으면 자기 것.
  → MS Project 가 요약 행에 적는 하위 합을 두 번 세지 않고, MH 가 Lv6 에 있든 Lv5 에만 있든 한 규칙으로 받는다.
  MH 0·빈 칸은 '적히지 않음'. 날짜가 없거나 거꾸로인 업무의 MH 는 '날짜 없음' 으로 따로 센다 (어느 달인지 모름).
한 업무: 시작~완료(양끝 포함) 의 월~금 날 수로 나눈 값을 하루치로 본다. 평일이 하루도 없으면(주말만) 달력 날 수로 나눈다.
보기 HTML 이 같은 규칙을 JS 로 다시 계산한다 (거르기를 따라가려고) — 이 모듈의 월 합계를 검산값으로 넣는다.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import timedelta

from .parse import Task, Wbs


@dataclass
class Workload:
    contributors: list[int]                       # MH 를 쓴 업무 번호 (엑셀 순)
    cells: dict[tuple[str, int], float]           # ('YYYY-MM', 구성품 번호) → MH
    months: list[str]                             # 첫 달 ~ 끝 달 (빈 달 포함)
    modules: list[int]                            # 구성품 번호 (엑셀 순)
    undated: list[int] = field(default_factory=list)
    undated_mh: float = 0.0
    total_mh: float = 0.0                         # 달에 나눈 MH 합
    mh_below: dict[int, float] = field(default_factory=dict)   # 요약 작업 → 아래에서 쓴 MH 합 (아래에 적힌 것이 있을 때만)
    leaves: int = 0
    leaves_without_mh: int = 0                    # 자기도 위 묶음도 MH 를 쓰지 않은 하위 업무
    by_level: dict[int, int] = field(default_factory=dict)     # MH 를 쓴 업무의 레벨별 수

    def month_totals(self) -> dict[str, float]:
        out = {m: 0.0 for m in self.months}
        for (m, _), v in self.cells.items():
            out[m] += v
        return out


def has_mh(t: Task) -> bool:
    return bool(t.man_hours)


def spread(t: Task) -> dict[str, float]:
    """한 업무의 MH → {'YYYY-MM': MH}."""
    days = [t.start + timedelta(days=i) for i in range((t.finish - t.start).days + 1)]
    work = [d for d in days if d.weekday() < 5] or days
    per = t.man_hours / len(work)
    out: dict[str, float] = defaultdict(float)
    for d in work:
        out[f"{d.year:04d}-{d.month:02d}"] += per
    return out


def month_range(first: str, last: str) -> list[str]:
    y, m = int(first[:4]), int(first[5:])
    out = []
    while f"{y:04d}-{m:02d}" <= last:
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def workload(w: Wbs, module_level: int = 1) -> Workload:
    tasks = w.tasks
    order = sorted(tasks.values(), key=lambda x: x.excel_row)
    deep_first = sorted(tasks.values(), key=lambda x: -x.level)
    below_any: dict[int, bool] = {}
    for t in deep_first:
        below_any[t.no] = any(below_any[c] or has_mh(tasks[c]) for c in t.children)
    contrib = [t.no for t in order if has_mh(t) and not below_any[t.no]]
    cset = set(contrib)

    def top(no: int) -> int:
        t = tasks[no]
        while t.level > module_level and t.parent is not None:
            t = tasks[t.parent]
        return t.no

    cells: dict[tuple[str, int], float] = defaultdict(float)
    undated, undated_mh, by_level = [], 0.0, defaultdict(int)
    for no in contrib:
        t = tasks[no]
        by_level[t.level] += 1
        if not (t.start and t.finish) or t.finish < t.start:
            undated.append(no)
            undated_mh += t.man_hours
            continue
        m = top(no)
        for ym, v in spread(t).items():
            cells[(ym, m)] += v
    # 요약 작업 → 아래에서 쓴 MH 합
    mh_below: dict[int, float] = {}
    sums: dict[int, float] = {}
    for t in deep_first:
        s = sum(sums[c] for c in t.children)
        if t.children and below_any[t.no]:
            mh_below[t.no] = s
        sums[t.no] = s + (t.man_hours if t.no in cset else 0.0)
    # 자기도 위 묶음도 MH 를 쓰지 않은 하위 업무
    leaves = [t for t in order if not t.children]
    covered = 0
    for t in leaves:
        x = t
        while x is not None:
            if x.no in cset:
                covered += 1
                break
            x = tasks[x.parent] if x.parent is not None else None
    ms = sorted({k[0] for k in cells})
    months = month_range(ms[0], ms[-1]) if ms else []
    modules = [t.no for t in order if t.level == module_level]
    return Workload(contrib, dict(cells), months, modules, undated, undated_mh, sum(cells.values()), mh_below,
                    len(leaves), len(leaves) - covered, dict(sorted(by_level.items())))
