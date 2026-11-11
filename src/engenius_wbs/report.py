"""보고서 xlsx — 요약 · 검사 · 단계 요약 · 분야 요약 · 월별 공수 · 업무(계산 열 포함) · 분야 행렬 · 분야 간 연결 · 가장 긴 사슬."""
from __future__ import annotations

from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .analyze import Analysis
from .parse import ISSUES, LEVELS_MSG, Wbs

BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")


def _sheet(wb, title, header, widths):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, wdt in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = wdt
    ws.freeze_panes = "A2"
    return ws


def summary_sheet(wb, title: str, w: Wbs, a: Analysis, level: int):
    """레벨 level 요소마다: 엑셀 날짜 · 하위 범위 · 차이 · 하위 업무 수 · MH · 마일스톤 · 밖과 오가는 선행 — '단계 기간이 왜 이러냐' 를 찾는 표."""
    tasks = w.tasks
    head = ["번호", "WBS", "구성품", "최상위 업무", "개발 단계"] + (["기술 분야"] if level >= 4 else []) + \
        ["엑셀 시작", "엑셀 완료", "하위 시작", "하위 완료", "시작 차이(일)", "완료 차이(일)", "하위 업무", "Man Hours(쓴 것)",
         "마일스톤", "밖에서 들어옴", "밖으로 나감"]
    s = _sheet(wb, title, head, [7, 16, 12, 12, 14] + ([12] if level >= 4 else []) + [11, 11, 11, 11, 8, 8, 8, 12, 8, 9, 9])

    def up(no: int, lv: int):
        t = tasks[no]
        while t.level > lv and t.parent is not None:
            t = tasks[t.parent]
        return t.no if t.level == lv else None

    def name_at(no: int, lv: int) -> str:
        x = up(no, lv)
        return tasks[x].name if x is not None else ""

    leaves, mh, ms = Counter(), Counter(), Counter()
    cset = set(a.workload.contributors)
    for t in tasks.values():
        x = up(t.no, level)
        if x is None:
            continue
        if not t.children:
            leaves[x] += 1
            ms[x] += t.is_milestone
        if t.no in cset:
            mh[x] += t.man_hours
    cin, cout = Counter(), Counter()
    for e in a.edges:
        p, q = up(e.pred, level), up(e.succ, level)
        if p != q:
            if p is not None:
                cout[p] += 1
            if q is not None:
                cin[q] += 1
    for t in sorted(tasks.values(), key=lambda x: x.excel_row):
        if t.level != level:
            continue
        sp = a.child_span.get(t.no)
        d0 = (t.start - sp[0]).days if sp and t.start else None
        d1 = (t.finish - sp[1]).days if sp and t.finish else None
        s.append([t.no, t.wbs, name_at(t.no, 1), name_at(t.no, 2), name_at(t.no, 3)] + ([t.name] if level >= 4 else []) +
                 [t.start, t.finish, sp[0] if sp else None, sp[1] if sp else None, d0 or None, d1 or None, leaves[t.no],
                  round(mh[t.no], 1) if mh[t.no] else None, ms[t.no] or None, cin[t.no] or None, cout[t.no] or None])
    s.auto_filter.ref = s.dimensions


def month_sheet(wb, w: Wbs, a: Analysis):
    """행 = 달, 열 = 구성품, 값 = Man Hours (평일 균등 배분). 맨 아래 합계·날짜 없음."""
    wl, tasks = a.workload, w.tasks
    keys = list(wl.modules) + sorted({m for _, m in wl.cells} - set(wl.modules))
    s = _sheet(wb, "월별 공수", ["달"] + [tasks[m].name for m in keys] + ["합계"], [9] + [11] * len(keys) + [12])
    tot = Counter()
    for ym in wl.months:
        vals = [wl.cells.get((ym, m), 0.0) for m in keys]
        for m, v in zip(keys, vals):
            tot[m] += v
        s.append([ym] + [round(v, 1) or None for v in vals] + [round(sum(vals), 1)])
    s.append(["합계"] + [round(tot[m], 1) for m in keys] + [round(wl.total_mh, 1)])
    s.append(["날짜 없음"] + [None] * len(keys) + [round(wl.undated_mh, 1)])
    for c in s[s.max_row - 1]:
        c.font = BOLD
    for row in s.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = "#,##0"


def all_issues(w: Wbs, a: Analysis):
    return sorted(w.issues + a.issues, key=lambda i: (LEVELS_MSG[i.level], i.code, i.no or 0))


def write_report(w: Wbs, a: Analysis, cfg: dict, path: Path) -> Path:
    tasks = w.tasks
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    issues = all_issues(w, a)
    wl = a.workload
    lv = Counter(t.level for t in tasks.values())
    rows = [
        ("원본 파일", Path(w.source).name), ("시트", w.sheet),
        ("업무 수", len(tasks)),
        *[(f"  Lv{k} {cfg['levels'].get(k, {}).get('label', '하위')}", lv[k]) for k in sorted(lv)],
        ("선행 관계 (두 끝이 있는 것)", len(a.edges)),
        ("순환 묶음", len(a.cycles)),
        ("가장 긴 사슬 (달력 없는 근사, 근무일)", round(a.project_days, 1)),
        ("가장 긴 사슬의 하위 업무 수", len(a.critical_chain)),
        ("CPM 에서 뺀 순환 간선", a.cpm_dropped_edges),
        ("마일스톤 (기간 0 인 하위 업무)", sum(1 for t in tasks.values() if t.is_milestone)),
        ("월별 공수: 배분 방식", "Man Hours 를 시작~완료 평일에 고르게 (휴일·사내 달력 모름 — 근사, 총합은 같음)"),
        ("월별 공수: 쓴 행", f"{len(wl.contributors):,} (아래에 적힌 MH 가 있으면 아래 것만 — "
                          + " · ".join(f"Lv{k} {v:,}" for k, v in wl.by_level.items()) + ")"),
        ("월별 공수: 달에 나눈 MH", round(wl.total_mh, 1)),
        ("월별 공수: 날짜 없어 못 나눈 MH", f"{wl.undated_mh:,.1f} ({len(wl.undated)}행)"),
        ("월별 공수: MH 가 없는 하위 업무", f"{wl.leaves_without_mh:,} / {wl.leaves:,} (마일스톤 포함)"),
        ("읽은 칸", ", ".join(f"{k}={v}" for k, v in w.columns.items())),
        ("읽지 않은 칸 (인명·설정 밖)", ", ".join(w.skipped_columns)),
    ]
    for k, v in rows:
        ws.append([k, v])
    ws.append([])
    ws.append(["문제 코드", "수준", "이름", "건수"])
    for c in ws[ws.max_row]:
        c.font = BOLD
    cnt = Counter(i.code for i in issues)
    for code in sorted(cnt, key=lambda c: (LEVELS_MSG[ISSUES[c][0]], c)):
        ws.append([code, ISSUES[code][0], ISSUES[code][1], cnt[code]])
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 50

    s = _sheet(wb, "검사", ["수준", "코드", "이름", "번호", "엑셀 행", "WBS", "작업 이름", "내용"], [7, 6, 36, 8, 8, 16, 40, 60])
    for i in issues:
        name = tasks[i.no].name if i.no in tasks else ""
        s.append([i.level, i.code, i.title, i.no, i.excel_row, i.wbs, name, i.detail])

    summary_sheet(wb, "단계 요약", w, a, cfg.get("phase_level", 3))
    summary_sheet(wb, "분야 요약", w, a, cfg.get("discipline_level", 4))
    month_sheet(wb, w, a)

    critical = set(a.critical_chain)
    s = _sheet(wb, "업무", ["번호", "WBS", "레벨", "작업 이름", "구성품", "개발 단계", "기술 분야", "기간(일)", "시작", "완료",
                            "선행", "선행 수", "후속 수", "앞쪽 전체", "뒤쪽 전체", "CPM 이른 시작(근무일)", "여유(근무일)",
                            "가장 긴 사슬", "투입인력", "Man Hours", "산출물"],
               [7, 16, 5, 40, 12, 12, 12, 8, 11, 11, 20, 7, 7, 9, 9, 10, 9, 8, 8, 9, 30])
    for t in sorted(tasks.values(), key=lambda x: x.excel_row):
        preds = ", ".join(f"{d.pred}{'' if d.type == 'FS' else d.type}{d.lag_text}" for d in t.deps)
        s.append([t.no, t.wbs, t.level, ("  " * (t.level - 1)) + t.name, t.module, t.phase, t.discipline, t.duration_days,
                  t.start, t.finish, preds, len(a.preds_of.get(t.no, [])), len(a.succs_of.get(t.no, [])),
                  a.up_count.get(t.no, 0), a.down_count.get(t.no, 0), round(a.es.get(t.no, 0), 1),
                  round(a.total_float.get(t.no, 0), 1), "●" if t.no in critical else "", t.staff, t.man_hours,
                  "; ".join(t.deliverables)])
    s.auto_filter.ref = s.dimensions

    discs = sorted({d for pair in a.discipline_flows for d in pair})
    s = _sheet(wb, "분야 행렬", ["선행 분야 \\ 후속 분야"] + discs, [22] + [10] * len(discs))
    for pd in discs:
        s.append([pd] + [a.discipline_flows.get((pd, sd)) or None for sd in discs])
    for r in s.iter_rows(min_row=2, min_col=1, max_col=1):
        r[0].font = BOLD

    s = _sheet(wb, "분야 간 연결", ["선행 분야", "후속 분야", "선행 번호", "선행 WBS", "선행 작업", "관계", "후속 번호", "후속 WBS",
                                 "후속 작업", "구성품(후속)", "개발 단계(후속)"], [12, 12, 8, 16, 36, 8, 8, 16, 36, 12, 12])
    for e in sorted(a.edges, key=lambda e: (tasks[e.pred].discipline, tasks[e.succ].discipline, e.pred, e.succ)):
        p, q = tasks[e.pred], tasks[e.succ]
        if p.discipline and q.discipline and p.discipline != q.discipline:
            s.append([p.discipline, q.discipline, p.no, p.wbs, p.name, e.type + e.lag_text, q.no, q.wbs, q.name, q.module, q.phase])
    s.auto_filter.ref = s.dimensions

    s = _sheet(wb, "가장 긴 사슬", ["순서", "번호", "WBS", "작업 이름", "기술 분야", "기간(일)", "CPM 이른 시작", "엑셀 시작", "엑셀 완료"],
               [6, 8, 16, 40, 12, 8, 10, 11, 11])
    for k, no in enumerate(a.critical_chain, start=1):
        t = tasks[no]
        s.append([k, t.no, t.wbs, t.name, t.discipline, t.duration_days, round(a.es.get(no, 0), 1), t.start, t.finish])

    for sh in wb.worksheets:
        for row in sh.iter_rows(min_row=2):
            for c in row:
                if hasattr(c.value, "year"):
                    c.number_format = "yyyy-mm-dd"
    wb.save(path)
    return path
