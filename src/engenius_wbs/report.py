"""보고서 xlsx — 요약 · 검사 · 업무(계산 열 포함) · 분야 행렬 · 분야 간 연결 · 가장 긴 사슬."""
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


def all_issues(w: Wbs, a: Analysis):
    return sorted(w.issues + a.issues, key=lambda i: (LEVELS_MSG[i.level], i.code, i.no or 0))


def write_report(w: Wbs, a: Analysis, cfg: dict, path: Path) -> Path:
    tasks = w.tasks
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    issues = all_issues(w, a)
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
