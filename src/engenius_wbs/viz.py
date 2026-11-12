"""WBS 보기 HTML 한 장 (v3, 피벗 방식 + v0.2 마일스톤·월별 공수 탭) — 외부 JS·CDN 없음 (폐쇄망에서 파일을 그대로 연다).

Lv6 업무마다 (구성품, 최상위 업무, 개발 단계, 기술 분야, 상위 기술 업무) 다섯 꼬리표가 있다고 보고, 엑셀 피벗처럼 다룬다:
  거르기(레벨마다 이름 여러 개 체크 + 트리 체크로 손으로 고른 묶음) × 묶기(상자 = 어느 레벨) × 줄 나누기(없음 | Lv1~4 이름 | 고른 항목별) × 배치(시간표 | 흐름도).
  빠른 보기 4개(단계 하나·분야별 줄 / 분야 하나·모든 단계 / 한 레벨 전체·줄 없음 / 트리에서 고른 것), '지금 보기' 한 줄 문장, 보기 주소(#) 복사.
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
    sch = a.schedule
    tasks = []
    for t in sorted(w.tasks.values(), key=lambda x: x.excel_row):
        tasks.append([t.no, t.wbs, t.level, t.name, t.duration_days, t.start.isoformat() if t.start else "",
                      t.finish.isoformat() if t.finish else "", a.up_count.get(t.no, 0), a.down_count.get(t.no, 0),
                      round(a.total_float.get(t.no, 0), 1), 1 if t.no in crit else 0, t.parent, "; ".join(t.deliverables),
                      t.man_hours, t.staff, sch.float_.get(t.no)])
    edges = [[e.pred, e.succ, ("" if e.type == "FS" else e.type) + e.lag_text, e.type, e.lag_days or 0, g]
             for e, g in zip(a.edges, sch.gaps)]
    issues = [[i.code, i.level, i.title, i.no, i.wbs, i.detail] for i in all_issues(w, a)]
    wl = a.workload
    mh = {"months": {m: round(v, 1) for m, v in wl.month_totals().items()}, "total": round(wl.total_mh, 1),
          "undated_mh": round(wl.undated_mh, 1), "undated": len(wl.undated), "rows": len(wl.contributors)}
    return {"source": Path(w.source).name, "tasks": tasks, "edges": edges, "issues": issues,
            "project_days": round(a.project_days, 1), "cycles": len(a.cycles), "mh": mh,
            "cp": {"crit_days": sch.crit_days, "near_days": sch.near_days, "critical": len(sch.critical), "near": len(sch.near),
                   "start": sch.start_date.isoformat() if sch.start_date else "", "end": sch.end_date.isoformat() if sch.end_date else "",
                   "next_float": sch.next_float, "excluded": len(sch.excluded)}}


def write_html(w: Wbs, a: Analysis, path: Path, title: str = "WBS 선후 관계") -> Path:
    data = json.dumps(payload(w, a), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = TEMPLATE.replace("__TITLE__", html.escape(title)).replace("__DATA__", data)
    path.write_text(page, encoding="utf-8", newline="\n")
    return path


TEMPLATE = (Path(__file__).with_name("template_v3.html")).read_text(encoding="utf-8")
