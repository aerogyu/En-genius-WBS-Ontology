"""가정 편집 적용 — 보기 HTML 이 내려준 변경 목록(.json)을 원본 엑셀에 반영해 **새 판 엑셀**을 쓴다 (v0.4).

원본은 건드리지 않는다: `<원본>_수정_v01.xlsx` (있으면 v02 …) + 같은 이름 `.changes.json` (적용한 변경 목록 사본 = 판 이력).
바뀐 칸만 고친다 — openpyxl 로 원본을 열어 셀 값만 바꾸므로 서식·다른 시트·열 순서가 그대로다.
날짜는 '간격 유지' 규칙(schedule.reschedule)으로 다시 짜고, 보기 HTML 이 계산한 결과(변경 목록의 result)와 대조한다.
인명 칸은 읽지도 쓰지도 않는다.
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from .analyze import Edge, analyze
from .config import nfkc
from .parse import WEEKDAYS_KO, find_header, parse_deps, read_wbs
from .schedule import date_schedule, event_values, finish_date, reschedule, start_date, widx

FORMAT = "engenius-wbs-changes/1"
FIELDS = ("duration", "start", "preds", "name", "man_hours", "staff")
DATE_TXT = re.compile(r"^(\s*)(\d{2,4})([-./])(\d{1,2})([-./])(\d{1,2})(\s*)(\(\s*[월화수목금토일]\s*\))?(.*)$")
DUR_TXT = re.compile(r"^(\s*<?\s*)([0-9][0-9,]*(?:\.[0-9]+)?)(\s*)([^\d\s?<>]*)(.*)$")
DAY_UNITS = {"", "일", "d", "day", "days", "경과일"}


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def next_name(src: Path, out_dir: Path) -> Path:
    stem = re.sub(r"_수정_v\d+$", "", src.stem)
    for k in range(1, 1000):
        p = out_dir / f"{stem}_수정_v{k:02d}.xlsx"
        if not p.exists():
            return p
    raise RuntimeError("판 번호가 999 를 넘었다")


def date_like(orig, d: date):
    """원래 칸 모양 그대로 날짜를 쓴다 — 글자 '27-01-01 (금)' 이면 같은 모양, 날짜 셀이면 날짜."""
    if isinstance(orig, datetime):
        return datetime(d.year, d.month, d.day, orig.hour, orig.minute)
    if isinstance(orig, date):
        return d
    m = DATE_TXT.match(nfkc(orig)) if isinstance(orig, str) else None
    if not m:
        return f"{d:%y-%m-%d} ({WEEKDAYS_KO[d.weekday()]})"
    lead, y, s1, mo, s2, dd, sp, wd, tail = m.groups()
    ys = f"{d.year % 100:02d}" if len(y) == 2 else str(d.year)
    ms = f"{d.month:02d}" if len(mo) == 2 else str(d.month)
    ds = f"{d.day:02d}" if len(dd) == 2 else str(d.day)
    return f"{lead}{ys}{s1}{ms}{s2}{ds}" + (f"{sp}({WEEKDAYS_KO[d.weekday()]})" if wd else "") + tail


def duration_like(orig, days: float):
    n = int(days) if float(days).is_integer() else round(days, 2)
    if isinstance(orig, (int, float)) and not isinstance(orig, bool):
        return n
    m = DUR_TXT.match(nfkc(orig)) if isinstance(orig, str) else None
    if m and m.group(4).strip().lower() in DAY_UNITS:
        return f"{m.group(1)}{n}{m.group(3)}{m.group(4)}{m.group(5)}"
    return f"{n} 일"


def succ_text(no: int, tasks) -> str:
    toks = []
    for t in sorted(tasks.values(), key=lambda x: x.excel_row):
        for d in t.deps:
            if d.pred == no:
                toks.append(f"{t.no}{'' if d.type == 'FS' and not d.lag_text else d.type}{d.lag_text}")
    return ",".join(toks)


@dataclass
class ApplyResult:
    out: Path
    cells: int = 0
    moved: list[int] = field(default_factory=list)
    kept_summaries: list[int] = field(default_factory=list)    # 엑셀 요약 날짜가 원래 하위 범위와 달라(C11) 그대로 둔 것
    end_before: date | None = None
    end_after: date | None = None
    mismatches: list[str] = field(default_factory=list)       # 보기 HTML 결과와 다른 것
    checked: bool = False                                     # 변경 목록에 보기 HTML 결과가 있어 대조했나


def load_changes(path: Path) -> dict:
    ch = json.loads(Path(path).read_text(encoding="utf-8"))
    if ch.get("format") != FORMAT:
        raise ValueError(f"변경 목록 형식이 아니다 (format={ch.get('format')!r})")
    for c in ch.get("changes", []):
        if c.get("field") not in FIELDS or not isinstance(c.get("no"), int):
            raise ValueError(f"읽을 수 없는 변경: {c}")
    return ch


def apply_changes(src: Path, changes: Path, cfg: dict, out_dir: Path | None = None, force: bool = False) -> ApplyResult:
    src, changes = Path(src), Path(changes)
    ch = load_changes(changes)
    sha = sha256_of(src)
    if ch.get("sha256") and ch["sha256"] != sha and not force:
        raise ValueError("이 변경 목록은 다른 판의 엑셀에서 만든 것이다 (sha256 다름) — 같은 원본을 고르거나 --force")
    w = read_wbs(src, cfg, ch.get("sheet") or None)
    tasks = w.tasks
    a = analyze(w, cfg)
    val0, _ = event_values(tasks)
    units = {nfkc(k).lower(): float(v) for k, v in cfg["unit_days"].items()}
    dur, start, edited = {}, {}, {}
    mh_old = {no: t.man_hours for no, t in tasks.items()}
    for c in ch["changes"]:
        no, f, to = c["no"], c["field"], c.get("to")
        if no not in tasks:
            raise ValueError(f"없는 번호 {no}")
        t = tasks[no]
        edited.setdefault(no, set()).add(f)
        if f == "duration":
            if t.children:
                raise ValueError(f"{no}: 요약 작업의 기간은 고치지 않는다")
            dur[no] = float(to)
        elif f == "start":
            if to:
                start[no] = widx(date.fromisoformat(to))
        elif f == "preds":
            deps, bad = parse_deps(to or None, units)
            if bad:
                raise ValueError(f"{no}: 선행 표기를 읽지 못함 {bad}")
            missing = [d.pred for d in deps if d.pred not in tasks or d.pred == no]
            if missing:
                raise ValueError(f"{no}: 없는 번호·자기 자신을 선행으로 {missing}")
            t.deps = deps
        elif f == "name":
            t.name = str(to or "").strip()
        elif f == "man_hours":
            t.man_hours = None if to in (None, "") else float(to)
        elif f == "staff":
            t.staff = None if to in (None, "") else float(to)
    edges = [Edge(d.pred, t.no, d.type, d.lag_days, d.lag_text) for t in tasks.values() for d in t.deps
             if d.pred in tasks and d.pred != t.no]
    orig = {(e.pred, e.succ, e.type, e.lag_days or 0.0) for e in a.edges}
    val = reschedule(tasks, edges, val0, dur, start, orig)
    sc0, sc1 = a.schedule, date_schedule(tasks, edges, 0, 5, val=val)
    res = ApplyResult(next_name(src, Path(out_dir) if out_dir else src.parent))
    res.end_before, res.end_after = sc0.end_date, sc1.end_date

    def dates(no, v):
        s, f = v[2 * no], v[2 * no + 1]
        return (finish_date(f) if f - s < 1e-9 else start_date(s)), finish_date(f)

    new_dates = {}
    for t in tasks.values():
        if 2 * t.no not in val or 2 * t.no not in val0:
            continue
        if abs(val[2 * t.no] - val0[2 * t.no]) < 1e-9 and abs(val[2 * t.no + 1] - val0[2 * t.no + 1]) < 1e-9:
            continue
        if t.children and not (t.start or t.finish):
            continue                                               # 엑셀에 요약 날짜가 비어 있으면 비워 둠
        if t.children and (t.start, t.finish) != a.child_span.get(t.no):
            res.kept_summaries.append(t.no)                        # C11: 엑셀 요약 날짜를 손으로 잡은 것 — 그대로
            continue
        new_dates[t.no] = dates(t.no, val)
    res.moved = sorted(new_dates)
    # 보기 HTML 결과와 대조
    result = ch.get("result") or {}
    res.checked = "moved" in result
    if "moved" in result:
        want = {int(k): tuple(v) for k, v in result["moved"].items()}
        got = {no: (s.isoformat(), f.isoformat()) for no, (s, f) in new_dates.items()}
        for no in sorted(set(want) | set(got)):
            if want.get(no) != got.get(no):
                res.mismatches.append(f"{no}: 보기 {'~'.join(want[no]) if no in want else '안 바뀜'} · apply {'~'.join(got[no]) if no in got else '안 바뀜'}")
    if result.get("end") and res.end_after and result["end"] != res.end_after.isoformat():
        res.mismatches.append(f"사업 완료: 보기 {result['end']} · apply {res.end_after}")

    # 쓰기 — 원본 사본을 열어 바뀐 칸만
    out = res.out
    shutil.copyfile(src, out)
    wb = load_workbook(out)
    ws = wb[w.sheet]
    _, cols, _, _ = find_header(ws, cfg)

    def put(t, key, value):
        c = cols.get(key)
        if not c:
            return
        cell = ws.cell(row=t.excel_row, column=c)
        if cell.value != value:
            cell.value = value
            res.cells += 1

    for no, fs in edited.items():
        t = tasks[no]
        if "duration" in fs:
            put(t, "duration", duration_like(ws.cell(row=t.excel_row, column=cols["duration"]).value if "duration" in cols else None, dur[no]))
        if "preds" in fs:
            put(t, "preds", ",".join(f"{d.pred}{'' if d.type == 'FS' and not d.lag_text else d.type}{d.lag_text}" for d in t.deps) or None)
        if "name" in fs:
            old = ws.cell(row=t.excel_row, column=cols["name"]).value
            lead = re.match(r"^[ \t　 ]*", str(old or "")).group(0)
            put(t, "name", lead + t.name)
        if "man_hours" in fs:
            put(t, "man_hours", t.man_hours)
        if "staff" in fs:
            put(t, "staff", t.staff)
    # 요약 Man Hours: 원래 '적힌 값 = 아래 합' 이던 위 요약만 차이만큼 고침 (C12 인 것은 그대로)
    below = a.workload.mh_below
    delta: dict[int, float] = {}
    for no, fs in edited.items():
        if "man_hours" in fs:
            dv = (tasks[no].man_hours or 0) - (mh_old[no] or 0)
            p = tasks[no].parent
            while p is not None and dv:
                if mh_old[p] and below.get(p) is not None and abs(mh_old[p] - below[p]) <= max(0.5, abs(mh_old[p]) * 1e-6):
                    delta[p] = delta.get(p, 0.0) + dv
                p = tasks[p].parent
    for p, dv in delta.items():
        if p not in edited or "man_hours" not in edited[p]:
            tasks[p].man_hours = round(mh_old[p] + dv, 6)
            put(tasks[p], "man_hours", tasks[p].man_hours)
    for no, (s, f) in new_dates.items():
        t = tasks[no]
        for key, d in (("start", s), ("finish", f)):
            if key in cols:
                put(t, key, date_like(ws.cell(row=t.excel_row, column=cols[key]).value, d))
    pred_edits = [no for no, fs in edited.items() if "preds" in fs]
    if "succs" in cols and pred_edits:                               # 후속 열(H): 후속 목록이 바뀐 선행 행만 다시 씀
        orig = read_wbs(src, cfg, w.sheet).tasks
        touched = {d.pred for no in pred_edits for d in orig[no].deps + tasks[no].deps}
        for p in sorted(touched):
            if p in tasks and set(succ_keys(p, orig)) != set(succ_keys(p, tasks)):
                put(tasks[p], "succs", succ_text(p, tasks) or None)
    wb.save(out)
    wb.close()
    log = dict(ch)
    log["applied"] = {"output": out.name, "at": datetime.now().replace(microsecond=0).isoformat(), "cells": res.cells,
                      "moved": len(res.moved), "end_before": str(res.end_before), "end_after": str(res.end_after),
                      "kept_summaries": res.kept_summaries, "mismatches": res.mismatches}
    out.with_suffix(".changes.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    return res


def succ_keys(no: int, tasks) -> list[tuple]:
    return [(t.no, d.type, d.lag_text) for t in tasks.values() for d in t.deps if d.pred == no]
