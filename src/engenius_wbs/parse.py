"""WBS 엑셀 읽기 — 헤더 이름으로 칸을 찾고, 행마다 업무(Task) 하나를 만든다.

읽는 것: 번호(A)·WBS·작업 이름(앞 띄어쓰기 = 들여쓰기)·기간·시작·완료(글자 '27-01-01 (금)')·선행·후속·투입인력·Man Hours·산출물.
읽지 않는 것: 작성자 등 인명 칸 (config person_columns), 설정에 없는 칸.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from .config import nfkc, norm_header

WEEKDAYS_KO = "월화수목금토일"
LEVELS_MSG = {"오류": 0, "경고": 1, "알림": 2}

# 문제 코드 → (수준, 이름). P = 읽기, C = 그래프 검사
ISSUES = {
    "P01": ("오류", "헤더를 찾지 못함"),
    "P02": ("오류", "번호(A열)가 없거나 숫자가 아님 — 행을 읽지 않음"),
    "P03": ("오류", "번호 중복 — 뒤 행을 읽지 않음"),
    "P04": ("오류", "WBS 코드 없음·형식 이상"),
    "P05": ("오류", "WBS 코드 중복"),
    "P06": ("경고", "상위 WBS 코드가 없음"),
    "P07": ("알림", "들여쓰기와 WBS 레벨이 어긋남"),
    "P08": ("경고", "날짜를 읽지 못함"),
    "P09": ("알림", "날짜와 요일이 어긋남"),
    "P10": ("경고", "기간을 읽지 못함"),
    "P11": ("오류", "선행 작업 표기를 읽지 못함"),
    "P12": ("경고", "숫자 칸을 읽지 못함"),
    "P13": ("알림", "WBS 코드가 숫자 서식 — 1.10 과 1.1 을 구별할 수 없음"),
    "C01": ("오류", "없는 번호를 선행으로 적음"),
    "C02": ("오류", "자기 자신을 선행으로 적음"),
    "C03": ("오류", "선행 관계가 순환함"),
    "C04": ("알림", "상위 묶음(요약 작업)이 선행 관계에 걸림"),
    "C05": ("경고", "선행 조건과 날짜가 어긋남"),
    "C06": ("알림", "기간과 시작·완료 날짜가 맞지 않음 (휴일 모름 — 근사)"),
    "C07": ("알림", "후속 작업(H열)이 선행(G열)을 뒤집은 것과 다름"),
    "C08": ("알림", "선행·후속이 하나도 없는 하위 업무"),
    "C09": ("경고", "시작이 완료보다 늦음"),
    "C10": ("알림", "상위 묶음과 그 하위 업무 사이의 선행 관계"),
    "C11": ("알림", "요약 작업 날짜가 하위 업무 범위와 다름"),
    "C12": ("알림", "요약 작업 Man Hours 가 하위 합과 다름"),
}


@dataclass
class Issue:
    code: str
    no: int | None = None          # 업무 번호 (A열)
    excel_row: int | None = None
    wbs: str = ""
    detail: str = ""

    @property
    def level(self) -> str:
        return ISSUES[self.code][0]

    @property
    def title(self) -> str:
        return ISSUES[self.code][1]


@dataclass
class Dep:
    pred: int
    type: str = "FS"
    lag_days: float | None = 0.0   # None = 단위를 모름 (비율 등)
    lag_text: str = ""


@dataclass
class Task:
    no: int
    excel_row: int
    wbs: str
    level: int
    name: str
    indent: int = 0
    duration_days: float | None = None
    duration_text: str = ""
    estimated: bool = False
    start: date | None = None
    finish: date | None = None
    deps: list[Dep] = field(default_factory=list)
    succs_declared: list[int] | None = None
    staff: float | None = None
    man_hours: float | None = None
    deliverables: list[str] = field(default_factory=list)
    parent: int | None = None
    children: list[int] = field(default_factory=list)
    module: str = ""
    phase: str = ""
    discipline: str = ""

    @property
    def is_summary(self) -> bool:
        return bool(self.children)

    @property
    def is_milestone(self) -> bool:
        """마일스톤 = 하위 업무이고 기간 0 (MS Project 의 '0 일')."""
        return not self.children and self.duration_days == 0


@dataclass
class Wbs:
    source: str
    sheet: str
    tasks: dict[int, Task]
    issues: list[Issue]
    columns: dict[str, str]        # 칸 key → 찾은 헤더 문구
    skipped_columns: list[str]     # 읽지 않은 헤더 (인명 칸 포함)
    uids: dict[int, int] = field(default_factory=dict)   # 번호 → 고유 ID (도구 DB 에서 읽었을 때만, store.py)


# --- 값 읽기 -------------------------------------------------------------------------------------

DATE_RE = re.compile(r"(\d{2,4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?\s*(?:\(\s*([월화수목금토일])\s*\))?")
DUR_RE = re.compile(r"^<?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*([^\d\s?<>]*)\s*(\?)?\s*>?$")
TYPE_KO = {"완시": "FS", "시시": "SS", "완완": "FF", "시완": "SF"}
DEP_RE = re.compile(r"^(\d+)\s*(FS|SS|FF|SF|완시|시시|완완|시완)?\s*(?:([+-])\s*([0-9]+(?:\.[0-9]+)?)\s*(%|[^\d\s%]*))?$", re.I)


def parse_date(v, century: int = 2000) -> tuple[date | None, str | None]:
    """→ (날짜, 적힌 요일 글자). 못 읽으면 (None, None). 빈 칸은 (None, '')."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, ""
    if isinstance(v, datetime):
        return v.date(), ""
    if isinstance(v, date):
        return v, ""
    m = DATE_RE.search(nfkc(v))
    if not m:
        return None, None
    y, mo, d, wd = m.groups()
    y = int(y) + (century if len(y) == 2 else 0)
    try:
        return date(y, int(mo), int(d)), wd or ""
    except ValueError:
        return None, None


def to_days(num: float, unit: str, units: dict) -> float | None:
    u = nfkc(unit).strip().lower()
    if u in ("", "일"):
        return num
    return num * units[u] if u in units else None


def parse_duration(v, units: dict) -> tuple[float | None, bool, bool]:
    """→ (일수, 추정 '?', 읽었나). 빈 칸은 (None, False, True)."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, False, True
    if isinstance(v, (int, float)):
        return float(v), False, True
    m = DUR_RE.match(nfkc(v).strip())
    if not m:
        return None, False, False
    days = to_days(float(m.group(1).replace(",", "")), m.group(2), units)
    return days, bool(m.group(3)), days is not None


def parse_deps(v, units: dict) -> tuple[list[Dep], list[str]]:
    """'12,15FS+10일,20FF' → Dep 목록, 못 읽은 조각."""
    if v is None:
        return [], []
    if isinstance(v, (int, float)):
        return [Dep(int(v))], []
    deps, bad = [], []
    for tok in re.split(r"[,;]", nfkc(v)):
        tok = tok.strip()
        if not tok:
            continue
        m = DEP_RE.match(tok)
        if not m:
            bad.append(tok)
            continue
        no, typ, sign, num, unit = m.groups()
        typ = TYPE_KO.get(typ, (typ or "FS").upper())
        lag, text = 0.0, ""
        if sign:
            text = f"{sign}{num}{unit or ''}"
            if unit == "%":
                lag = None
            else:
                d = to_days(float(num), unit or "", units)
                lag = None if d is None else (d if sign == "+" else -d)
        deps.append(Dep(int(no), typ, lag, text))
    return deps, bad


def parse_number(v) -> tuple[float | None, bool]:
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, True
    if isinstance(v, (int, float)):
        return float(v), True
    m = re.search(r"-?[0-9][0-9,]*(?:\.[0-9]+)?", nfkc(v))
    if not m:
        return None, False
    return float(m.group(0).replace(",", "")), True


def parse_int(v) -> int | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    s = nfkc(v).strip() if v is not None else ""
    return int(s) if s.isdigit() else None


def norm_wbs(v) -> tuple[str, bool]:
    """→ (코드, 숫자 서식이었나). 끝의 '.' 은 뗀다."""
    if v is None:
        return "", False
    if isinstance(v, (int, float)):
        s = str(int(v)) if float(v).is_integer() else repr(float(v))
        return s, True
    return nfkc(v).strip().rstrip("."), False


# --- 시트 읽기 -----------------------------------------------------------------------------------

def find_header(ws, cfg: dict) -> tuple[int, dict[str, int], dict[str, str], list[str]]:
    want = {}
    for key, names in cfg["columns"].items():
        for n in names:
            want.setdefault(norm_header(n), key)
    persons = {norm_header(n) for n in cfg.get("person_columns", [])}
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=cfg.get("header_scan_rows", 15), values_only=True), start=1):
        cols, found, skipped = {}, {}, []
        for c, v in enumerate(row, start=1):
            if v is None or not str(v).strip():
                continue
            h = norm_header(v)
            key = want.get(h)
            if key and key not in cols and h not in persons:
                cols[key], found[key] = c, nfkc(v).strip()
            else:
                skipped.append(nfkc(v).strip())
        if all(k in cols for k in cfg["required"]):
            return r, cols, found, skipped
    return 0, {}, {}, []


def read_wbs(path: str | Path, cfg: dict, sheet: str | None = None) -> Wbs:
    path = Path(path)
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    issues: list[Issue] = []
    hr, cols, found, skipped = find_header(ws, cfg)
    if not hr:
        wb.close()
        issues.append(Issue("P01", detail=f"필수 칸 {', '.join(cfg['required'])} 의 헤더가 위 {cfg.get('header_scan_rows', 15)}행 안에 없음"))
        return Wbs(str(path), ws.title, {}, issues, {}, [])
    units = {nfkc(k).lower(): float(v) for k, v in cfg["unit_days"].items()}
    century = int(cfg.get("century", 2000))
    splits = cfg.get("deliverable_split", [","])
    tasks: dict[int, Task] = {}
    by_code: dict[str, int] = {}

    def cell(row, key):
        c = cols.get(key)
        return row[c - 1] if c and c - 1 < len(row) else None

    for r, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), start=hr + 1):
        if not row or all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        no = parse_int(cell(row, "row_no"))
        if no is None:
            issues.append(Issue("P02", excel_row=r))
            continue
        if no in tasks:
            issues.append(Issue("P03", no, r, detail=f"앞 행 {tasks[no].excel_row}"))
            continue
        code, numeric = norm_wbs(cell(row, "wbs"))
        if numeric:
            issues.append(Issue("P13", no, r, code))
        if not code or not re.fullmatch(r"[0-9A-Za-z]+(\.[0-9A-Za-z]+)*", code):
            issues.append(Issue("P04", no, r, code))
            continue
        if code in by_code:
            issues.append(Issue("P05", no, r, code, detail=f"번호 {by_code[code]} 와 같음"))
            continue
        raw_name = cell(row, "name")
        name_s = "" if raw_name is None else str(raw_name).replace("\t", "    ")
        indent = len(name_s) - len(name_s.lstrip(" 　 "))
        t = Task(no, r, code, code.count(".") + 1, nfkc(name_s).strip(), indent)
        # 기간
        dv = cell(row, "duration")
        t.duration_text = "" if dv is None else nfkc(dv).strip()
        t.duration_days, t.estimated, ok = parse_duration(dv, units)
        if not ok:
            issues.append(Issue("P10", no, r, code, detail=t.duration_text))
        # 날짜
        for key in ("start", "finish"):
            d, wd = parse_date(cell(row, key), century)
            if wd is None:
                issues.append(Issue("P08", no, r, code, detail=f"{found.get(key, key)}: {nfkc(cell(row, key))}"))
            elif wd and d and WEEKDAYS_KO[d.weekday()] != wd:
                issues.append(Issue("P09", no, r, code, detail=f"{d.isoformat()} 은 {WEEKDAYS_KO[d.weekday()]}요일 (적힌 요일 {wd})"))
            setattr(t, key, d)
        # 선행·후속
        t.deps, bad = parse_deps(cell(row, "preds"), units)
        for b in bad:
            issues.append(Issue("P11", no, r, code, detail=b))
        if "succs" in cols:
            succs, _ = parse_deps(cell(row, "succs"), units)
            t.succs_declared = [d.pred for d in succs]
        # 숫자
        for key, attr in (("staff", "staff"), ("man_hours", "man_hours")):
            val, ok = parse_number(cell(row, key))
            if not ok:
                issues.append(Issue("P12", no, r, code, detail=f"{found.get(key, key)}: {nfkc(cell(row, key))}"))
            setattr(t, attr, val)
        # 산출물
        dl = cell(row, "deliverable")
        if dl is not None and str(dl).strip():
            parts = [nfkc(dl)]
            for sep in ["\n"] + list(splits):
                parts = [p for x in parts for p in x.split(sep)]
            t.deliverables = [p.strip() for p in parts if p.strip()]
        tasks[no] = t
        by_code[code] = no
    wb.close()
    link_tree(tasks, by_code, issues, cfg)
    return Wbs(str(path), ws.title, tasks, issues, found, skipped)


def link_tree(tasks: dict[int, Task], by_code: dict[str, int], issues: list[Issue], cfg: dict) -> None:
    """WBS 코드로 상위·하위를 잇고, 구성품·개발 단계·기술 분야 이름을 물려준다."""
    for t in tasks.values():
        if t.level == 1:
            continue
        pcode = t.wbs.rsplit(".", 1)[0]
        p = by_code.get(pcode)
        if p is None:
            issues.append(Issue("P06", t.no, t.excel_row, t.wbs, detail=f"상위 코드 {pcode}"))
            continue
        t.parent = p
        tasks[p].children.append(t.no)
        if t.indent <= tasks[p].indent and (t.indent or tasks[p].indent):
            issues.append(Issue("P07", t.no, t.excel_row, t.wbs,
                                detail=f"들여쓰기 {t.indent}칸 · 상위({tasks[p].wbs}) {tasks[p].indent}칸"))
    lv = {"module": cfg["module_level"], "phase": cfg["phase_level"], "discipline": cfg["discipline_level"]}
    for t in sorted(tasks.values(), key=lambda x: x.level):
        p = tasks[t.parent] if t.parent is not None else None
        for attr, level in lv.items():
            if t.level == level:
                setattr(t, attr, t.name)
            elif p is not None:
                setattr(t, attr, getattr(p, attr))
