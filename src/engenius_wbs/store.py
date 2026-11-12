"""저장소 (v0.5 S1) — WBS 의 정본을 SQLite 파일 하나에 둔다. 엑셀은 들여오기·내보내기 형식 (docs/05).

  import_xlsx  엑셀 → DB. 업무마다 고유 ID(uid, 다시 쓰지 않음)를 붙이고, 날짜가 하나도 움직이지 않게 '시작 하한'과 '느슨함'을 굳힌다.
  load         DB → 기존 엔진의 Wbs 객체 (analyze·report·viz·rdfout 을 그대로 쓴다). 번호·WBS 코드·레벨은 트리 위치로 계산.
  recompute    날짜를 앞으로 다시 짠다 (schedule.forward) — 값이 바뀐 업무만 고쳐 쓴다.
  export_xlsx  DB → 원본 양식 엑셀 (서식 틀은 들여올 때 값을 지운 채 DB 에 보관 — 원본 자료·인명은 보관하지 않는다).

DB 에 두는 것과 계산하는 것:
  두는 것  트리(parent_uid·pos) · 이름 · 적힌 기간(duration_days, 원래 글자 dur_text) · 일하는 평일 수(work_days) · 시작 하한(start_floor)
           · 날짜(start·finish, 원래 글자 모양 start_text·finish_text) · 투입인력 · Man Hours(가장 아래 적힌 값만) · 산출물 · 선행(dep)
  계산     번호(앞에서부터 순서) · WBS 코드(형제 순서) · 레벨(깊이) · 후속 칸 · 요약의 날짜·Man Hours 합
인명 칸은 읽지 않고, 내보낼 때 비워 둔다. 설정에 없는 칸은 버린다 (keep_columns 에 헤더를 적으면 값 그대로 보관).
"""
from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.packaging.core import DocumentProperties
from openpyxl.utils import get_column_letter

from .analyze import Edge, analyze
from .apply import date_like, duration_like, sha256_of
from .config import nfkc, norm_header
from .parse import DEP_RE, Dep, Task, Wbs, find_header, link_tree, parse_date, parse_duration, read_wbs
from .schedule import EPS, ONE, event_values, finish_date, forward, start_date, widx

SCHEMA_VERSION = 1
INDENT_CHARS = " 　 "
CELL = "@cell"          # start_text·finish_text: 엑셀 날짜 셀이었음 / dur_text: 숫자 셀이었음
UID_HEADER = "UID"

SCHEMA = """
CREATE TABLE meta (key TEXT PRIMARY KEY, value);
CREATE TABLE task (
  uid INTEGER PRIMARY KEY AUTOINCREMENT,      -- 고유 ID. 지워도 다시 쓰지 않는다
  parent_uid INTEGER REFERENCES task(uid),    -- 위 묶음 (없으면 맨 위)
  pos INTEGER NOT NULL,                       -- 형제 사이 순서
  name TEXT NOT NULL DEFAULT '',
  duration_days REAL, dur_text TEXT,          -- 적힌 기간(일) · 원래 글자
  work_days REAL,                             -- 시작~완료 평일 수 (일정 계산은 이 값으로 — 휴일 때문에 적힌 기간과 다를 수 있다)
  start_floor TEXT,                           -- 시작 하한 (이 날 이후). 선행에 바로 밀리던 업무는 비어 있다
  start TEXT, finish TEXT,                    -- 날짜 (계산 결과를 적어 둔 것)
  start_text TEXT, finish_text TEXT,          -- 원래 칸 글자 (날짜가 그대로면 그대로 내보낸다)
  staff REAL, man_hours REAL,
  mh_sum INTEGER NOT NULL DEFAULT 0,          -- 요약: 원본이 Man Hours 칸에 합을 적어 두었나 (내보낼 때 아래 합을 적는다)
  deliverables TEXT, extra TEXT,              -- extra = keep_columns 값 (JSON)
  deleted INTEGER NOT NULL DEFAULT 0, rev INTEGER NOT NULL DEFAULT 1,
  src_no INTEGER, src_wbs TEXT, src_row INTEGER   -- 들여온 엑셀에서의 번호·WBS 코드·행 (추적용)
);
CREATE INDEX task_parent ON task(parent_uid, pos);
CREATE TABLE dep (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  succ_uid INTEGER NOT NULL REFERENCES task(uid), pred_uid INTEGER NOT NULL REFERENCES task(uid),
  ord INTEGER NOT NULL,                       -- 선행 칸 안 순서
  type TEXT NOT NULL DEFAULT 'FS', lag_days REAL, lag_text TEXT NOT NULL DEFAULT '',
  suffix TEXT NOT NULL DEFAULT '',            -- 원래 표기에서 번호 뒤 ('SS+5일', '완시' …)
  relax_days REAL NOT NULL DEFAULT 0          -- 들여올 때 이미 어겨 있던 평일 수 (그만큼 느슨하게 본다)
);
CREATE INDEX dep_succ ON dep(succ_uid, ord);
CREATE INDEX dep_pred ON dep(pred_uid);
CREATE TABLE txn (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, role TEXT NOT NULL, kind TEXT NOT NULL, note TEXT);
CREATE TABLE change (id INTEGER PRIMARY KEY AUTOINCREMENT, txn_id INTEGER NOT NULL REFERENCES txn(id),
  task_uid INTEGER, field TEXT NOT NULL, old TEXT, new TEXT);
CREATE TABLE note (id INTEGER PRIMARY KEY AUTOINCREMENT, txn_id INTEGER NOT NULL REFERENCES txn(id),
  task_uid INTEGER, col TEXT NOT NULL, old TEXT, new TEXT);   -- 들여올 때 원본과 달라진 칸 (내보내면 이렇게 나온다)
"""


@dataclass
class Model:
    wbs: Wbs
    rows: dict[int, sqlite3.Row]                # 번호 → task 행
    deps: dict[int, list[sqlite3.Row]]          # 번호 → dep 행 (칸 안 순서)
    edges: list[Edge]
    relax: list[float]
    floor: dict[int, float]
    work: dict[int, float]


@dataclass
class ImportReport:
    tasks: int = 0
    deps: int = 0
    moved: int = 0                              # 굳힌 뒤 날짜가 달라진 하위 업무 (0 이어야 한다)
    floors: int = 0                             # 시작 하한을 둔 하위 업무
    relaxed: int = 0                            # 느슨함을 둔 선행
    notes: dict[str, int] = field(default_factory=dict)        # 칸 → 원본과 달라진 칸 수
    dropped_columns: dict[str, int] = field(default_factory=dict)   # 버린 헤더 → 값이 있던 칸 수
    kept_columns: list[str] = field(default_factory=list)
    person_columns: int = 0                     # 비워 둘 인명 칸 수 (헤더 문구도 적지 않는다)
    skipped_rows: int = 0
    cleared_top_cells: int = 0                  # 헤더 위쪽에서 지운 칸 수
    issues: dict[str, int] = field(default_factory=dict)


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def same_cell(a, b) -> bool:
    if _blank(a) and _blank(b):
        return True
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-6 * max(1.0, abs(a))
    return type(a) is type(b) and a == b


def _num(x):
    if x is None:
        return None
    return int(x) if float(x).is_integer() else x


def _iso(d: date | None) -> str | None:
    return d.isoformat() if d else None


def _date(s: str | None) -> date | None:
    return date.fromisoformat(s) if s else None


def _units(cfg: dict) -> dict:
    return {nfkc(k).lower(): float(v) for k, v in cfg["unit_days"].items()}


def _own(t: Task) -> tuple[float, float] | None:
    """업무에 적혀 있는 날짜의 사건 값 (event_values 와 같은 규칙, 요약도 자기 날짜로)."""
    if not (t.start and t.finish) or t.start > t.finish:
        return None
    f = widx(t.finish + ONE)
    return (f, f) if t.is_milestone else (widx(t.start), f)


def _dates(s: float, f: float) -> tuple[date, date]:
    return (start_date(s) if f > s else finish_date(f)), finish_date(f)


def _dep_suffix(d: Dep) -> str:
    return ("" if d.type == "FS" else d.type) + d.lag_text


def split_deliverables(text, cfg: dict) -> list[str]:
    if text is None or not str(text).strip():
        return []
    parts = [nfkc(text)]
    for sep in ["\n"] + list(cfg.get("deliverable_split", [","])):
        parts = [p for x in parts for p in x.split(sep)]
    return [p.strip() for p in parts if p.strip()]


class Store:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(self.path))
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA foreign_keys = ON")
        if not self.db.execute("SELECT 1 FROM sqlite_master WHERE name = 'meta'").fetchone():
            self.db.executescript(SCHEMA)
            self.set_meta("schema_version", SCHEMA_VERSION)
            self.db.commit()
        elif self.meta("schema_version") != SCHEMA_VERSION:
            raise ValueError(f"DB 판({self.meta('schema_version')})이 이 코드({SCHEMA_VERSION})와 다르다: {self.path}")

    def close(self) -> None:
        self.db.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # --- meta -----------------------------------------------------------------------------------
    def meta(self, key: str, default=None):
        r = self.db.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
        if r is None:
            return default
        v = r["value"]
        return json.loads(v) if isinstance(v, str) else v

    def set_meta(self, key: str, value) -> None:
        v = value if isinstance(value, bytes) else json.dumps(value, ensure_ascii=False)
        self.db.execute("INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value", (key, v))

    def count(self) -> int:
        return self.db.execute("SELECT COUNT(*) FROM task WHERE deleted = 0").fetchone()[0]

    def _txn(self, kind: str, role: str = "system", note: str = "") -> int:
        cur = self.db.execute("INSERT INTO txn(at, role, kind, note) VALUES(?, ?, ?, ?)",
                              (datetime.now().isoformat(timespec="seconds"), role, kind, note))
        return cur.lastrowid

    # --- DB → Wbs -------------------------------------------------------------------------------
    def model(self, cfg: dict) -> Model:
        units = _units(cfg)
        kids: dict[int | None, list] = defaultdict(list)
        for r in self.db.execute("SELECT * FROM task WHERE deleted = 0 ORDER BY pos, uid"):
            kids[r["parent_uid"]].append(r)
        unit = int(self.meta("indent_unit", 3))
        hr = int(self.meta("header_row", 1))
        tasks: dict[int, Task] = {}
        rows: dict[int, sqlite3.Row] = {}
        by_code: dict[str, int] = {}
        uids: dict[int, int] = {}
        no_of: dict[int, int] = {}
        stack = [(r, 1, str(i + 1)) for i, r in enumerate(kids[None])][::-1]
        while stack:
            r, lv, code = stack.pop()
            no = len(tasks) + 1
            txt = r["dur_text"]
            est = False
            if txt and txt != CELL:
                est = parse_duration(txt, units)[1]
                dtext = nfkc(txt).strip()
            else:
                dtext = "" if r["duration_days"] is None else f"{r['duration_days']:g}"
            t = Task(no, hr + no, code, lv, nfkc(r["name"]).strip(), unit * (lv - 1), r["duration_days"], dtext, est,
                     _date(r["start"]), _date(r["finish"]), staff=r["staff"], man_hours=r["man_hours"],
                     deliverables=split_deliverables(r["deliverables"], cfg))
            tasks[no], rows[no], by_code[code], uids[no], no_of[r["uid"]] = t, r, no, r["uid"], no
            stack.extend([(c, lv + 1, f"{code}.{i + 1}") for i, c in enumerate(kids.get(r["uid"], ()))][::-1])
        deps: dict[int, list] = defaultdict(list)
        for d in self.db.execute("SELECT * FROM dep ORDER BY succ_uid, ord, id"):
            s, p = no_of.get(d["succ_uid"]), no_of.get(d["pred_uid"])
            if s is not None and p is not None:
                deps[s].append(d)
                tasks[s].deps.append(Dep(p, d["type"], d["lag_days"], d["lag_text"]))
        link_tree(tasks, by_code, [], cfg)
        edges, relax = [], []
        for t in tasks.values():
            for d, row in zip(t.deps, deps[t.no]):
                edges.append(Edge(d.pred, t.no, d.type, d.lag_days, d.lag_text))
                relax.append(row["relax_days"] or 0.0)
        floor, work = {}, {}
        for no, r in rows.items():
            t = tasks[no]
            if t.children:
                continue
            if r["start_floor"]:
                d = _date(r["start_floor"])
                floor[no] = widx(d + ONE) if t.is_milestone else widx(d)
            if r["work_days"] is not None:
                work[no] = r["work_days"]
        w = Wbs(self.meta("source_name", self.path.name), self.meta("sheet", "WBS"), tasks, [], self.meta("headers", {}), [], uids)
        return Model(w, rows, deps, edges, relax, floor, work)

    def load(self, cfg: dict) -> Wbs:
        return self.model(cfg).wbs

    # --- 날짜 다시 짜기 ---------------------------------------------------------------------------
    def recompute(self, cfg: dict, m: Model | None = None) -> list[int]:
        """앞으로 다시 짜서 날짜가 달라진 업무만 고쳐 쓴다 → 바뀐 uid."""
        m = m or self.model(cfg)
        tasks = m.wbs.tasks
        val0, _ = event_values(tasks)
        val, _ = forward(tasks, m.edges, val0, m.floor, m.work, m.relax)
        changed = []
        for t in tasks.values():
            s, f = val.get(2 * t.no), val.get(2 * t.no + 1)
            if s is None or f is None:
                continue
            own = _own(t)
            if own and abs(own[0] - s) < EPS and abs(own[1] - f) < EPS:
                continue
            ns, nf = _dates(s, f)
            r = m.rows[t.no]
            dur = r["duration_days"]
            if t.children:
                dur = (f - s) if own is None or dur is None else max(dur + (f - s) - (own[1] - own[0]), 0.0)
            self.db.execute("UPDATE task SET start = ?, finish = ?, duration_days = ?, rev = rev + 1 WHERE uid = ?",
                            (_iso(ns), _iso(nf), dur, r["uid"]))
            changed.append(r["uid"])
        return changed

    # --- 엑셀 → DB -------------------------------------------------------------------------------
    def import_xlsx(self, path: str | Path, cfg: dict, sheet: str | None = None) -> ImportReport:
        path = Path(path)
        if self.count():
            raise ValueError(f"이미 업무가 들어 있는 DB 다: {self.path} (새 파일에 들여오거나 --force)")
        w = read_wbs(path, cfg, sheet)
        if not w.tasks:
            raise ValueError("읽은 업무가 없다: " + " / ".join(f"{i.code} {i.title} {i.detail}" for i in w.issues[:3]))
        an = analyze(w, cfg)
        tasks = w.tasks
        rep = ImportReport(tasks=len(tasks), issues=dict(sorted(Counter(i.code for i in w.issues + an.issues).items())))
        wb = load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        hr, cols, found, _ = find_header(ws, cfg)
        raw = {r: row for r, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), start=hr + 1)}
        ncol = ws.max_column

        def cell(r, key):
            c = cols.get(key)
            row = raw.get(r, ())
            return row[c - 1] if c and c - 1 < len(row) else None

        # 칸 분류: 읽는 칸 · 인명 칸(비움) · 보관 칸(keep_columns) · 버리는 칸
        persons = {norm_header(n) for n in cfg.get("person_columns", [])}
        keep = {norm_header(n) for n in cfg.get("keep_columns") or []}
        mapped = set(cols.values())
        keep_cols: dict[str, int] = {}
        drop_cols: dict[int, str] = {}
        for c in range(1, ncol + 1):
            v = ws.cell(hr, c).value
            if c in mapped or _blank(v):
                continue
            h = norm_header(v)
            if h in persons:
                rep.person_columns += 1
            elif h in keep:
                keep_cols[nfkc(v).strip()] = c
            elif h != norm_header(UID_HEADER):
                drop_cols[c] = nfkc(v).strip()
        rep.kept_columns = list(keep_cols)
        used = {t.excel_row for t in tasks.values()}
        for c, h in drop_cols.items():
            rep.dropped_columns[h] = sum(1 for r in used if c - 1 < len(raw[r]) and not _blank(raw[r][c - 1]))
        rep.skipped_rows = sum(1 for r, row in raw.items() if r not in used and any(not _blank(v) for v in row))

        # 글자 습관: 들여쓰기 단위 · 선행 구분자 · 숫자 한 개짜리 선행·후속 칸이 숫자 셀인가 · 후속 칸에 종류를 적나 · 날짜·기간 본보기
        ind = Counter(t.indent // (t.level - 1) for t in tasks.values() if t.level > 1 and t.indent % (t.level - 1) == 0)
        seps: Counter = Counter()
        pred_int = succ_int = succ_typed = False
        samples = {"date": None, "dur": None}
        for t in tasks.values():
            pv, sv = cell(t.excel_row, "preds"), cell(t.excel_row, "succs")
            if isinstance(pv, str):
                seps.update(re.findall(r"\s*[,;]\s*", pv))
            pred_int = pred_int or (isinstance(pv, int) and not isinstance(pv, bool))
            succ_int = succ_int or (isinstance(sv, int) and not isinstance(sv, bool))
            if isinstance(sv, str) and not succ_typed:
                succ_typed = any(not tok.strip().isdigit() and DEP_RE.match(tok.strip()) for tok in re.split(r"[,;]", nfkc(sv)))
            for key, kind in (("start", "date"), ("finish", "date"), ("duration", "dur")):
                v = cell(t.excel_row, key)
                if samples[kind] is None and not _blank(v):
                    samples[kind] = v if isinstance(v, str) else CELL

        # 트리 순서대로 uid (엑셀 행 순서 = 앞에서부터 순서일 때 uid = 번호 순)
        order: list[tuple[int, int | None]] = []            # (번호, 위 묶음 번호)
        roots = sorted((t for t in tasks.values() if t.parent is None), key=lambda t: t.excel_row)
        stack = [(t.no, None) for t in roots][::-1]
        while stack:
            no, parent = stack.pop()
            order.append((no, parent))
            stack.extend([(c, no) for c in sorted(tasks[no].children, key=lambda c: tasks[c].excel_row)][::-1])
        uid_of = {no: i + 1 for i, (no, _) in enumerate(order)}

        # 시작 하한·느슨함 — 고치지 않으면 날짜가 하나도 안 움직이게
        val0 = an.schedule.val
        leaves = [t.no for t in tasks.values() if not t.children and 2 * t.no in val0]
        work = {no: val0[2 * no + 1] - val0[2 * no] for no in leaves}
        relax = [(-g if g is not None and g < -EPS else 0.0) for g in an.schedule.gaps]
        _, need = forward(tasks, an.edges, val0, {no: val0[2 * no] for no in leaves}, work, relax)
        floor = {no: val0[2 * no] for no in leaves if no not in need or need[no] < val0[2 * no] - EPS}
        val1, _ = forward(tasks, an.edges, val0, floor, work, relax)
        rep.moved = sum(1 for no in leaves if abs(val1[2 * no] - val0[2 * no]) > EPS)
        rep.floors, rep.relaxed = len(floor), sum(1 for x in relax if x)

        txn = self._txn("import", note=json.dumps({"file": path.name, "sha256": sha256_of(path), "tasks": len(tasks)}, ensure_ascii=False))
        mh_below = an.workload.mh_below
        pos: Counter = Counter()

        def text_of(v):
            return None if _blank(v) else (v if isinstance(v, str) else CELL)

        for no, parent in order:
            t = tasks[no]
            r = t.excel_row
            nm = cell(r, "name")
            name = "" if nm is None else str(nm).replace("\t", "    ").lstrip(INDENT_CHARS).rstrip()
            fl = floor.get(no)
            fdate = None if fl is None else (finish_date(fl) if t.is_milestone else start_date(fl))
            below = mh_below.get(no) if t.children else None
            dl = cell(r, "deliverable")
            extra = {h: raw[r][c - 1] for h, c in keep_cols.items() if c - 1 < len(raw[r]) and not _blank(raw[r][c - 1])}
            extra = {h: (v.isoformat() if isinstance(v, (date, datetime)) else v) for h, v in extra.items()}
            self.db.execute(
                "INSERT INTO task(uid, parent_uid, pos, name, duration_days, dur_text, work_days, start_floor, start, finish, start_text,"
                " finish_text, staff, man_hours, mh_sum, deliverables, extra, src_no, src_wbs, src_row) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (uid_of[no], uid_of.get(parent), pos[parent], name, t.duration_days, text_of(cell(r, "duration")), work.get(no), _iso(fdate),
                 _iso(t.start), _iso(t.finish), text_of(cell(r, "start")), text_of(cell(r, "finish")), t.staff,
                 None if below is not None else t.man_hours, int(below is not None and t.man_hours is not None),
                 None if _blank(dl) else str(dl), json.dumps(extra, ensure_ascii=False) if extra else None, no, t.wbs, r))
            pos[parent] += 1
        # 선행 — 읽을 수 있고 두 끝이 있는 것만 (an.edges 와 같은 순서)
        ei = 0
        for t in tasks.values():
            pv = cell(t.excel_row, "preds")
            toks = [mt for mt in (DEP_RE.match(x.strip()) for x in re.split(r"[,;]", nfkc(pv))) if mt] if isinstance(pv, str) else []
            aligned = len(toks) == len(t.deps) and all(int(mt.group(1)) == d.pred for mt, d in zip(toks, t.deps))
            for k, d in enumerate(t.deps):
                if d.pred == t.no or d.pred not in tasks:
                    continue
                suffix = toks[k].group(0)[len(toks[k].group(1)):] if aligned else _dep_suffix(d)
                self.db.execute("INSERT INTO dep(succ_uid, pred_uid, ord, type, lag_days, lag_text, suffix, relax_days) VALUES(?,?,?,?,?,?,?,?)",
                                (uid_of[t.no], uid_of[d.pred], k, d.type, d.lag_days, d.lag_text, suffix, relax[ei]))
                ei += 1
        if ei != len(an.edges):
            raise RuntimeError(f"선행 수가 검사 엔진과 다르다 ({ei} / {len(an.edges)})")
        rep.deps = ei
        for k, v in {"header_row": hr, "sheet": ws.title, "source_name": path.name, "source_sha256": sha256_of(path), "columns": cols,
                     "headers": found, "keep_columns": keep_cols, "indent_unit": ind.most_common(1)[0][0] if ind else 0,
                     "pred_sep": seps.most_common(1)[0][0] if seps else ",", "pred_int": pred_int, "succ_int": succ_int,
                     "succ_typed": succ_typed, "date_sample": samples["date"], "dur_sample": samples["dur"],
                     "imported_at": datetime.now().isoformat(timespec="seconds")}.items():
            self.set_meta(k, v)
        self.recompute(cfg)                      # 요약 날짜를 하위 범위로 (C11)
        # 원본과 달라지는 칸을 적어 둔다 (내보내면 이렇게 나온다)
        m = self.model(cfg)
        notes: Counter = Counter()
        for no, vals in self._values(m, cfg).items():
            row = m.rows[no]
            src = raw[row["src_row"]]
            for key, c in cols.items():
                old = src[c - 1] if c - 1 < len(src) else None
                if not same_cell(old, vals.get(c)):
                    notes[key] += 1
                    self.db.execute("INSERT INTO note(txn_id, task_uid, col, old, new) VALUES(?,?,?,?,?)",
                                    (txn, row["uid"], key, None if old is None else str(old)[:200], None if vals.get(c) is None else str(vals[c])[:200]))
        rep.notes = dict(notes)
        blob, rep.cleared_top_cells = self._template(wb, ws, hr, tasks)
        self.set_meta("template", blob)
        self.db.commit()
        return rep

    def _template(self, wb, ws, hr: int, tasks: dict[int, Task]) -> tuple[bytes, int]:
        """서식 틀 — 헤더 행과 레벨마다 서식 본보기 한 줄만 남기고 값을 다 지운다. 다른 시트·메모·문서 속성도 지운다."""
        for other in list(wb.worksheets):
            if other is not ws:
                wb.remove(other)
        ncol = ws.max_column
        first: dict[int, int] = {}
        for t in sorted(tasks.values(), key=lambda t: t.excel_row):
            first.setdefault(t.level, t.excel_row)
        levels = sorted(first)
        samples = [([copy(ws.cell(first[lv], c)._style) for c in range(1, ncol + 1)], ws.row_dimensions[first[lv]].height) for lv in levels]
        ws.delete_rows(hr + 1, ws.max_row - hr)
        for r in [r for r in ws.row_dimensions if r > hr]:
            del ws.row_dimensions[r]
        for rng in [rng for rng in ws.merged_cells.ranges if rng.max_row > hr or rng.min_row < hr]:
            ws.unmerge_cells(str(rng))
        cleared = 0
        for row in ws.iter_rows(min_row=1, max_row=hr):
            for c in row:
                c.comment = None
                c.hyperlink = None
                if c.row < hr and c.value is not None:
                    c.value = None
                    cleared += 1
        for i, (styles, height) in enumerate(samples):
            for c, st in enumerate(styles, start=1):
                ws.cell(hr + 1 + i, c)._style = st
            if height:
                ws.row_dimensions[hr + 1 + i].height = height
        wb.properties = DocumentProperties()
        props = getattr(wb, "custom_doc_props", None)
        if props is not None:
            for name in [p.name for p in props.props]:
                del props[name]
        self.set_meta("template_levels", levels)
        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue(), cleared

    # --- DB → 엑셀 -------------------------------------------------------------------------------
    def _values(self, m: Model, cfg: dict) -> dict[int, dict[int, object]]:
        """번호 → {칸 번호: 내보낼 값}."""
        units = _units(cfg)
        century = int(cfg.get("century", 2000))
        cols = self.meta("columns")
        keep_cols = self.meta("keep_columns", {})
        unit, sep = int(self.meta("indent_unit", 3)), self.meta("pred_sep", ",")
        pred_int, succ_int, succ_typed = self.meta("pred_int", False), self.meta("succ_int", False), self.meta("succ_typed", True)
        date_sample, dur_sample = self.meta("date_sample"), self.meta("dur_sample")
        tasks = m.wbs.tasks
        below: dict[int, float] = {}
        for t in sorted(tasks.values(), key=lambda x: -x.level):
            if t.children and any(c in below or tasks[c].man_hours for c in t.children):
                below[t.no] = sum(below[c] if c in below else (tasks[c].man_hours or 0.0) for c in t.children)
        succs: dict[int, list[str]] = defaultdict(list)
        if "succs" in cols:
            for t in tasks.values():
                for d in t.deps:
                    succs[d.pred].append(f"{t.no}" + (_dep_suffix(d) if succ_typed else ""))

        def date_cell(text, d):
            if d is None:
                return None if text in (None, CELL) else text
            if text == CELL or (text is None and date_sample == CELL):
                return datetime(d.year, d.month, d.day)
            if text is not None and parse_date(text, century)[0] == d:
                return text
            return date_like(text if text is not None else date_sample, d)

        def dur_cell(text, days):
            if days is None:
                return None if text in (None, CELL) else text
            if text == CELL or (text is None and dur_sample == CELL):
                return _num(days)
            if text is not None and parse_duration(text, units)[0] == days:
                return text
            return duration_like(text if text is not None else dur_sample, days)

        def join(toks, as_int):
            s = sep.join(toks) if toks else None
            return int(s) if as_int and s and s.isdigit() else s

        out = {}
        for no, t in tasks.items():
            r = m.rows[no]
            v = {cols["row_no"]: no, cols["wbs"]: t.wbs, cols["name"]: " " * (unit * (t.level - 1)) + r["name"]}
            if "duration" in cols:
                v[cols["duration"]] = dur_cell(r["dur_text"], r["duration_days"])
            if "start" in cols:
                v[cols["start"]] = date_cell(r["start_text"], t.start)
            if "finish" in cols:
                v[cols["finish"]] = date_cell(r["finish_text"], t.finish)
            if "preds" in cols:
                v[cols["preds"]] = join([f"{d.pred}{row['suffix']}" for d, row in zip(t.deps, m.deps[no])], pred_int)
            if "succs" in cols:
                v[cols["succs"]] = join(succs.get(no), succ_int)
            if "staff" in cols:
                v[cols["staff"]] = _num(r["staff"])
            if "man_hours" in cols:
                v[cols["man_hours"]] = _num(below[no] if r["mh_sum"] else None) if no in below else _num(r["man_hours"])
            if "deliverable" in cols:
                v[cols["deliverable"]] = r["deliverables"]
            if r["extra"]:
                ex = json.loads(r["extra"])
                for h, c in keep_cols.items():
                    v[c] = ex.get(h)
            out[no] = v
        return out

    def export_xlsx(self, out: str | Path, cfg: dict) -> Path:
        """원본 양식 그대로 — 헤더·칸 순서·칸 너비·레벨별 서식. 인명 칸은 비고, 맨 끝에 숨긴 UID 칸을 붙인다."""
        out = Path(out)
        blob = self.meta("template")
        if not blob:
            raise ValueError("서식 틀이 없다 — 엑셀을 들여온 DB 가 아니다")
        m = self.model(cfg)
        wb = load_workbook(BytesIO(blob))
        ws = wb.worksheets[0]
        hr = int(self.meta("header_row", 1))
        levels = self.meta("template_levels", [])
        ncol = ws.max_column
        uid_col = next((c for c in range(1, ncol + 1) if norm_header(ws.cell(hr, c).value or "") == norm_header(UID_HEADER)), ncol + 1)
        samples = {lv: ([copy(ws.cell(hr + 1 + i, c)._style) for c in range(1, ncol + 1)], ws.row_dimensions[hr + 1 + i].height)
                   for i, lv in enumerate(levels)}
        for r in [r for r in ws.row_dimensions if r > hr]:
            del ws.row_dimensions[r]
        deepest = max(levels) if levels else None
        values = self._values(m, cfg)
        for no, t in m.wbs.tasks.items():
            r = hr + no
            styles, height = samples.get(t.level) or samples.get(deepest) or ([], None)
            vals = values[no]
            for c in range(1, ncol + 1):
                cl = ws.cell(r, c)
                if styles:
                    cl._style = copy(styles[c - 1])
                cl.value = vals.get(c)
            ws.cell(r, uid_col).value = m.rows[no]["uid"]
            if height:
                ws.row_dimensions[r].height = height
        for r in range(hr + len(m.wbs.tasks) + 1, hr + len(levels) + 1):     # 업무가 본보기 줄보다 적을 때 남는 줄
            for c in range(1, ncol + 1):
                ws.cell(r, c).value = None
        ws.cell(hr, uid_col).value = UID_HEADER
        ws.column_dimensions[get_column_letter(uid_col)].hidden = True
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A{hr}:{get_column_letter(max(ncol, uid_col))}{hr + len(m.wbs.tasks)}"
        out.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out)
        return out
