"""저장소 (v0.5 S1) — 엑셀 → DB → 엑셀 왕복. 달라지는 칸은 들여오기 기록(note)에 있는 것뿐이어야 한다."""
from __future__ import annotations

import copy
from collections import Counter
from datetime import date

import pytest
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from engenius_wbs.__main__ import main
from engenius_wbs.analyze import analyze
from engenius_wbs.parse import parse_date, parse_duration, read_wbs
from engenius_wbs.schedule import EPS, event_values, forward, reschedule
from engenius_wbs.store import UID_HEADER, Store, same_cell
from engenius_wbs.synth import generate
from test_wbs import CFG, UNITS, make_small

KEEP = ("C03", "C04", "C05", "C06", "C08", "C09", "C10")      # 들여와도 그대로 남아야 하는 검사 (나머지는 들여오며 풀린다)


def diff_cells(a_path, b_path, hr: int) -> dict[tuple[int, int], tuple]:
    a, b = load_workbook(a_path).worksheets[0], load_workbook(b_path).worksheets[0]
    out = {}
    rows_b = list(b.iter_rows(min_row=hr + 1, values_only=True))
    for i, ra in enumerate(a.iter_rows(min_row=hr + 1, values_only=True)):
        rb = rows_b[i] if i < len(rows_b) else ()
        for c, x in enumerate(ra):
            y = rb[c] if c < len(rb) else None
            if not same_cell(x, y):
                out[(hr + 1 + i, c + 1)] = (x, y)
    return out


@pytest.fixture(scope="module")
def synth_store(tmp_path_factory):
    d = tmp_path_factory.mktemp("store")
    info = generate(d / "wbs.xlsx", 12827, 7)
    st = Store(d / "wbs.db")
    rep = st.import_xlsx(d / "wbs.xlsx", CFG)
    st.export_xlsx(d / "out.xlsx", CFG)
    yield d, info, st, rep
    st.close()


def test_synthetic_roundtrip_cells(synth_store):
    d, info, st, rep = synth_store
    inj = info["injected"]
    assert rep.tasks == 12827 and rep.moved == 0 and rep.skipped_rows == 0 and rep.dropped_columns == {}
    cols = st.meta("columns")
    key_of = {c: k for k, c in cols.items()}
    row_of = {r["uid"]: r["src_row"] for r in st.db.execute("SELECT uid, src_row FROM task")}
    no_of = {r["src_row"]: r["src_no"] for r in st.db.execute("SELECT src_no, src_row FROM task")}
    notes = {(row_of[n["task_uid"]], cols[n["col"]]) for n in st.db.execute("SELECT task_uid, col FROM note")}
    diff = diff_cells(d / "wbs.xlsx", d / "out.xlsx", st.meta("header_row"))
    person = {k for k in diff if k[1] not in key_of}             # 인명 칸 — 비워서 나간다
    assert all(diff[k][1] is None for k in person)
    assert set(diff) - person == notes                          # 달라진 칸 = 들여오기 기록, 빠짐도 남음도 없다
    # 달라진 칸마다 심어 둔 문제로 설명된다
    why = {"preds": {"C01", "C02", "P11"}, "succs": {"C02", "C07"}, "start": {"C11"}, "finish": {"C11"}, "duration": {"C11"},
           "man_hours": {"C12"}, "name": {"P07"}}
    for (r, c) in notes:
        codes = {k for k, nos in inj.items() if no_of[r] in nos}
        assert codes & why[key_of[c]], (r, key_of[c], codes)
    assert Counter(key_of[c] for _, c in notes) == Counter(rep.notes)
    # 숨긴 UID 칸
    ws = load_workbook(d / "out.xlsx").worksheets[0]
    uc = ws.max_column
    assert ws.cell(1, uc).value == UID_HEADER and ws.column_dimensions[ws.cell(1, uc).column_letter].hidden
    assert [ws.cell(r, uc).value for r in (2, 3, 12828)] == [1, 2, 12827]
    assert st.db.execute("SELECT kind FROM txn WHERE id = 1").fetchone()[0] == "import"


def test_synthetic_same_analysis(synth_store):
    d, info, st, rep = synth_store
    w0 = read_wbs(d / "wbs.xlsx", CFG)
    a0 = analyze(w0, CFG)
    w1 = st.load(CFG)
    a1 = analyze(w1, CFG)
    assert list(w1.uids.values()) == list(range(1, 12828)) and not w1.issues
    c0, c1 = Counter(i.code for i in a0.issues), Counter(i.code for i in a1.issues)
    assert {k: c0[k] for k in KEEP} == {k: c1[k] for k in KEEP}
    assert not set(c1) - set(KEEP)                               # 나머지(C01·C02·C07·C11·C12)는 들여오며 풀렸다
    s0, s1 = a0.schedule, a1.schedule
    assert s0.val == s1.val and s0.critical == s1.critical and s0.near == s1.near and s0.float_ == s1.float_ and s0.gaps == s1.gaps
    assert a0.workload.month_totals() == a1.workload.month_totals()
    assert st.recompute(CFG) == []                               # 고치지 않으면 아무것도 안 움직인다
    # 다시 내보낸 엑셀을 읽어도 같은 일정
    w2 = read_wbs(d / "out.xlsx", CFG)
    a2 = analyze(w2, CFG)
    assert a2.schedule.val == s1.val and a2.schedule.critical == s1.critical
    assert not {i.code for i in w2.issues} - {"P09"}


def test_forward_matches_reschedule_when_pushing(synth_store):
    """기간을 늘리거나 시작을 늦추면 v0.4 다시 짜기(reschedule)와 같은 날짜가 나와야 한다."""
    d, info, st, rep = synth_store
    m = st.model(CFG)
    tasks = m.wbs.tasks
    val0, _ = event_values(tasks)
    crit = sorted(no for no in analyze(m.wbs, CFG).schedule.critical if not tasks[no].children and not tasks[no].is_milestone)
    picks = [crit[len(crit) // 5], crit[len(crit) // 2], crit[-len(crit) // 4]]
    work = dict(m.work)
    dur = {}
    for no in picks:
        work[no] = dur[no] = m.work[no] + 7
    late = next(no for no in m.floor if no not in picks and not tasks[no].is_milestone)
    floor = dict(m.floor)
    floor[late] = m.floor[late] + 30
    got, _ = forward(tasks, m.edges, val0, floor, work, m.relax)
    want = reschedule(tasks, m.edges, val0, dur=dur, start={late: floor[late]})
    moved = [k for k in want if abs(want[k] - val0[k]) > EPS]
    assert len(moved) > 100
    assert [k for k in want if abs(want[k] - got[k]) > EPS] == []


def styled_small(path):
    make_small(path)
    wb = load_workbook(path)
    ws = wb.active
    ws.cell(3, 1).font = Font(bold=True)
    ws.column_dimensions["C"].width = 44
    for r in range(4, ws.max_row + 1):
        if "." not in str(ws.cell(r, 2).value):
            for c in range(1, 14):
                ws.cell(r, c).fill = PatternFill("solid", fgColor="FFDDEEFF")
    wb.save(path)
    return path


def test_small_format_and_columns(tmp_path):
    src = styled_small(tmp_path / "small.xlsx")
    with Store(tmp_path / "a.db") as st:
        rep = st.import_xlsx(src, CFG)
        assert rep.moved == 0 and rep.person_columns == 1 and rep.cleared_top_cells == 1
        assert list(rep.dropped_columns) == ["비고"] and rep.dropped_columns["비고"] == rep.tasks
        with pytest.raises(ValueError):
            st.import_xlsx(src, CFG)                             # 들어 있는 DB 에 또 들여오지 않는다
        out = st.export_xlsx(tmp_path / "a.xlsx", CFG)
        a, b = load_workbook(src).active, load_workbook(out).active
        assert b.cell(1, 1).value is None                        # 헤더 위 글자는 지운다
        assert [c.value for c in b[3]][:13] == [c.value for c in a[3]][:13]
        assert b.cell(3, 1).font.bold and b.column_dimensions["C"].width == 44
        cols = st.meta("columns")
        noted = {(n["src_row"], cols[n["col"]]) for n in st.db.execute(
            "SELECT t.src_row, n.col FROM note n JOIN task t ON t.uid = n.task_uid")}
        # 원본에 비어 있던 묶음 행의 기간·시작·완료·MH 는 하위에서 채워 나간다 — 그것만 기록에 있다
        assert noted and all("." not in str(a.cell(r, 2).value) or a.cell(r, 4).value is None for r, _ in noted)
        assert {c for _, c in noted} <= {4, 5, 6, 11}
        for r in range(4, a.max_row + 1):
            assert b.cell(r, 2).fill.fgColor.rgb == a.cell(r, 2).fill.fgColor.rgb, r
            assert b.cell(r, 9).value is None and b.cell(r, 13).value is None       # 작성자·비고
            for c in (1, 2, 3, 4, 5, 6, 7, 10, 11, 12):
                assert (r, c) in noted or same_cell(a.cell(r, c).value, b.cell(r, c).value), (r, c)
        blob = st.meta("template")
        assert "가나다".encode() not in blob and "메모".encode() not in blob
        # 기간을 늘리면 뒤 업무가 밀리고, 날짜·기간 글자 모양은 그대로
        m = st.model(CFG)
        no = next(t.no for t in m.wbs.tasks.values() if not t.children and t.duration_days and any(e.pred == t.no for e in m.edges))
        uid = m.rows[no]["uid"]
        st.db.execute("UPDATE task SET duration_days = duration_days + 3, work_days = work_days + 3 WHERE uid = ?", (uid,))
        changed = st.recompute(CFG)
        assert uid in changed and len(changed) > 1
        st.export_xlsx(tmp_path / "b.xlsx", CFG)
        c = load_workbook(tmp_path / "b.xlsx").active
        r = 3 + no
        old_f, new_f = parse_date(b.cell(r, 6).value)[0], parse_date(c.cell(r, 6).value)
        assert isinstance(c.cell(r, 6).value, str) and new_f[0] > old_f and new_f[1] is not None
        assert new_f[0].weekday() < 5 and c.cell(r, 5).value == b.cell(r, 5).value
        assert parse_duration(c.cell(r, 4).value, UNITS)[0] == parse_duration(b.cell(r, 4).value, UNITS)[0] + 3
        assert c.cell(r, 4).value.endswith(" 일")
        w = read_wbs(tmp_path / "b.xlsx", CFG)
        assert not [i for i in analyze(w, CFG).issues if i.code in ("C05", "C11")]
    cfg = copy.deepcopy(CFG)
    cfg["keep_columns"] = ["비고"]
    with Store(tmp_path / "k.db") as st:
        rep = st.import_xlsx(src, cfg)
        assert rep.kept_columns == ["비고"] and not rep.dropped_columns
        k = load_workbook(st.export_xlsx(tmp_path / "k.xlsx", cfg)).active
        assert k.cell(4, 13).value == "메모" and k.cell(4, 9).value is None


def test_cli_import_export(tmp_path, capsys):
    src = make_small(tmp_path / "small.xlsx")
    db = tmp_path / "db" / "wbs.db"
    assert main(["import", str(src), "--db", str(db)]) == 0
    assert db.exists() and (db.parent / "wbs.import_notes.json").exists()
    assert main(["import", str(src), "--db", str(db)]) == 2      # 덮어쓰지 않는다
    assert main(["import", str(src), "--db", str(db), "--force"]) == 0
    assert len(list(db.parent.glob("wbs.db.bak-*"))) == 1         # 옛 DB 는 이름을 바꿔 남긴다
    out = tmp_path / "x.xlsx"
    assert main(["export", "--db", str(db), "--out", str(out)]) == 0
    assert out.exists() and "가나다" not in capsys.readouterr().out
    assert main(["export", "--db", str(tmp_path / "none.db")]) == 2
    assert date.today().isoformat()[:4] in Store(db).meta("imported_at")
