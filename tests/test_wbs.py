"""WBS 읽기·검사·CPM·출력 — 손으로 셈한 작은 WBS + 가짜 12,827행 (심은 문제를 모두 찾는지)."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from rdflib import Graph

from engenius_wbs.analyze import analyze, weekdays_between
from engenius_wbs.config import load_config
from engenius_wbs.parse import parse_date, parse_deps, parse_duration, read_wbs
from engenius_wbs.rdfout import build_graphs
from engenius_wbs.report import write_report
from engenius_wbs.analyze import Edge
from engenius_wbs.apply import apply_changes, sha256_of
from engenius_wbs.schedule import date_schedule, event_values, finish_date, reschedule, start_date, wdate, widx
from engenius_wbs.synth import date_text, generate, workday
from engenius_wbs.viz import write_html

ROOT = Path(__file__).resolve().parents[1]
CFG = load_config(ROOT / "config" / "wbs.yaml")
UNITS = {k.lower(): float(v) for k, v in CFG["unit_days"].items()}


@pytest.mark.parametrize("text,want", [
    ("27-01-01 (금)", (date(2027, 1, 1), "금")),
    ("39-12-02(금)", (date(2039, 12, 2), "금")),
    ("2028/3/7", (date(2028, 3, 7), "")),
    ("", (None, "")),
    ("미정", (None, None)),
])
def test_parse_date(text, want):
    assert parse_date(text) == want


@pytest.mark.parametrize("text,want", [
    ("120 일", (120.0, False, True)), ("<120 일 >", (120.0, False, True)), ("12일?", (12.0, True, True)),
    ("3주", (15.0, False, True)), ("0 일", (0.0, False, True)), ("2 fortnights", (None, False, False)), (None, (None, False, True)),
])
def test_parse_duration(text, want):
    assert parse_duration(text, UNITS) == want


def test_parse_deps():
    deps, bad = parse_deps("12, 15FS+10일,20FF,7SS-2일,9완시,3FS+50%, 약 4", UNITS)
    assert [(d.pred, d.type, d.lag_days, d.lag_text) for d in deps] == [
        (12, "FS", 0.0, ""), (15, "FS", 10.0, "+10일"), (20, "FF", 0.0, ""), (7, "SS", -2.0, "-2일"), (9, "FS", 0.0, ""),
        (3, "FS", None, "+50%")]
    assert bad == ["약 4"]


def test_weekdays_between():
    assert weekdays_between(date(2027, 1, 1), date(2027, 1, 7)) == 5      # 금~목
    assert weekdays_between(date(2027, 1, 2), date(2027, 1, 3)) == 0


# 작은 WBS: (번호, WBS, 이름, 기간, 근무일 시작, 선행, 후속, 산출물)
SMALL = [
    (1, "1", "Fan", None, None, "", "", ""),
    (2, "1.1", "설계통합", None, None, "", "", ""),
    (3, "1.1.1", "기본설계", None, None, "", "", ""),
    (4, "1.1.1.1", "성능", None, None, "", "", ""),
    (5, "1.1.1.1.1", "성능 해석", None, None, "", "", "성능 보고서, 검토 자료"),
    (6, "1.1.1.1.1.1", "A", 5, 0, "", "7,11", ""),
    (7, "1.1.1.1.1.2", "B", 3, 5, "6", "10", ""),
    (8, "1.1.1.2", "이차유로", None, None, "", "", ""),
    (9, "1.1.1.2.1", "유로 해석", None, None, "", "", ""),
    (10, "1.1.1.2.1.1", "C", 4, 10, "7FS+2일", "12", ""),
    (11, "1.1.1.2.1.2", "D", 10, 0, "6SS", "12", ""),
    (12, "1.1.1.2.1.3", "E", 2, 14, "10,11FF", "", ""),
]


def make_small(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.append(["가상 사업 WBS (시험용)"])
    ws.append([])
    ws.append(["번호", "WBS", "작업 이름", "기간", "시작", "완료", "선행 작업", "후속 작업", "작성자", "투입인력(명)", "Man Hours", "산출물", "비고"])
    for no, code, name, dur, es, preds, succs, deliv in SMALL:
        lv = code.count(".") + 1
        st = date_text(workday(es)) if es is not None else None
        fi = date_text(workday(es + dur - 1)) if es is not None else None
        ws.append([no, code, "   " * (lv - 1) + name, f"{dur} 일" if dur else None, st, fi, preds or None, succs or None,
                   "가나다", 2 if dur else None, dur * 16 if dur else None, deliv or None, "메모"])
    wb.save(path)
    return path


@pytest.fixture(scope="module")
def small(tmp_path_factory):
    p = make_small(tmp_path_factory.mktemp("small") / "small.xlsx")
    w = read_wbs(p, CFG)
    return w, analyze(w)


def test_small_read(small):
    w, a = small
    assert len(w.tasks) == 12 and not w.issues and not a.issues
    assert "작성자" in w.skipped_columns and "비고" in w.skipped_columns and "staff" in w.columns
    e = w.tasks[12]
    assert (e.level, e.parent, e.module, e.phase, e.discipline) == (6, 9, "Fan", "기본설계", "이차유로")
    assert w.tasks[5].deliverables == ["성능 보고서", "검토 자료"]
    assert w.tasks[10].deps[0].lag_days == 2.0 and w.tasks[11].deps[0].type == "SS"
    assert w.tasks[6].start == date(2027, 1, 1) and w.tasks[6].finish == date(2027, 1, 7)


def test_small_cpm_and_counts(small):
    w, a = small
    assert a.project_days == 16
    assert a.critical_chain == [6, 7, 10, 12]
    assert a.total_float[11] == 6 and a.total_float[10] == 0
    assert a.es[12] == 14 and a.es[11] == 0
    assert a.up_count[12] == 4 and a.down_count[6] == 4 and a.up_count[6] == 0
    assert a.discipline_flows == {("성능", "성능"): 1, ("성능", "이차유로"): 2, ("이차유로", "이차유로"): 2}


def test_widx_roundtrip():
    assert widx(date(2000, 1, 3)) == 0 and widx(date(2000, 1, 8)) == 5 and widx(date(2000, 1, 10)) == 5   # 토요일 = 다음 월요일
    d = date(2026, 12, 1)
    for _ in range(400):
        if d.weekday() < 5:
            assert wdate(widx(d)) == d and start_date(widx(d)) == d and finish_date(widx(d) + 1) == d
        d = date.fromordinal(d.toordinal() + 1)
    assert widx(date(2027, 1, 11)) - widx(date(2027, 1, 1)) == 6                                        # 금 → 다음다음 월: 평일 6


def test_small_date_schedule(small):
    """A(0~4) → B(5~7) → FS+2 → C(10~13) → E(14~15) 가 주경로. D(SS A, 0~9) 는 E 와 FF — E 완료(16) − D 완료(10) = 여유 6."""
    w, a = small
    sc = a.schedule
    assert sc.end - sc.start == 16 and sc.start_date == date(2027, 1, 1) and sc.end_date == w.tasks[12].finish
    assert sc.critical == [6, 7, 10, 12] and sc.near == [] and sc.next_float == 6
    assert sc.float_[11] == 6 and sc.float_[9] == 0 and sc.float_[4] == 0 and sc.float_[1] == 0
    assert {(e.pred, e.succ): g for e, g in zip(a.edges, sc.gaps)} == {(6, 7): 0, (7, 10): 0, (6, 11): 0, (10, 12): 0, (11, 12): 6}
    sc2 = date_schedule(w.tasks, a.edges, 0, 6)                                                          # 준주경로 기준을 넓히면 D 가 들어온다
    assert sc2.near == [11] and sc2.critical == sc.critical
    sc3 = date_schedule(w.tasks, a.edges, 6, 6)
    assert sc3.critical == [6, 11, 7, 10, 12] and sc3.next_float is None


def test_reschedule_rules(small):
    """'간격 유지' 다시 짜기 — 고치지 않으면 그대로 · 구동된 업무는 따라 당겨지고 밀림 · 여유가 있으면 여유를 다 쓸 때까지 그대로."""
    w, a = small
    T, val0 = w.tasks, event_values(w.tasks)[0]
    b = val0[2 * 6]                                                                  # A 시작 = 평일 0
    assert reschedule(T, a.edges, val0) == val0
    v = reschedule(T, a.edges, val0, dur={7: 1.0})                                   # B 3 → 1: C·E 2 평일 당겨짐, D 그대로
    assert (v[2 * 10] - b, v[2 * 12] - b, v[2 * 12 + 1] - b, v[2 * 11] - b) == (8, 12, 14, 0)
    sc = date_schedule(T, a.edges, 0, 5, val=v)
    assert sc.end - sc.start == 14 and sc.float_[11] == 4 and sc.critical == [6, 7, 10, 12]
    v = reschedule(T, a.edges, val0, dur={11: 18.0})                                 # D 10 → 18: FF 여유 6 을 넘어 E 를 2 밂
    assert (v[2 * 12] - b, v[2 * 12 + 1] - b) == (16, 18)
    assert date_schedule(T, a.edges, 0, 5, val=v).critical == [6, 11, 12]
    v = reschedule(T, a.edges, val0, start={11: b + 3})                                  # D '이 날 이후' +3 → D 만 밀림 (E 는 C 가 쥠)
    assert (v[2 * 11] - b, v[2 * 12] - b) == (3, 14)
    only_ff = [e for e in a.edges if (e.pred, e.succ) != (10, 12)]                   # E 의 선행을 D(FF, 여유 6) 하나로 → 구동 안 됨
    v = reschedule(T, only_ff, val0, dur={11: 8.0})
    assert v[2 * 12] - b == 14                                                           # 여유 안에서 줄어든 것은 E 를 당기지 않음
    orig = {(e.pred, e.succ, e.type, e.lag_days or 0.0) for e in a.edges}
    v = reschedule(T, a.edges + [Edge(10, 11, "FS", 0.0, "")], val0, orig=orig)      # 새 선행 C → D: D 가 C 뒤로 밀리고 FF 로 E 도
    assert (v[2 * 11] - b, v[2 * 11 + 1] - b, v[2 * 12] - b, v[2 * 12 + 1] - b) == (14, 24, 22, 24)
    v = reschedule(T, a.edges + [Edge(6, 12, "FS", 0.0, "")], val0, orig=orig)       # 이미 지켜지는 새 선행(A → E)은 아무것도 안 바꿈
    assert v == val0


def test_apply_roundtrip(small, tmp_path):
    import json
    import shutil
    w, a = small
    src = tmp_path / "small.xlsx"
    shutil.copyfile(w.source, src)
    sha = sha256_of(src)
    ch = {"format": "engenius-wbs-changes/1", "source": "small.xlsx", "sheet": w.sheet, "sha256": sha,
          "changes": [{"no": 7, "field": "duration", "from": 3, "to": 1},
                      {"no": 12, "field": "preds", "from": "10,11FF", "to": "10"},
                      {"no": 11, "field": "name", "from": "D", "to": "D 바꿈"},
                      {"no": 11, "field": "man_hours", "from": 160, "to": 100}],
          "result": {"end": "2027-01-20", "moved": {"7": ["2027-01-08", "2027-01-08"], "10": ["2027-01-13", "2027-01-18"],
                                                     "12": ["2027-01-19", "2027-01-20"]}}}   # 요약은 엑셀 날짜가 비어 있어 안 씀
    (tmp_path / "ch.json").write_text(json.dumps(ch, ensure_ascii=False), encoding="utf-8")
    r = apply_changes(src, tmp_path / "ch.json", CFG)
    assert r.out.name == "small_수정_v01.xlsx" and sha256_of(src) == sha and not r.mismatches, r.mismatches
    assert (r.end_before, r.end_after) == (date(2027, 1, 22), date(2027, 1, 20)) and r.out.with_suffix(".changes.json").exists()
    w2 = read_wbs(r.out, CFG)
    t7, t12, t11, t10 = w2.tasks[7], w2.tasks[12], w2.tasks[11], w2.tasks[10]
    assert (t7.duration_text, t7.start, t7.finish) == ("1 일", date(2027, 1, 8), date(2027, 1, 8))
    assert [d.pred for d in t12.deps] == [10] and (t11.name, t11.man_hours) == ("D 바꿈", 100)
    assert w2.tasks[1].finish is None and not r.kept_summaries
    ws = load_workbook(r.out, read_only=True).worksheets[0]
    cells = {row[0]: row for row in ws.iter_rows(min_row=4, values_only=True)}
    assert cells[10][4] == "27-01-13 (수)" and cells[11][7] is None and cells[7][7] == "10"     # 날짜 글자 모양 · 후속 열(H) 다시 씀
    assert cells[11][2].startswith("               D 바꿈") and cells[11][8] == "가나다"      # 들여쓰기 유지 · 다른 칸 그대로
    a2 = analyze(w2)
    assert not a2.issues and a2.schedule.end_date == date(2027, 1, 20)
    assert apply_changes(src, tmp_path / "ch.json", CFG).out.name == "small_수정_v02.xlsx"           # 덮어쓰지 않음
    ch["result"]["end"] = "2027-01-21"
    (tmp_path / "bad.json").write_text(json.dumps(ch, ensure_ascii=False), encoding="utf-8")
    (tmp_path / "o").mkdir()
    assert any("사업 완료" in m for m in apply_changes(src, tmp_path / "bad.json", CFG, tmp_path / "o").mismatches)   # 보기 결과와 다르면 알림
    ch["sha256"] = "0" * 64
    (tmp_path / "other.json").write_text(json.dumps(ch, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ValueError):
        apply_changes(src, tmp_path / "other.json", CFG)


def test_small_outputs(small, tmp_path):
    w, a = small
    g, d = build_graphs(w, a, CFG, "WRUN-TEST")
    q = """PREFIX wbs: <http://en-genius.local/onto/wbs/>
           SELECT (COUNT(*) AS ?n) WHERE { ?t a wbs:Task ; wbs:precededBy ?p . ?p a wbs:Task }"""
    assert int(next(iter(g.query(q)))[0]) == 5
    assert int(next(iter(d.query("PREFIX wbs: <http://en-genius.local/onto/wbs/> SELECT (COUNT(*) AS ?n) WHERE { ?t wbs:onLongestChain true }")))[0]) == 4
    rep = write_report(w, a, CFG, tmp_path / "r.xlsx")
    html = write_html(w, a, tmp_path / "v.html")
    wb = load_workbook(rep, read_only=True)
    assert wb.sheetnames == ["요약", "검사", "주경로", "단계 요약", "분야 요약", "월별 공수", "업무", "분야 행렬", "분야 간 연결"]
    cp = list(wb["주경로"].iter_rows(values_only=True))
    assert [r[2] for r in cp[1:]] == [6, 7, 10, 12] and cp[2][13] == "6" and cp[3][13] == "7+2일" and cp[4][13] == "10"   # 구동 선행
    text = html.read_text(encoding="utf-8")
    assert "가나다" not in text and "<script src" not in text and "http" not in text.split("<script>")[1].replace("http://www.w3.org/2000/svg", "")
    for piece in ('id="tree"', 'id="filters"', 'id="grain"', 'id="lane"', 'id="quick"', 'data-v="time"', 'data-v="flow"', 'id="side"',
                  'id="copyLink"', 'data-v="gantt"', 'id="lines"', 'id="gantt"', '"source":"small.xlsx"',
                  'data-t="mh"', 'id="mhChart"', 'id="mhComp"', 'id="msk"', '"mh":{"months":',
                  'data-t="cp"', 'id="cpList"', 'id="cpGrp"', 'data-q="cp"', '"cp":{"crit_days":0',
                  'id="edChip"', 'function reschedule(', 'function edDownload(', '"sha256":"', '"units":{',
                  'id="tgTree"', 'id="tgSide"', 'function emh(', 'function mhOf('):   # 양옆 켜고 끄기 · 업무별 MH (보기 v3.3)
        assert piece in text, piece


def test_header_missing(tmp_path):
    wb = Workbook()
    wb.active.append(["이름", "날짜"])
    wb.save(tmp_path / "x.xlsx")
    w = read_wbs(tmp_path / "x.xlsx", CFG)
    assert not w.tasks and w.issues[0].code == "P01"


@pytest.fixture(scope="module")
def synthetic(tmp_path_factory):
    d = tmp_path_factory.mktemp("synth")
    info = generate(d / "wbs.xlsx", 12827, 7)
    w = read_wbs(d / "wbs.xlsx", CFG)
    return d, info, w, analyze(w)


def test_synthetic_finds_every_injected_issue(synthetic):
    d, info, w, a = synthetic
    assert info["rows"] == len(w.tasks) == 12827
    found = {(i.code, i.no) for i in w.issues + a.issues}
    for code, nos in info["injected"].items():
        for no in nos:
            assert (code, no) in found, (code, no)
    assert {c for c, _ in found} == set(info["injected"]), "심지 않은 문제가 나왔다"
    for code, nos in info["injected"].items():
        assert sum(1 for c, _ in found if c == code) == len(nos), code
    assert len(a.cycles) == 1 and len(a.cycles[0]) == 3


def test_synthetic_critical_path(synthetic):
    d, info, w, a = synthetic
    sc = a.schedule
    assert set(a.critical_chain) <= set(sc.critical)                     # 합성은 날짜 = 선행 계산 → 기간 합 최장 사슬은 주경로 안에
    assert sc.near and sc.next_float is not None and 0 < sc.next_float <= 5   # 나란한 두 업무 → 준주경로가 생긴다
    assert sc.excluded == info["undated"]
    last = max(sc.critical, key=lambda n: sc.val[2 * n + 1])
    assert sc.val[2 * last + 1] == sc.end and sc.float_[last] == 0
    neg = {e.succ for e, g in zip(a.edges, sc.gaps) if g is not None and g < -1e-6}
    c05 = {i.no for i in a.issues if i.code == "C05"}
    c10 = {i.no for i in a.issues if i.code == "C10"}
    assert neg <= c05 | c10                                              # 음수 간격 = 날짜가 선행을 어긴 것뿐
    assert all(sc.float_[n] >= -1e-6 for n in sc.float_ if not w.tasks[n].children and n not in c05 | c10)


def test_synthetic_milestones_and_workload(synthetic, tmp_path):
    d, info, w, a = synthetic
    ms = sorted(t.no for t in w.tasks.values() if t.is_milestone)
    assert ms == sorted(info["milestones"]) and len(ms) > 100
    wl = a.workload
    assert wl.undated == info["undated"] and len(wl.undated) == 1
    used = sum(w.tasks[n].man_hours for n in wl.contributors)
    assert abs(wl.total_mh + wl.undated_mh - used) < 1e-6                     # 달에 나눈 것 + 날짜 없음 = 쓴 MH 전부
    assert abs(sum(wl.month_totals().values()) - wl.total_mh) < 1e-6
    assert wl.by_level == {6: len(wl.contributors)}                           # 합성은 MH 가 Lv6 에 다 있다 → 요약 행은 안 씀
    assert wl.leaves_without_mh == len(ms)                                    # MH 없는 하위 업무 = 마일스톤뿐
    assert len(wl.months) == len(set(wl.months)) and wl.months == sorted(wl.months)
    g, _ = build_graphs(w, a, CFG, "WRUN-TEST")
    q = "PREFIX wbs: <http://en-genius.local/onto/wbs/> SELECT (COUNT(*) AS ?n) WHERE { ?t wbs:isMilestone true }"
    assert int(next(iter(g.query(q)))[0]) == len(ms)
    wb = load_workbook(write_report(w, a, CFG, tmp_path / "r.xlsx"), read_only=True)
    rows = list(wb["월별 공수"].iter_rows(values_only=True))
    assert rows[0][0] == "달" and rows[-2][0] == "합계" and rows[-1][0] == "날짜 없음"
    assert abs(rows[-2][-1] - round(wl.total_mh, 1)) < 0.05 and len(rows) == len(wl.months) + 3
    ph = list(wb["단계 요약"].iter_rows(values_only=True))
    c11 = info["injected"]["C11"][0]
    row = next(r for r in ph if r[0] == c11)
    assert row[10] == 7                                                       # 완료 차이(일) = 심은 7일


# 공수 규칙을 손으로 셈하는 작은 WBS: (WBS, 이름, 기간, 시작, 완료, Man Hours)
#   A(Lv2) 에 100 이 적혔지만 아래 a1·a2 에 60·30 → 아래 것만 (C12) · B(Lv2) 에만 40 → B 를 씀 (아래 b1·b2 는 MH 없음)
#   C(Lv2) 하위 없음·날짜 없음 → 날짜 없음 20 · M(Lv1) 날짜는 하위 범위보다 이름 (C11) · b2 = 마일스톤
WL = [
    ("1", "M", None, date(2027, 1, 1), date(2027, 3, 5), None),
    ("1.1", "A", None, None, None, 100),
    ("1.1.1", "a1", 10, date(2027, 1, 25), date(2027, 2, 5), 60),     # 평일 10일: 1월 5일 · 2월 5일
    ("1.1.2", "a2", 1, date(2027, 2, 1), date(2027, 2, 1), 30),
    ("1.2", "B", None, date(2027, 3, 1), date(2027, 3, 5), 40),
    ("1.2.1", "b1", 5, date(2027, 3, 1), date(2027, 3, 5), None),
    ("1.2.2", "b2", 0, date(2027, 3, 5), date(2027, 3, 5), None),
    ("1.3", "C", 3, None, None, 20),
    ("1.4", "D", 2, date(2027, 3, 6), date(2027, 3, 7), 8),              # 주말만 → 달력 날 수로 나눔
]


def test_workload_rules(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["번호", "WBS", "작업 이름", "기간", "시작", "완료", "Man Hours"])
    for i, (code, name, dur, st, fi, mh) in enumerate(WL, start=1):
        lv = code.count(".") + 1
        ws.append([i, code, "   " * (lv - 1) + name, None if dur is None else f"{dur} 일",
                   date_text(st) if st else None, date_text(fi) if fi else None, mh])
    wb.save(tmp_path / "wl.xlsx")
    w = read_wbs(tmp_path / "wl.xlsx", CFG)
    a = analyze(w)
    wl = a.workload
    assert wl.contributors == [3, 4, 5, 8, 9]
    assert wl.cells == {("2027-01", 1): 30.0, ("2027-02", 1): 60.0, ("2027-03", 1): 48.0}
    assert wl.months == ["2027-01", "2027-02", "2027-03"] and wl.undated == [8] and wl.undated_mh == 20
    assert wl.mh_below == {1: 158.0, 2: 90.0} and wl.leaves == 6 and wl.leaves_without_mh == 0
    assert wl.by_level == {2: 3, 3: 2}
    assert [t.no for t in w.tasks.values() if t.is_milestone] == [7]
    found = sorted((i.code, i.no) for i in a.issues)
    assert ("C12", 2) in found and ("C11", 1) in found and not any(c == "C11" and n != 1 for c, n in found)
    assert a.child_span[1] == (date(2027, 1, 25), date(2027, 3, 7)) and a.child_span[5] == (date(2027, 3, 1), date(2027, 3, 5))


def test_synthetic_no_person_names_in_outputs(synthetic, tmp_path):
    d, info, w, a = synthetic
    g, dg = build_graphs(w, a, CFG, "WRUN-TEST")
    assert "가나다" not in g.serialize(format="nt") and "가나다" not in dg.serialize(format="nt")
    assert "가나다" not in write_html(w, a, tmp_path / "v.html").read_text(encoding="utf-8")
    wb = load_workbook(write_report(w, a, CFG, tmp_path / "r.xlsx"), read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            assert not any(isinstance(v, str) and "가나다" in v for v in row)


def test_cli_build(tmp_path):
    src = make_small(tmp_path / "small.xlsx")
    env = dict(os.environ, PYTHONUTF8="1", PYTHONPATH=str(ROOT / "src"))
    r = subprocess.run([sys.executable, "-m", "engenius_wbs", "build", str(src), "--out", str(tmp_path / "out")],
                       cwd=ROOT, capture_output=True, text=True, encoding="utf-8", env=env)
    assert r.returncode == 0, r.stdout + r.stderr
    for name in ("report.xlsx", "wbs_view.html", "wbs.ttl", "derived.ttl", "run_log.json"):
        assert (tmp_path / "out" / name).exists(), name
    log = json.loads((tmp_path / "out" / "run_log.json").read_text(encoding="utf-8"))
    assert log["tasks"] == 12 and log["issues"] == {} and log["project_workdays"] == 16
    Graph().parse(tmp_path / "out" / "wbs.ttl")
    r = subprocess.run([sys.executable, "-m", "engenius_wbs", "build", str(tmp_path / "없음.xlsx")], cwd=ROOT,
                       capture_output=True, text=True, encoding="utf-8", env=env)
    assert r.returncode == 2
