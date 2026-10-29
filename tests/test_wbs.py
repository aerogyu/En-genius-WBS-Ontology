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
    assert wb.sheetnames == ["요약", "검사", "업무", "분야 행렬", "분야 간 연결", "가장 긴 사슬"]
    text = html.read_text(encoding="utf-8")
    assert "가나다" not in text and "<script src" not in text and "http" not in text.split("<script>")[1].replace("http://www.w3.org/2000/svg", "")


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
