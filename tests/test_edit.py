"""칸 고치기 저장 (v0.5 S3) — 권한 · 겹침 거절 · 이력 · 날짜 다시 짜기 · 내보낸 엑셀에 반영."""
from __future__ import annotations

import json
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from engenius_wbs import auth
from engenius_wbs.analyze import analyze
from engenius_wbs.edit import EditError, apply_edits, history
from engenius_wbs.parse import parse_date, parse_duration, read_wbs
from engenius_wbs.store import Store
from engenius_wbs.synth import date_text, workday
from test_wbs import CFG, SMALL, UNITS

PW = "test-pass-1234"


def make_two(path):
    """작은 WBS 두 모듈 (Fan 1~12 · HPC 13~24). HPC 의 첫 업무는 Fan 의 E(12) 뒤."""
    wb = Workbook()
    ws = wb.active
    ws.append(["번호", "WBS", "작업 이름", "기간", "시작", "완료", "선행 작업", "후속 작업", "투입인력(명)", "Man Hours", "산출물"])
    for k, (mod, off) in enumerate((("Fan", 0), ("HPC", 12))):
        for no, code, name, dur, es, preds, succs, deliv in SMALL:
            lv = code.count(".") + 1
            if off:
                es = None if es is None else es + 16
                preds = ",".join(str(int(p.split("F")[0].split("S")[0]) + off) + p[len(p.split("F")[0].split("S")[0]):]
                                 for p in preds.split(",")) if preds else ""
                if no == 6:
                    preds = "12"
            st = date_text(workday(es)) if es is not None else None
            fi = date_text(workday(es + dur - 1)) if es is not None else None
            ws.append([no + off, str(k + 1) + code[1:], "   " * (lv - 1) + (mod if no == 1 else name), f"{dur} 일" if dur else None, st, fi,
                       preds or None, None, 2 if dur else None, dur * 16 if dur else None, deliv or None])
    wb.save(path)
    return path


@pytest.fixture()
def two(tmp_path):
    db = tmp_path / "wbs.db"
    st = Store(db)
    rep = st.import_xlsx(make_two(tmp_path / "two.xlsx"), CFG)
    assert rep.moved == 0 and rep.tasks == 24
    roots = [r[0] for r in st.db.execute("SELECT uid FROM task WHERE parent_uid IS NULL ORDER BY pos")]
    roles = {"admin": auth.add_role(st.db, "admin1", "관리자 1", "admin", PW),
             "fan": auth.add_role(st.db, "fan", "FAN 리더", "leader", PW, roots[0]),
             "hpc": auth.add_role(st.db, "hpc", "HPC 리더", "leader", PW, roots[1]),
             "view": auth.add_role(st.db, "view", "조회", "viewer", PW)}
    yield st, roles, tmp_path
    st.close()


def uid(st, no):
    return st.model(CFG).wbs.uids[no]


def dates(st):
    m = st.model(CFG)
    return {no: (t.start, t.finish) for no, t in m.wbs.tasks.items()}


def test_leader_saves_own_module_and_dates_push(two):
    st, R, _ = two
    before = dates(st)
    res = apply_edits(st, CFG, R["fan"], 1, [{"uid": uid(st, 7), "duration": 8}])      # B: 3 → 8 평일
    assert res["rev"] == 2 and res["tasks"] == 1 and res["moved"] > 1 and res["end_after"] > res["end_before"]
    after = dates(st)
    assert after[7][0] == before[7][0] and after[7][1] > before[7][1]
    assert after[10][0] > before[10][0] and after[13][0] > before[13][0]           # 뒤 업무·다른 모듈(HPC)도 밀린다 — 막지 않는다
    assert after[6] == before[6] and after[11] == before[11]                        # 앞·옆 업무는 그대로
    h = history(st, CFG)
    assert h[0]["role"] == "FAN 리더" and h[0]["changes"][0]["field"] == "duration" and h[0]["changes"][0]["old"] == "3"
    assert h[0]["note"]["moved"] == res["moved"] and history(st, CFG, uid(st, 7))[0]["txn"] == 2
    assert st.recompute(CFG) == []


def test_saved_change_reaches_exported_excel(two):
    st, R, tmp = two
    apply_edits(st, CFG, R["admin"], 1, [{"uid": uid(st, 7), "duration": 8, "man_hours": 100}, {"uid": uid(st, 10), "name": "C 고침"}])
    out = st.export_xlsx(tmp / "out.xlsx", CFG)
    ws = load_workbook(out).active
    row = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
    assert parse_duration(ws.cell(row[7], 4).value, UNITS)[0] == 8 and ws.cell(row[7], 10).value == 100
    assert ws.cell(row[10], 3).value.strip() == "C 고침"
    w = read_wbs(out, CFG)
    m = st.model(CFG)
    assert {no: (t.start, t.finish) for no, t in w.tasks.items()} == {no: (t.start, t.finish) for no, t in m.wbs.tasks.items()}
    assert not [i for i in analyze(w, CFG).issues if i.code in ("C05", "C11", "C12")]


def test_permissions(two):
    st, R, _ = two
    for who, no in (("view", 7), ("fan", 19), ("hpc", 7), ("fan", 1)):          # 조회자 · 남의 모듈 · Lv1 줄
        with pytest.raises(EditError) as e:
            apply_edits(st, CFG, R[who], 1, [{"uid": uid(st, no), "name": "x"}])
        assert e.value.kind == "denied"
    assert st.db.execute("SELECT COUNT(*) FROM txn").fetchone()[0] == 1          # 아무것도 안 씀
    apply_edits(st, CFG, R["hpc"], 1, [{"uid": uid(st, 19), "duration": 9}])
    apply_edits(st, CFG, R["admin"], 2, [{"uid": uid(st, 1), "name": "Fan 모듈"}])
    # 내 업무의 선행으로 다른 모듈 업무를 고르는 것은 된다
    apply_edits(st, CFG, R["hpc"], 3, [{"uid": uid(st, 18), "preds": "6,7"}])
    assert [no for no in [t.pred for t in st.model(CFG).wbs.tasks[18].deps]] == [6, 7]


def test_conflict(two):
    st, R, _ = two
    apply_edits(st, CFG, R["fan"], 1, [{"uid": uid(st, 7), "duration": 6}])
    with pytest.raises(EditError) as e:                                           # 판 1 을 보고 같은 업무를 고침 → 거절
        apply_edits(st, CFG, R["admin"], 1, [{"uid": uid(st, 7), "name": "B2"}, {"uid": uid(st, 10), "name": "C2"}])
    assert e.value.kind == "conflict" and e.value.items[0]["role"] == "FAN 리더" and e.value.items[0]["no"] == 7
    assert st.model(CFG).wbs.tasks[10].name == "C"                               # 묶음 전체가 안 들어감
    apply_edits(st, CFG, R["admin"], 1, [{"uid": uid(st, 10), "name": "C2"}])    # 다른 업무만이면 판 1 을 봤어도 된다
    apply_edits(st, CFG, R["admin"], 3, [{"uid": uid(st, 7), "name": "B2"}])     # 새 판을 보고 다시 하면 된다


def test_preds_rules(two):
    st, R, _ = two
    before = dates(st)
    # D(11) 의 선행 6SS 를 지움 — 시작 하한이 없던 업무라도 저절로 당겨지지 않는다
    r11 = st.db.execute("SELECT start_floor FROM task WHERE uid = ?", (uid(st, 11),)).fetchone()[0]
    apply_edits(st, CFG, R["fan"], 1, [{"uid": uid(st, 11), "preds": ""}])
    assert dates(st)[11] == before[11]
    if r11 is None:
        assert st.db.execute("SELECT start_floor FROM task WHERE uid = ?", (uid(st, 11),)).fetchone()[0] == before[11][0].isoformat()
    # E(12) 에 선행 7FS+10일 을 더함 → 밀린다, 표기는 그대로 내보낸다
    apply_edits(st, CFG, R["fan"], 2, [{"uid": uid(st, 12), "preds": "10,11FF,7FS+10일"}])
    assert dates(st)[12][0] > before[12][0]
    deps = st.db.execute("SELECT suffix FROM dep WHERE succ_uid = ? ORDER BY ord", (uid(st, 12),)).fetchall()
    assert [d[0] for d in deps] == ["", "FF", "FS+10일"]
    # 시작 하한 '이 날 이후' — 늦추면 밀리고, 비우지 않는 한 남는다
    late = before[6][0].replace(day=1).replace(month=before[6][0].month % 12 + 1)
    apply_edits(st, CFG, R["fan"], 3, [{"uid": uid(st, 6), "start": late.isoformat()}])
    assert dates(st)[6][0] >= late


def test_invalid(two):
    st, R, _ = two
    cases = [{"uid": uid(st, 5), "duration": 3},                 # 요약 업무 기간
             {"uid": uid(st, 7), "preds": "6,xx"},               # 못 읽는 선행
             {"uid": uid(st, 7), "preds": "7"},                  # 자기 자신
             {"uid": uid(st, 7), "preds": "999"},                # 없는 번호
             {"uid": uid(st, 7), "duration": -1},
             {"uid": uid(st, 7), "start": "2027-13-40"},
             {"uid": uid(st, 7), "name": "  "},
             {"uid": uid(st, 7), "duration": 3}]                 # 바뀐 것 없음
    for c in cases:
        with pytest.raises(EditError) as e:
            apply_edits(st, CFG, R["admin"], 1, [c])
        assert e.value.kind == "invalid", c
    with pytest.raises(EditError):
        apply_edits(st, CFG, R["admin"], 1, [{"uid": 99999, "name": "x"}])
    assert st.db.execute("SELECT COUNT(*) FROM txn").fetchone()[0] == 1


# --- 서버 API ---------------------------------------------------------------------------------------
pytest.importorskip("starlette")
pytest.importorskip("httpx")


def test_save_api(two):
    from starlette.testclient import TestClient

    from engenius_wbs.server import create_app
    st, R, tmp = two
    app = create_app(st.path)

    def client(who):
        c = TestClient(app)
        assert c.post("/login", data={"login": who, "password": PW, "mode": "role"}, follow_redirects=False).status_code == 303
        return c
    fan, hpc, view = client("fan"), client("hpc"), client("view")
    page = fan.get("/").text
    assert '"kind": "leader"' in page and "forwardJS" in page
    d = json.loads(fan.get("/data.js").text[len("window.__WBS_DATA__="):-1])
    assert len(d["uid"]) == len(d["tasks"]) == 24 and len(d["fl"]) == 24 and all(len(e) == 7 for e in d["edges"])
    u7 = d["uid"][[t[0] for t in d["tasks"]].index(7)]
    body = {"base_rev": 1, "edits": [{"uid": u7, "duration": 8}]}
    assert view.post("/api/save", json=body).status_code == 403
    assert hpc.post("/api/save", json=body).status_code == 403
    assert fan.post("/api/save", data="x", headers={"content-type": "text/plain"}).status_code == 415
    r = fan.post("/api/save", json=body)
    assert r.status_code == 200 and r.json()["rev"] == 2
    r = fan.post("/api/save", json={"base_rev": 1, "edits": [{"uid": u7, "name": "B3"}]})
    assert r.status_code == 409 and r.json()["items"][0]["role"] == "FAN 리더"
    assert view.get("/api/status").json()["rev"] == 2                       # 다른 사람 화면: '새 판이 있음'
    h = view.get("/api/history").json()
    assert h[0]["role"] == "FAN 리더" and h[0]["changes"][0]["new"] == "8"
    x = view.get("/export.xlsx")
    ws = load_workbook(BytesIO(x.content)).active
    r7 = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == 7)
    assert parse_duration(ws.cell(r7, 4).value, UNITS)[0] == 8                 # 저장한 것은 내려받은 엑셀에 들어간다
