"""S5 협업 (docs/05 14장) — 스키마 판 2 · 판 확정·비교 · 백업·복구 · 수정 요청 · 모듈별 내보내기 · 편집 열림/닫힘."""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from io import BytesIO

import pytest
from openpyxl import load_workbook

from engenius_wbs import backup, collab, snapshot
from engenius_wbs.__main__ import main
from engenius_wbs.edit import EditError, apply_edits, history
from engenius_wbs.rows import apply_row_op
from engenius_wbs.store import SCHEMA_VERSION, Store, module_of
from test_edit import PW, make_two, two  # noqa: F401  (fixture)
from test_wbs import CFG


def U(st, no):
    return st.model(CFG).wbs.uids[no]


def cells(path) -> list[tuple]:
    ws = load_workbook(path).worksheets[0]
    return [tuple(c.value for c in row) for row in ws.iter_rows()]


# --- 스키마 판 2 ------------------------------------------------------------------------------------
def test_schema_v1_db_is_upgraded(tmp_path):
    db = tmp_path / "wbs.db"
    with Store(db) as st:
        st.import_xlsx(make_two(tmp_path / "two.xlsx"), CFG)
        n = st.count()
    con = sqlite3.connect(db)                                    # S4 까지의 DB 처럼 되돌린다
    con.executescript("DROP TABLE snap_task; DROP TABLE snap_dep; DROP TABLE request; DROP TABLE snapshot;")
    con.execute("UPDATE meta SET value = '1' WHERE key = 'schema_version'")
    con.commit()
    con.close()
    with Store(db) as st:
        assert st.meta("schema_version") == SCHEMA_VERSION == 2 and st.count() == n
        assert st.db.execute("SELECT COUNT(*) FROM snapshot").fetchone()[0] == 0
        # 판 사본 표의 열 = 업무·선행 표의 열 (나중에 열을 더하면 둘 다 고칠 것)
        assert st.columns("snap_task")[1:] == st.columns("task")
        assert st.columns("snap_dep")[1:] == st.columns("dep")


# --- 판 확정 ----------------------------------------------------------------------------------------
def test_freeze_open_export_same(two):
    st, R, tmp = two
    before = st.export_xlsx(tmp / "before.xlsx", CFG)
    rev0 = st.rev()
    with pytest.raises(EditError):
        snapshot.freeze(st, CFG, R["fan"], "1차본")              # 관리자만
    s1 = snapshot.freeze(st, CFG, R["admin"], "1차본", "모듈 취합 뒤")
    assert s1["label"] == "1차본" and s1["rev"] == rev0 and s1["summary"]["tasks"] == 24 and s1["summary"]["end"]
    assert st.rev() == rev0                                      # 판 번호는 그대로 (일정이 안 바뀜)
    with pytest.raises(EditError):
        snapshot.freeze(st, CFG, R["admin"], " 1차본 ")          # 같은 이름
    # 고친 뒤에도 옛 판 엑셀 = 확정 때 내보낸 엑셀
    apply_edits(st, CFG, R["fan"], st.rev(), [{"uid": U(st, 7), "duration": 9, "name": "B 수정"}])
    mem = snapshot.open_snapshot(st, s1["id"])
    old = mem.export_xlsx(tmp / "snap.xlsx", CFG)
    mem.close()
    assert cells(old) == cells(before)
    assert cells(st.export_xlsx(tmp / "now.xlsx", CFG)) != cells(before)
    assert snapshot.find_snapshot(st, "1차본") == snapshot.find_snapshot(st, str(s1["id"])) == s1["id"]
    h = history(st, CFG)
    assert [x["kind"] for x in h][:2] == ["edit", "snapshot"] and h[1]["kind_ko"] == "판 확정"


def test_compare(two):
    st, R, tmp = two
    s1 = snapshot.freeze(st, CFG, R["admin"], "1차본")
    mem = snapshot.open_snapshot(st, s1["id"])
    same = snapshot.compare(mem, mem, CFG)
    assert not same["items"] and not any(same["counts"].values()) and same["end_shift"] == 0
    # 기간 · 이름 · 선행 · 줄 추가 · 줄 삭제 (줄을 넣으면 번호가 밀리므로 uid 를 먼저 잡아 둔다)
    u = {no: U(st, no) for no in (6, 10, 12, 19, 24)}
    apply_edits(st, CFG, R["fan"], st.rev(), [{"uid": u[10], "duration": 8}, {"uid": u[6], "name": "A 바꿈"}])
    apply_edits(st, CFG, R["hpc"], st.rev(), [{"uid": u[19], "preds": ""}])
    apply_row_op(st, CFG, R["fan"], st.rev(), {"op": "add", "where": "after", "uid": u[12], "name": "새 업무", "duration": 3})
    apply_row_op(st, CFG, R["hpc"], st.rev(), {"op": "delete", "uid": u[24]})          # E — 후속이 없다
    res = snapshot.compare(mem, st, CFG)
    c = res["counts"]
    assert (c["added"], c["deleted"], c["duration"], c["name"], c["preds"]) == (1, 1, 1, 1, 1)
    assert c["changed"] == 3 and c["dates"] >= 2 and res["end_shift"] > 0
    kinds = {(i["kind"], i["name"]) for i in res["items"]}
    assert ("added", "새 업무") in kinds and ("deleted", "E") in kinds and ("changed", "A 바꿈") in kinds
    dur = next(i for i in res["items"] if any(ch["field"] == "duration" for ch in i["changes"]))
    assert dur["name"] == "C" and next(ch for ch in dur["changes"] if ch["field"] == "duration")["b"] == "8" and dur["finish_shift"] == 4
    fan = next(r for r in res["modules"] if r["module"] == "Fan")
    assert fan["added"] == 1 and fan["changed"] == 2 and fan["finish_max"] >= 4
    out = snapshot.compare_xlsx(res, tmp / "cmp.xlsx", "1차본", "지금")
    wb = load_workbook(out)
    assert wb.sheetnames == ["요약", "모듈별", "바뀐 업무"]
    rows = list(wb["바뀐 업무"].iter_rows(values_only=True))
    assert rows[0][0] == "구분" and sum(1 for r in rows[1:] if r[8] == "기간(평일)") == 1
    mem.close()


# --- 백업 · 복구 ------------------------------------------------------------------------------------
def test_backup_restore(two, tmp_path):
    st, R, tmp = two
    cfg = CFG | {"backup": {"keep": 2}}
    db = st.path
    p, chk = backup.make_backup(db, cfg, "manual")
    assert chk["ok"] and chk["tasks"] == 24 and p.parent == db.parent / "backup" and backup.NAME_RE.match(p.name)
    for _ in range(3):
        backup.make_backup(db, cfg, "manual")
    backup.make_backup(db, cfg, "snapshot")
    files = backup.list_backups(db.parent / "backup")
    assert sum(b.kind == "manual" for b in files) == 2 and sum(b.kind == "snapshot" for b in files) == 1   # 수동은 최근 2개, 판 직전은 남김
    keep = files[-2].path                                        # 고치기 전 사본
    rev0 = st.rev()
    apply_edits(st, CFG, R["admin"], rev0, [{"uid": U(st, 7), "duration": 9}])
    st.close()
    with backup.ServingMarker(db):                               # 서버가 켜져 있으면 거절
        with pytest.raises(ValueError, match="서버가 켜져"):
            backup.restore(db, keep)
    bak, chk = backup.restore(db, keep)
    assert bak.exists() and chk["ok"]
    with Store(db) as again:
        assert again.rev() == rev0 and again.model(CFG).rows[7]["work_days"] != 9
    bad = tmp_path / "bad.db"
    bad.write_bytes(b"not a db")
    with pytest.raises(ValueError, match="검사 실패"):
        backup.restore(db, bad)


def test_daily_due(two):
    st, R, tmp = two
    cfg = CFG | {"backup": {"daily_at": "18:00"}}
    db = st.path
    day = datetime(2026, 11, 13, 17, 0)
    assert not backup.daily_due(db, cfg, day)                    # 아직 18시 전
    assert backup.daily_due(db, cfg, day.replace(hour=18, minute=5))   # 사본이 없음
    backup.make_backup(db, cfg, "start")
    assert not backup.daily_due(db, cfg, datetime.now().replace(hour=23, minute=0))   # 마지막 사본 뒤 저장 없음
    apply_edits(st, CFG, R["admin"], st.rev(), [{"uid": U(st, 7), "duration": 9}])
    assert backup.daily_due(db, cfg, datetime.now().replace(hour=23, minute=0))
    backup.make_backup(db, cfg, "daily")
    assert not backup.daily_due(db, cfg, datetime.now().replace(hour=23, minute=30))  # 오늘 것이 있음


# --- 수정 요청 ------------------------------------------------------------------------------------
def test_requests(two):
    st, R, tmp = two
    hpc_mod, fan_mod = U(st, 13), U(st, 1)
    with pytest.raises(EditError):
        collab.create_request(st, CFG, R["view"], {"kind": "pred", "text": "x", "to_module": fan_mod})
    with pytest.raises(EditError, match="내 모듈"):
        collab.create_request(st, CFG, R["hpc"], {"kind": "pred", "text": "x", "to_module": hpc_mod})
    with pytest.raises(EditError, match="내 모듈 업무"):
        collab.create_request(st, CFG, R["hpc"], {"kind": "pred", "text": "x", "to_module": fan_mod, "ref_uid": U(st, 7)})
    # HPC 리더: 'Fan 에 내 업무 19 의 선행이 될 업무가 없음'
    r = collab.create_request(st, CFG, R["hpc"], {"kind": "pred", "text": "시험 리그 준비 업무가 필요", "to_module": fan_mod, "ref_uid": U(st, 19)})
    assert r["status"] == "open" and r["to_module_name"] == "Fan" and r["ref"]["no"] == 19 and r["outbox"] and not r["inbox"]
    # 다른 모듈 업무를 대상으로 — 받는 모듈은 그 업무의 모듈
    r2 = collab.create_request(st, CFG, R["hpc"], {"kind": "fix", "text": "기간 확인", "task_uid": U(st, 12)})
    assert r2["to_module"] == fan_mod and r2["task"]["name"] == "E"
    assert collab.inbox_count(st, R["fan"]) == 2 and collab.inbox_count(st, R["hpc"]) == 0 and collab.inbox_count(st, R["admin"]) == 2
    assert collab.list_requests(st, CFG, R["view"]) == []
    with pytest.raises(EditError):
        collab.act_request(st, CFG, R["hpc"], r["id"], "done")            # 받은 쪽이 아님
    with pytest.raises(EditError, match="반려 이유"):
        collab.act_request(st, CFG, R["fan"], r["id"], "rejected")
    # Fan 리더가 업무를 만들고 처리 → 결과 업무
    res = apply_row_op(st, CFG, R["fan"], st.rev(), {"op": "add", "where": "after", "uid": U(st, 12), "name": "시험 리그 준비", "duration": 4})
    done = collab.act_request(st, CFG, R["fan"], r["id"], "done", "만들었음", res["uid"])
    assert done["status"] == "done" and done["closed_role"] == "FAN 리더" and done["result"]["name"] == "시험 리그 준비"
    with pytest.raises(EditError, match="닫힌"):
        collab.act_request(st, CFG, R["hpc"], r["id"], "withdrawn")
    w = collab.act_request(st, CFG, R["hpc"], r2["id"], "withdrawn")
    assert w["status"] == "withdrawn" and collab.inbox_count(st, R["fan"]) == 0
    mine = collab.list_requests(st, CFG, R["hpc"])
    assert [x["id"] for x in mine] == [r2["id"], r["id"]] and all(x["outbox"] for x in mine)
    assert st.rev() == res["rev"]                                # 요청은 판 번호를 올리지 않는다


# --- 모듈별 내보내기 --------------------------------------------------------------------------------
def test_module_export(two):
    st, R, tmp = two
    full = cells(st.export_xlsx(tmp / "full.xlsx", CFG))
    m = st.model(CFG)
    lv1 = module_of(m)
    mods = [m.wbs.uids[no] for no, t in m.wbs.tasks.items() if t.level == 1]
    parts, outside = [], 0
    for u in mods:
        p = st.export_xlsx(tmp / f"m{u}.xlsx", CFG, u)
        wb = load_workbook(p)
        rows = [tuple(c.value for c in r) for r in wb.worksheets[0].iter_rows()]
        assert rows[0] == full[0]                                # 같은 머리글
        parts += rows[1:]
        assert wb.sheetnames[1] == "밖 연결"
        outside += wb["밖 연결"].max_row - 1
    assert sorted(parts, key=lambda r: r[0]) == full[1:]         # 모듈 파일 줄을 합치면 전체 (번호·선행 그대로)
    cross = sum(1 for e in m.edges if lv1[e.pred] != lv1[e.succ])
    assert cross >= 1 and outside == 2 * cross                   # 한 연결이 양쪽 파일에 한 번씩
    hpc = load_workbook(tmp / f"m{mods[1]}.xlsx")["밖 연결"]
    first = [c.value for c in next(hpc.iter_rows(min_row=2))]
    assert first[0].startswith("들어옴") and first[5] == 12 and first[8] == "Fan"
    with pytest.raises(ValueError):
        st.export_xlsx(tmp / "x.xlsx", CFG, U(st, 7))            # Lv1 이 아님
    # 명령
    assert main(["export", "--db", str(st.path), "--module", "HPC", "--out", str(tmp / "hpc.xlsx")]) == 0
    assert cells(tmp / "hpc.xlsx") == cells(tmp / f"m{mods[1]}.xlsx")
    assert main(["export", "--db", str(st.path), "--by-module", "--out", str(tmp / "mods")]) == 0
    assert len(list((tmp / "mods").glob("*.xlsx"))) == 2


# --- 편집 열림/닫힘 --------------------------------------------------------------------------------
def test_edit_lock(two):
    st, R, tmp = two
    with pytest.raises(EditError):
        collab.set_edit_open(st, R["fan"], False)
    rev0 = st.rev()
    assert collab.set_edit_open(st, R["admin"], False, "11/20 까지 검토") == {"open": False, "note": "11/20 까지 검토"}
    assert st.rev() == rev0
    with pytest.raises(EditError) as e:
        apply_edits(st, CFG, R["fan"], st.rev(), [{"uid": U(st, 7), "duration": 9}])
    assert e.value.kind == "locked" and "11/20" in str(e.value)
    with pytest.raises(EditError) as e:
        apply_row_op(st, CFG, R["fan"], st.rev(), {"op": "add", "uid": U(st, 7), "name": "x"}, dry_run=True)
    assert e.value.kind == "locked"
    apply_edits(st, CFG, R["admin"], st.rev(), [{"uid": U(st, 7), "duration": 9}])          # 관리자는 됨
    collab.create_request(st, CFG, R["hpc"], {"kind": "etc", "text": "닫힌 동안 요청", "to_module": U(st, 1)})   # 요청은 됨
    assert collab.set_edit_open(st, R["admin"], True)["open"]
    apply_edits(st, CFG, R["fan"], st.rev(), [{"uid": U(st, 7), "duration": 8}])
    assert [x["kind"] for x in history(st, CFG)][:4] == ["edit", "lock", "edit", "lock"]


def test_cli_backup_snapshot(two, capsys):
    st, R, tmp = two
    snapshot.freeze(st, CFG, R["admin"], "1차본")
    db = str(st.path)
    assert main(["snapshot", "list", "--db", db]) == 0 and "1차본" in capsys.readouterr().out
    assert main(["backup", "--db", db]) == 0
    assert main(["backup", "list", "--db", db]) == 0 and "_manual.db" in capsys.readouterr().out
    assert main(["export", "--db", db, "--snapshot", "1차본", "--out", str(tmp / "s.xlsx")]) == 0
    assert main(["export", "--db", db, "--snapshot", "없는 판", "--out", str(tmp / "t.xlsx")]) == 2


# --- 서버 ---------------------------------------------------------------------------------------------
@pytest.fixture()
def client_two(two):
    pytest.importorskip("starlette")
    pytest.importorskip("httpx")
    from starlette.testclient import TestClient

    from engenius_wbs.server import create_app
    st, R, tmp = two
    app = create_app(st.path)

    def as_(who):
        c = TestClient(app)
        r = c.post("/login", data={"login": who, "password": PW, "mode": "role"}, follow_redirects=False)
        assert r.status_code == 303
        return c
    return app, as_, st, R


def test_server_snapshot_compare_export(client_two):
    app, as_, st, R = client_two
    admin, fan, view = as_("admin1"), as_("fan"), as_("view")
    assert fan.post("/api/snapshots", json={"label": "1차본"}).status_code == 403
    r = admin.post("/api/snapshots", json={"label": "1차본", "note": "취합"})
    assert r.status_code == 200 and r.json()["backup"].endswith("_snapshot.db")
    assert admin.post("/api/snapshots", json={"label": "1차본"}).status_code == 400
    sid = r.json()["id"]
    assert [x["label"] for x in view.get("/api/snapshots").json()] == ["1차본"]
    rev = fan.get("/api/status").json()["rev"]
    assert fan.post("/api/save", json={"base_rev": rev, "edits": [{"uid": U(st, 10), "duration": 8}]}).status_code == 200
    c = view.get(f"/api/compare?a={sid}&b=now").json()
    assert c["name_a"] == "1차본" and c["name_b"] == "지금" and c["counts"]["duration"] == 1 and c["items_total"] >= 1
    x = view.get(f"/compare.xlsx?a={sid}")
    assert x.status_code == 200 and load_workbook(BytesIO(x.content)).sheetnames == ["요약", "모듈별", "바뀐 업무"]
    s = view.get(f"/snapshot/{sid}.xlsx")
    assert s.status_code == 200 and "filename*=UTF-8''" in s.headers["content-disposition"]
    m = view.get(f"/snapshot/{sid}.xlsx?module={U(st, 13)}")
    assert load_workbook(BytesIO(m.content)).sheetnames[1] == "밖 연결"
    assert view.get(f"/export.xlsx?module={U(st, 1)}").status_code == 200
    assert view.get("/export.xlsx?module=x").status_code == 400
    assert view.get("/snapshot/99.xlsx").status_code == 400
    rp = view.get(f"/report.xlsx?snapshot={sid}")
    assert rp.status_code == 200 and "요약" in load_workbook(BytesIO(rp.content)).sheetnames
    assert view.get("/report.xlsx").status_code == 200


def test_server_requests_lock_backup(client_two):
    app, as_, st, R = client_two
    admin, fan, hpc, view = as_("admin1"), as_("fan"), as_("hpc"), as_("view")
    assert view.post("/api/requests", json={"kind": "pred", "text": "x", "to_module": U(st, 1)}).status_code == 403
    r = hpc.post("/api/requests", json={"kind": "pred", "text": "선행 필요", "to_module": U(st, 1), "ref_uid": U(st, 19)})
    assert r.status_code == 200
    rid = r.json()["id"]
    st_ = fan.get("/api/status").json()
    assert st_["inbox"] == 1 and st_["edit_open"] is True
    page = fan.get("/").text
    assert '"inbox": 1' in page and 'id="srvInbox"' in page and 'id="coTab"' in page and "const COLLAB=window.COLLAB" in page
    assert fan.get("/api/requests").json()[0]["can_act"] is True
    assert hpc.post(f"/api/requests/{rid}", json={"action": "done"}).status_code == 403
    d = fan.post(f"/api/requests/{rid}", json={"action": "done", "reply": "만듦", "result_uid": U(st, 12)})
    assert d.status_code == 200 and d.json()["result"]["no"] == 12
    assert fan.post(f"/api/requests/{rid}", json={"action": "rejected", "reply": "x"}).status_code == 409
    # 편집 닫힘
    assert fan.post("/api/lock", json={"open": False}).status_code == 403
    assert admin.post("/api/lock", json={"open": False, "note": "검토 기간"}).json() == {"open": False, "note": "검토 기간"}
    rev = fan.get("/api/status").json()["rev"]
    sv = fan.post("/api/save", json={"base_rev": rev, "edits": [{"uid": U(st, 7), "duration": 9}]})
    assert sv.status_code == 423 and sv.json()["kind"] == "locked"
    assert fan.post("/api/rows", json={"base_rev": rev, "dry_run": True, "op": {"op": "up", "uid": U(st, 7)}}).status_code == 423
    assert fan.get("/api/status").json()["edit_note"] == "검토 기간"
    admin.post("/api/lock", json={"open": True})
    assert fan.post("/api/save", json={"base_rev": rev, "edits": [{"uid": U(st, 7), "duration": 9}]}).status_code == 200
    # 백업
    assert fan.get("/api/backups").status_code == 403
    b = admin.post("/api/backups")
    assert b.status_code == 200 and b.json()["name"].endswith("_manual.db")
    assert [f["kind"] for f in admin.get("/api/backups").json()["files"]] == ["manual"]
    srv = app.state.wbs
    assert srv.backup_tick(datetime.now().replace(hour=23)) is None      # 마지막 사본 뒤 저장 없음
