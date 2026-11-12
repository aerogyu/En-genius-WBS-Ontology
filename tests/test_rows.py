"""줄 편집 (v0.5 S4) — 추가·삭제·이동·풀기 · 권한 · 미리 계산 = 저장 · 겹침 · 순환 · 내보낸 엑셀.

작은 WBS 두 모듈 (test_edit.make_two): Fan 1~12 · HPC 13~24.
  1 Fan > 2 설계통합 > 3 기본설계 > { 4 성능 > 5 성능 해석 > {6 A, 7 B} , 8 이차유로 > 9 유로 해석 > {10 C, 11 D, 12 E} }
  HPC 의 A(18) 는 Fan 의 E(12) 뒤.
"""
from __future__ import annotations

from io import BytesIO

import pytest
from openpyxl import load_workbook

from engenius_wbs.edit import EditError, apply_edits, history
from engenius_wbs.rows import apply_row_op
from test_edit import PW, two  # noqa: F401  (fixture)
from test_wbs import CFG


def M(st):
    return st.model(CFG)


def U(st, no):
    return M(st).wbs.uids[no]


def rev(st):
    return st.db.execute("SELECT COALESCE(MAX(id), 0) FROM txn").fetchone()[0]


def snap(st):
    """uid → (부모, 순서, 이름, 시작, 완료, 지움) · 선행 — 미리 계산이 아무것도 안 남기는지 볼 때."""
    t = st.db.execute("SELECT uid, parent_uid, pos, name, start, finish, deleted, start_floor FROM task ORDER BY uid").fetchall()
    d = st.db.execute("SELECT succ_uid, pred_uid, type, lag_days FROM dep ORDER BY id").fetchall()
    return [tuple(r) for r in t], [tuple(r) for r in d], rev(st)


def dates(st):
    return {r["uid"]: (r["start"], r["finish"]) for r in st.db.execute("SELECT uid, start, finish FROM task WHERE deleted = 0")}


def test_add_dry_run_equals_save(two):
    st, R, _ = two
    base, before = rev(st), snap(st)
    op = {"op": "add", "uid": U(st, 7), "where": "after", "name": "B2", "duration": 3, "preds": "7"}
    dry = apply_row_op(st, CFG, R["fan"], base, op, dry_run=True)
    assert snap(st) == before                                              # 미리 계산은 아무것도 남기지 않는다
    res = apply_row_op(st, CFG, R["fan"], base, op)
    for k in ("no", "wbs", "renumbered", "moved", "moved_other", "end_before", "end_after"):
        assert dry[k] == res[k], k                                          # 확인 창 숫자 = 저장 결과
    m = M(st)
    t = m.wbs.tasks[res["no"]]
    assert (res["no"], t.wbs, t.name, t.level) == (8, "1.1.1.1.1.3", "B2", 6)
    assert res["renumbered"] == 17                                         # 9~25 번 (옛 8~24) 이 한 칸씩
    b = m.wbs.tasks[7]
    assert t.deps[0].pred == 7 and t.start > b.finish                       # 선행 7 (B) 뒤에 시작
    hpc_a = m.wbs.tasks[19]
    assert hpc_a.name == "A" and [d.pred for d in hpc_a.deps] == [13]       # 번호가 바뀌어도 HPC A → Fan E 연결은 그대로 (uid)
    h = history(st, CFG, limit=1)[0]
    assert h["kind"] == "rows" and h["role"] == "FAN 리더" and {c["field"] for c in h["changes"]} >= {"add", "preds", "children"}


def test_add_rules(two):
    st, R, _ = two
    base = rev(st)
    with pytest.raises(EditError) as e:
        apply_row_op(st, CFG, R["hpc"], base, {"op": "add", "uid": U(st, 7), "name": "x"})
    assert e.value.kind == "denied"
    with pytest.raises(EditError) as e:
        apply_row_op(st, CFG, R["view"], base, {"op": "add", "uid": U(st, 7), "name": "x"})
    assert e.value.kind == "denied"
    with pytest.raises(EditError) as e:                                     # Lv1 옆에 = 새 모듈 → 관리자만
        apply_row_op(st, CFG, R["fan"], base, {"op": "add", "uid": U(st, 1), "name": "x"})
    assert e.value.kind == "denied"
    with pytest.raises(EditError, match="깊게"):                             # Lv6 아래 = Lv7 (max_level 6)
        apply_row_op(st, CFG, R["fan"], base, {"op": "add", "uid": U(st, 6), "where": "child", "name": "x"})
    with pytest.raises(EditError, match="이름"):
        apply_row_op(st, CFG, R["fan"], base, {"op": "add", "uid": U(st, 6), "name": "  "})
    with pytest.raises(EditError, match="없는 번호"):
        apply_row_op(st, CFG, R["fan"], base, {"op": "add", "uid": U(st, 6), "name": "x", "preds": "999"})
    res = apply_row_op(st, CFG, R["fan"], base, {"op": "add", "uid": U(st, 9), "where": "child", "name": "F", "duration": 0})
    t = M(st).wbs.tasks[res["no"]]
    assert t.wbs == "1.1.1.2.1.4" and t.is_milestone                       # 하위 맨 뒤, 기간 0 = 마일스톤
    res = apply_row_op(st, CFG, R["admin"], rev(st), {"op": "add", "uid": U(st, 1), "name": "LPT"})
    assert M(st).wbs.tasks[res["no"]].wbs == "2" and M(st).wbs.tasks[res["no"]].level == 1


def test_delete_other_module_successor(two):
    st, R, _ = two
    e_uid, hpc_a = U(st, 12), U(st, 18)
    with pytest.raises(EditError) as e:                                     # HPC A 가 Fan E 를 선행으로 씀 → 리더는 못 지움
        apply_row_op(st, CFG, R["fan"], rev(st), {"op": "delete", "uid": e_uid})
    assert e.value.kind == "denied" and e.value.items[0]["uid"] == hpc_a
    with pytest.raises(EditError) as e:                                     # 요약(유로 해석)을 지워도 아래의 E 때문에 같음
        apply_row_op(st, CFG, R["fan"], rev(st), {"op": "delete", "uid": U(st, 9)})
    assert e.value.kind == "denied"
    d0 = dates(st)
    res = apply_row_op(st, CFG, R["admin"], rev(st), {"op": "delete", "uid": e_uid})
    assert res["deleted"] == 1 and res["moved_other"] == 0 and any("다른 모듈" in w for w in res["warnings"])
    m = M(st)
    a = m.wbs.tasks[{u: n for n, u in m.wbs.uids.items()}[hpc_a]]
    assert a.deps == [] and dates(st)[hpc_a] == d0[hpc_a]                   # 연결은 빠지고, 날짜는 당겨지지 않음 (시작 하한)
    assert st.db.execute("SELECT start_floor FROM task WHERE uid = ?", (hpc_a,)).fetchone()[0] == d0[hpc_a][0]
    assert st.db.execute("SELECT deleted FROM task WHERE uid = ?", (e_uid,)).fetchone()[0] == 1   # 지우지 않고 표시만
    assert st.count() == 23


def test_delete_subtree_same_module(two):
    st, R, _ = two
    d0 = dates(st)
    c, d = U(st, 10), U(st, 11)
    res = apply_row_op(st, CFG, R["fan"], rev(st), {"op": "delete", "uid": U(st, 4)})   # 성능 > 성능 해석 > A, B
    assert res["deleted"] == 4 and st.count() == 20 and res["moved"] == 0
    assert any("아래 업무 3개" in w for w in res["warnings"])
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    assert m.wbs.tasks[no_of[c]].deps == [] and m.wbs.tasks[no_of[d]].deps == []        # A·B 에 걸린 연결이 빠짐
    assert dates(st)[c] == d0[c] and dates(st)[d] == d0[d]
    assert m.wbs.tasks[no_of[c]].wbs == "1.1.1.1.1.1"                                   # 이차유로가 1.1.1.1 이 됨


def test_move_up_down_and_conflict(two):
    st, R, _ = two
    base = rev(st)
    perf, sec = U(st, 4), U(st, 8)
    with pytest.raises(EditError, match="맨 위"):
        apply_row_op(st, CFG, R["fan"], base, {"op": "up", "uid": perf})
    d0 = dates(st)
    res = apply_row_op(st, CFG, R["fan"], base, {"op": "down", "uid": perf})
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    assert m.wbs.tasks[no_of[perf]].wbs == "1.1.1.2" and m.wbs.tasks[no_of[sec]].wbs == "1.1.1.1"
    assert res["moved"] == 0 and dates(st) == d0                           # 순서만 바뀜, 날짜 그대로
    h = history(st, CFG, uid=perf, limit=1)[0]
    assert any(c["field"] == "move" and c["old"] == "1.1.1.1" and c["new"] == "1.1.1.2" for c in h["changes"])
    with pytest.raises(EditError) as e:                                     # 같은 판을 보고 같은 묶음 순서를 또 바꿈 → 거절
        apply_row_op(st, CFG, R["fan"], base, {"op": "up", "uid": sec})
    assert e.value.kind == "conflict" and e.value.items[0]["role"] == "FAN 리더"
    apply_row_op(st, CFG, R["hpc"], base, {"op": "down", "uid": U(st, 18)})  # 다른 묶음은 그 위에 반영
    with pytest.raises(EditError) as e:                                     # 옮긴 업무를 옛 판에서 칸 고치기 → 거절
        apply_edits(st, CFG, R["fan"], base, [{"uid": perf, "name": "성능2"}])
    assert e.value.kind == "conflict"


def test_indent_outdent(two):
    st, R, _ = two
    cfg7 = dict(CFG, max_level=7)
    a, b = U(st, 6), U(st, 7)
    with pytest.raises(EditError, match="깊어"):
        apply_row_op(st, CFG, R["fan"], rev(st), {"op": "indent", "uid": b})
    with pytest.raises(EditError, match="위 형제"):
        apply_row_op(st, cfg7, R["fan"], rev(st), {"op": "indent", "uid": a})
    res = apply_row_op(st, cfg7, R["fan"], rev(st), {"op": "indent", "uid": b})
    assert any("요약이 된다" in w for w in res["warnings"])
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    assert m.wbs.tasks[no_of[b]].wbs == "1.1.1.1.1.1.1" and m.wbs.tasks[no_of[a]].children == [no_of[b]]
    res = apply_row_op(st, cfg7, R["fan"], rev(st), {"op": "outdent", "uid": b})       # 되돌림 → A 는 다시 하위 업무
    assert any("하위 업무가 된다" in w for w in res["warnings"])
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    assert m.wbs.tasks[no_of[b]].wbs == "1.1.1.1.1.2" and not m.wbs.tasks[no_of[a]].children
    with pytest.raises(EditError) as e:                                     # Lv2 내어쓰기 = 새 모듈 → 관리자만
        apply_row_op(st, CFG, R["fan"], rev(st), {"op": "outdent", "uid": U(st, 2)})
    assert e.value.kind == "denied"


def test_outdent_leaves_parent_as_leaf(two):
    st, R, _ = two
    d0 = dates(st)
    perf = U(st, 4)
    apply_row_op(st, CFG, R["fan"], rev(st), {"op": "outdent", "uid": U(st, 5)})       # 성능 해석 → 성능 옆
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    assert not m.wbs.tasks[no_of[perf]].children and dates(st)[perf] == d0[perf]       # 성능 = 하위 업무, 날짜 그대로
    assert dates(st) == d0


def test_ungroup_levels_up_dates_same(two):
    """Lv 줄이기: Lv3 '기본설계' 를 풀면 아래가 한 레벨씩 올라가고 날짜는 그대로."""
    st, R, _ = two
    d0 = dates(st)
    res = apply_row_op(st, CFG, R["fan"], rev(st), {"op": "ungroup", "uid": U(st, 3)})
    assert res["deleted"] == 1 and res["moved"] == 0 and st.count() == 23
    m = M(st)
    assert max(t.level for t in m.wbs.tasks.values() if t.wbs.startswith("1.")) == 5
    assert [m.wbs.tasks[n].name for n in m.wbs.tasks[2].children] == ["성능", "이차유로"]
    assert {u: v for u, v in dates(st).items()} == {u: v for u, v in d0.items() if u in dates(st)}


def test_ungroup_moves_summary_links(two):
    st, R, _ = two
    apply_edits(st, CFG, R["fan"], rev(st), [{"uid": U(st, 9), "preds": "5"}])         # 요약 '유로 해석' ← '성능 해석'
    d0 = dates(st)
    c, d, e = U(st, 10), U(st, 11), U(st, 12)
    res = apply_row_op(st, CFG, R["fan"], rev(st), {"op": "ungroup", "uid": U(st, 9)})
    assert res["moved"] == 0 and dates(st) == {u: v for u, v in d0.items() if u in dates(st)}   # FS 는 같은 뜻 → 날짜 그대로
    m = M(st)
    no_of = {u: n for n, u in m.wbs.uids.items()}
    perf_an = no_of[U(st, 5)]
    for u in (c, d, e):
        assert perf_an in [x.pred for x in m.wbs.tasks[no_of[u]].deps]


def test_cycle_rejected(two):
    st, R, _ = two
    with pytest.raises(EditError, match="고리"):                               # B(7) 는 A(6) 뒤 → A 를 B 뒤로 = 고리
        apply_edits(st, CFG, R["fan"], rev(st), [{"uid": U(st, 6), "preds": "7"}])
    assert rev(st) == 1


def test_export_after_rows(two):
    st, R, tmp = two
    d_uid = U(st, 11)                                                       # 넣기 전에 D 의 uid (넣으면 번호가 밀림)
    apply_row_op(st, CFG, R["fan"], rev(st), {"op": "add", "uid": U(st, 7), "name": "B2", "duration": 3, "preds": "7"})
    apply_row_op(st, CFG, R["fan"], rev(st), {"op": "delete", "uid": d_uid})
    out = st.export_xlsx(tmp / "out.xlsx", CFG)
    ws = load_workbook(out).active
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    names = [str(r[2]).strip() for r in rows if r[0]]
    assert "B2" in names and names.count("D") == 1 and len(names) == 24      # Fan 의 D 만 빠짐 (HPC 의 D 는 남음)
    nos = [r[0] for r in rows if r[0]]
    assert nos == list(range(1, 25))                                        # 번호는 1부터 빈틈 없이 다시
    b2 = next(r for r in rows if r[2] and str(r[2]).strip() == "B2")
    assert (b2[0], b2[1], str(b2[6])) == (8, "1.1.1.1.1.3", "7")


def test_rows_api(two):
    pytest.importorskip("httpx")
    from starlette.testclient import TestClient

    from engenius_wbs.server import create_app
    st, R, _ = two
    app = create_app(st.path)

    def client(who):
        c = TestClient(app)
        assert c.post("/login", data={"login": who, "password": PW, "mode": "role"}, follow_redirects=False).status_code == 303
        return c
    fan, view = client("fan"), client("view")
    body = {"base_rev": 1, "dry_run": True, "op": {"op": "add", "uid": U(st, 7), "name": "B2"}}
    assert view.post("/api/rows", json=body).status_code == 403
    assert fan.post("/api/rows", data="x", headers={"content-type": "text/plain"}).status_code == 415
    r = fan.post("/api/rows", json=body)
    assert r.status_code == 200 and r.json()["dry_run"] and r.json()["renumbered"] == 17 and r.json()["txn"] is None
    assert fan.get("/api/status").json()["rev"] == 1                       # 미리 계산은 판을 올리지 않음
    r = fan.post("/api/rows", json=body | {"dry_run": False})
    assert r.status_code == 200 and r.json()["rev"] == 2
    assert fan.post("/api/rows", json={"base_rev": 1, "op": {"op": "up", "uid": U(st, 7)}}).status_code == 409
    assert fan.post("/api/rows", json={"base_rev": 2, "op": {"op": "nope", "uid": 1}}).status_code == 400


def test_fix_relax(two):
    """들여올 때 어겨 있던 만큼(느슨함)을 바로잡으면 날짜가 선행을 따라 밀린다 — 칸 고치기(fix_relax)로 저장."""
    st, R, _ = two
    a, b = U(st, 6), U(st, 7)
    st.db.execute("UPDATE dep SET relax_days = 2 WHERE succ_uid = ? AND pred_uid = ?", (b, a))
    st.db.execute("UPDATE task SET start_floor = NULL WHERE uid = ?", (b,))
    st.recompute(CFG)
    st.db.commit()
    s0 = dates(st)[b][0]
    with pytest.raises(EditError):
        apply_edits(st, CFG, R["fan"], rev(st), [{"uid": b, "fix_relax": 1}])
    with pytest.raises(EditError) as e:
        apply_edits(st, CFG, R["hpc"], rev(st), [{"uid": b, "fix_relax": True}])
    assert e.value.kind == "denied"
    res = apply_edits(st, CFG, R["fan"], rev(st), [{"uid": b, "fix_relax": True}])
    assert res["moved"] >= 1 and dates(st)[b][0] > s0
    assert st.db.execute("SELECT SUM(relax_days) FROM dep WHERE succ_uid = ?", (b,)).fetchone()[0] == 0
    c = history(st, CFG, uid=b, limit=1)[0]["changes"][0]
    assert (c["label"], c["old"], c["new"]) == ("느슨함(평일)", "2", "0")
