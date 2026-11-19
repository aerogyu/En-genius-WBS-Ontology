"""일정 옮기기 (v0.6) — 전체·묶음을 n 평일 통째로: 정확히 n · 되돌리면 원래대로 · 권한 · 겹침 · 범위 밖 선행에 묶임 · 파일 모드(apply)."""
from __future__ import annotations

import json
from datetime import date

import pytest
from openpyxl import Workbook

from engenius_wbs.apply import apply_changes, sha256_of
from engenius_wbs.bulk import shift
from engenius_wbs.collab import set_edit_open
from engenius_wbs.edit import EditError, apply_edits, history
from engenius_wbs.parse import read_wbs
from engenius_wbs.schedule import widx
from engenius_wbs.store import Store
from test_collab import client_two  # noqa: F401  (fixture)
from test_edit import dates, make_two, two, uid  # noqa: F401  (fixture)
from test_wbs import CFG, make_small

LEAVES = (6, 7, 10, 11, 12)                 # Fan 의 하위 업무 (HPC = +12)


def wd(a: date, b: date) -> int:
    return widx(b) - widx(a)


def floors(st) -> dict:
    return {r["uid"]: r["start_floor"] for r in st.db.execute("SELECT uid, start_floor FROM task WHERE deleted = 0")}


def test_shift_all_exact_and_back(two):
    st, R, _ = two
    d0, f0 = dates(st), floors(st)
    dry = shift(st, CFG, R["admin"], 1, None, 7, dry_run=True)
    assert dry["dry_run"] and dry["txn"] is None and dates(st) == d0 and st.rev() == 1      # 미리 계산은 아무것도 남기지 않는다
    res = shift(st, CFG, R["admin"], 1, None, 7)
    assert res["rev"] == 2 and res["leaves"] == 10 and res["held"] == 0 and res["scope"] == "전체"
    assert {k: v for k, v in res.items() if k not in ("txn", "rev")} == {k: v for k, v in dry.items() if k not in ("txn", "rev", "dry_run")}
    d1 = dates(st)
    for no, (s, f) in d0.items():
        if s:                                                                               # 요약도 하위 범위를 따라 같이
            assert wd(s, d1[no][0]) == 7 and wd(f, d1[no][1]) == 7, no
    assert wd(date.fromisoformat(res["end_before"]), date.fromisoformat(res["end_after"])) == 7
    h = history(st, CFG)[0]
    assert h["role"] == "관리자 1" and h["note"]["shift"] == {"uid": None, "scope": "전체", "days": 7, "leaves": 10, "held": 0}
    assert h["n_changes"] == res["floors"] and all(c["field"] == "start" for c in h["changes"])
    back = shift(st, CFG, R["admin"], 2, None, -7)
    assert back["held"] == 0 and dates(st) == d0 and floors(st) == f0                       # 되돌리면 날짜도 하한도 원래대로


def test_shift_module_later_pushes_followers_only(two):
    st, R, _ = two
    d0, f0 = dates(st), floors(st)
    res = shift(st, CFG, R["hpc"], 1, uid(st, 13), 5)                                       # 리더가 자기 모듈(Lv1)을 통째로
    d1 = dates(st)
    assert res["leaves"] == 5 and res["held"] == 0 and res["scope"].endswith("HPC")
    assert all(d1[no] == d0[no] for no in range(1, 13))                                     # 앞 모듈(Fan)은 그대로
    assert all(wd(d0[no + 12][0], d1[no + 12][0]) == 5 and wd(d0[no + 12][1], d1[no + 12][1]) == 5 for no in LEAVES)
    f1 = floors(st)
    assert f1[uid(st, 18)] == d1[18][0].isoformat()                                         # Fan 의 E 에 붙어 있던 A' 는 새 시작이 하한
    assert f1[uid(st, 19)] is None and f0[uid(st, 19)] is None                              # 범위 안 선행을 따르던 B' 는 계속 따른다 (하한 없음)
    apply_edits(st, CFG, R["hpc"], 2, [{"uid": uid(st, 18), "duration": 3}])                # A' 를 줄이면 B' 가 따라 당겨진다
    assert dates(st)[19][0] < d1[19][0]


def test_shift_earlier_is_held_by_outside_preds(two):
    st, R, _ = two
    d0 = dates(st)
    with pytest.raises(EditError) as e:                                                     # HPC 는 통째로 Fan 의 E 에 붙어 있다 — 하나도 못 당기면 저장하지 않는다
        shift(st, CFG, R["hpc"], 1, uid(st, 13), -4)
    assert e.value.kind == "invalid" and st.rev() == 1 and dates(st) == d0
    shift(st, CFG, R["hpc"], 1, uid(st, 13), 5)                                             # 5 늦춘 뒤
    res = shift(st, CFG, R["hpc"], 2, uid(st, 13), -8)                                      # 8 당기면 — E 가 허락하는 5 만 당겨진다
    assert res["held"] == 5 and res["leaves"] == 5 and dates(st) == d0


def test_shift_leaf_and_rights(two):
    st, R, _ = two
    d0 = dates(st)
    for who, target, kind in (("view", uid(st, 6), "denied"), ("fan", None, "denied"), ("fan", uid(st, 18), "denied"), ("hpc", uid(st, 1), "denied")):
        with pytest.raises(EditError) as e:
            shift(st, CFG, R[who], 1, target, 3)
        assert e.value.kind == kind, who
    for bad in (0, 1.5, "3", True, 99999):
        with pytest.raises(EditError) as e:
            shift(st, CFG, R["admin"], 1, None, bad)
        assert e.value.kind == "invalid"
    assert dates(st) == d0 and st.rev() == 1
    res = shift(st, CFG, R["fan"], 1, uid(st, 11), 8)                                       # 하위 업무 하나 (D) — 뒤의 E(11FF) 가 밀린다
    d1 = dates(st)
    assert res["leaves"] == 1 and wd(d0[11][0], d1[11][0]) == 8 and d1[12][1] > d0[12][1] and d1[6] == d0[6]
    set_edit_open(st, R["admin"], False, "취합 중")                                          # 편집이 닫히면 관리자만
    with pytest.raises(EditError) as e:
        shift(st, CFG, R["fan"], st.rev(), uid(st, 11), 1)
    assert e.value.kind == "locked"
    assert shift(st, CFG, R["admin"], st.rev(), uid(st, 11), 1)["leaves"] == 1


def test_shift_conflict_with_other_save(two):
    st, R, _ = two
    apply_edits(st, CFG, R["fan"], 1, [{"uid": uid(st, 7), "duration": 8}])                 # 내가 본 판(1) 뒤에 Fan 리더가 B 를 고침
    with pytest.raises(EditError) as e:
        shift(st, CFG, R["admin"], 1, uid(st, 1), 5)
    assert e.value.kind == "conflict" and e.value.items[0]["no"] == 7 and e.value.items[0]["role"] == "FAN 리더"
    assert shift(st, CFG, R["hpc"], 1, uid(st, 13), 5)["leaves"] == 5                       # 다른 모듈은 겹치지 않는다
    with pytest.raises(EditError) as e:                                                     # 옮긴 뒤의 옛 화면 저장은 거절
        apply_edits(st, CFG, R["hpc"], 1, [{"uid": uid(st, 18), "duration": 2}])
    assert e.value.kind == "conflict"


def test_shift_summary_without_dated_children(tmp_path):
    """하위 업무에 날짜가 없고 요약에만 날짜가 적힌 묶음 — 적힌 날짜를 그대로 옮긴다."""
    wb = Workbook()
    ws = wb.active
    ws.append(["번호", "WBS", "작업 이름", "기간", "시작", "완료", "선행 작업"])
    for row in ((1, "1", "Fan", None, None, None, None), (2, "1.1", "   묶음 가", "5 일", "27-01-04 (월)", "27-01-08 (금)", None),
                (3, "1.1.1", "      날짜 없는 업무", None, None, None, None), (4, "1.2", "   업무 나", "2 일", "27-01-11 (월)", "27-01-12 (화)", None)):
        ws.append(row)
    wb.save(tmp_path / "bare.xlsx")
    from engenius_wbs import auth
    with Store(tmp_path / "wbs.db") as st:
        st.import_xlsx(tmp_path / "bare.xlsx", CFG)
        admin = auth.add_role(st.db, "admin1", "관리자 1", "admin", "test-pass-1234")
        d0 = dates(st)
        res = shift(st, CFG, admin, 1, None, 5)
        d1 = dates(st)
        assert res["leaves"] == 1 and d1[2] == (date(2027, 1, 11), date(2027, 1, 15)) and d1[4] == (date(2027, 1, 18), date(2027, 1, 19))
        assert d1[3] == d0[3] == (None, None) and d1[1][0] == date(2027, 1, 11)
        assert [c["field"] for c in history(st, CFG)[0]["changes"]] == ["start", "shift"]


def test_shift_api(client_two):
    app, as_, st, R = client_two
    admin, fan, view = as_("admin1"), as_("fan"), as_("view")
    assert view.post("/api/shift", json={"base_rev": 1, "uid": None, "days": 5}).status_code == 403
    assert fan.post("/api/shift", json={"base_rev": 1, "uid": None, "days": 5}).status_code == 403
    assert admin.post("/api/shift", json={"base_rev": 1, "uid": None, "days": 0}).status_code == 400
    r = admin.post("/api/shift", json={"base_rev": 1, "uid": None, "days": 5, "dry_run": True})
    assert r.status_code == 200 and r.json()["dry_run"] and admin.get("/api/status").json()["rev"] == 1
    r = admin.post("/api/shift", json={"base_rev": 1, "uid": None, "days": 5})
    assert r.status_code == 200 and r.json()["rev"] == 2 and r.json()["leaves"] == 10
    assert admin.get("/api/status").json()["rev"] == 2 and '"rev":2' in admin.get("/data.js").text
    h = admin.get("/api/history?limit=3").json()[0]
    assert h["note"]["shift"]["days"] == 5 and h["n_changes"] >= len(h["changes"])
    assert fan.post("/api/shift", json={"base_rev": 1, "uid": uid(st, 1), "days": 2}).status_code == 409     # 옛 판에서
    assert fan.post("/api/shift", json={"base_rev": 2, "uid": uid(st, 1), "days": 2}).status_code == 200


def test_apply_shift_file_mode(tmp_path):
    """파일 모드 — 변경 목록의 shift (no 0 = 전체) 를 apply 가 '이 날 이후' 로 풀어 다시 짠다."""
    src = make_small(tmp_path / "small.xlsx")
    w0 = read_wbs(src, CFG)

    def run(changes, name):
        ch = {"format": "engenius-wbs-changes/1", "source": "small.xlsx", "sheet": w0.sheet, "sha256": sha256_of(src), "changes": changes}
        (tmp_path / name).write_text(json.dumps(ch, ensure_ascii=False), encoding="utf-8")
        out = tmp_path / name.replace(".json", "")
        out.mkdir()
        return read_wbs(apply_changes(src, tmp_path / name, CFG, out).out, CFG)

    w1 = run([{"no": 0, "field": "shift", "to": 10}], "all.json")                           # 전체 +10 평일 = 달력 2주
    for no in LEAVES:
        assert wd(w0.tasks[no].start, w1.tasks[no].start) == 10 and wd(w0.tasks[no].finish, w1.tasks[no].finish) == 10, no
    w2 = run([{"no": 8, "field": "shift", "to": 3}], "sub.json")                            # 이차유로(8) 아래 C·D·E 만
    assert all(w2.tasks[no].start == w0.tasks[no].start for no in (6, 7))
    assert all(wd(w0.tasks[no].start, w2.tasks[no].start) == 3 for no in (10, 11, 12))
    w3 = run([{"no": 0, "field": "shift", "to": 10}, {"no": 6, "field": "start", "to": "2027-02-01"}], "mix.json")   # 손으로 고친 시작이 먼저
    assert w3.tasks[6].start == date(2027, 2, 1) and wd(w0.tasks[11].start, w3.tasks[11].start) >= 10
    with pytest.raises(ValueError):
        run([{"no": 999, "field": "shift", "to": 3}], "bad.json")
