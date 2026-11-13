"""칸 고치기 저장 (v0.5 S3) — 화면의 가정 편집 묶음을 도구 DB 에 한 번에 반영한다 (docs/05 5장).

  권한     관리자: 전부 · 모듈 리더: 자기 모듈(Lv1 아래) 업무만, Lv1 줄 자체는 관리자만 · 조회자: 없음 — 저장마다 서버가 검사
  겹침     내가 본 판(base_rev) 뒤에 다른 저장이 같은 업무를 고쳤으면 통째로 거절 (덮어쓰지 않는다). 다른 업무만 고쳤으면 그 위에 반영.
  반영     한 트랜잭션: txn(역할 이름) → 칸마다 change(old → new) → 날짜 다시 짜기(store.recompute = schedule.forward) → 커밋
  고칠 칸  하위 업무: 기간(일하는 평일 수) · 시작 하한('이 날 이후') · 선행 · 이름 · Man Hours / 요약 업무: 선행 · 이름

선행은 화면의 번호 표기(MS Project, '12,15SS+5일')로 받아 지금 번호 → uid 로 바꿔 둔다 (번호가 바뀌어도 연결은 그대로).
고치지 않은 선행의 '느슨함'(relax_days)은 그대로, 새 선행은 0. 선행을 지웠는데 시작 하한이 없으면 지금 시작을 하한으로 둔다
(지워도 저절로 당겨지지 않게 — v0.4 와 같은 동작, docs/05 2-2). 화면 미리 보기(JS)는 같은 규칙으로 계산한다.
"""
from __future__ import annotations

import json
import math
import re
from datetime import date

from .auth import Role
from .config import nfkc
from .parse import DEP_RE, parse_deps
from .rows import EditError, check_cycles, check_open, cycle_uids
from .store import Store, _units

FIELDS = ("duration", "start", "preds", "name", "man_hours", "fix_relax")
LEAF_ONLY = ("duration", "start", "man_hours")
LABEL = {"duration": "기간(평일)", "start": "시작 하한", "preds": "선행", "name": "이름", "man_hours": "Man Hours",
         "fix_relax": "느슨함(평일)", "relax": "느슨함(평일)", "add": "추가", "delete": "삭제", "move": "위치(WBS)", "children": "하위 줄", "ungroup": "풀기"}
KIND_KO = {"edit": "칸 고치기", "rows": "줄 편집", "snapshot": "판 확정", "lock": "편집 열고 닫기"}
MAX_EDITS = 2000


def _fmt(v) -> str | None:
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _num(v, what: str, lo: float = 0.0, hi: float = 1e7) -> float:
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or not lo <= v <= hi:
        raise EditError("invalid", f"{what} 는 {lo:g} 이상 숫자")
    return float(v)


def apply_edits(st: Store, cfg: dict, role: Role, base_rev: int, edits: list[dict]) -> dict:
    """→ {txn, rev, tasks, moved, end_before, end_after}. 안 되면 EditError (아무것도 쓰지 않음)."""
    if role.kind == "viewer":
        raise EditError("denied", "조회자는 저장할 수 없다")
    check_open(st, role)
    if not isinstance(edits, list) or not edits or len(edits) > MAX_EDITS:
        raise EditError("invalid", "고친 것이 없거나 너무 많다")
    m = st.model(cfg)
    tasks = m.wbs.tasks
    no_of = {uid: no for no, uid in m.wbs.uids.items()}
    units = _units(cfg)
    cyc0 = cycle_uids(m)

    def module_uid(no: int) -> int:
        while tasks[no].parent is not None:
            no = tasks[no].parent
        return m.wbs.uids[no]

    # --- 검사 (쓰기 전에 전부) ---
    plan: list[tuple[int, dict]] = []
    denied, bad = [], []
    for e in edits:
        if not isinstance(e, dict) or not isinstance(e.get("uid"), int) or e["uid"] not in no_of:
            raise EditError("invalid", f"없는 업무: {e.get('uid') if isinstance(e, dict) else e}")
        no = no_of[e["uid"]]
        t = tasks[no]
        if not role.is_admin and (t.level == 1 or module_uid(no) != role.module_uid):
            denied.append({"uid": e["uid"], "no": no, "wbs": t.wbs, "name": t.name})
            continue
        ch = {k: e[k] for k in FIELDS if k in e}
        if not ch:
            continue
        try:
            if t.children and any(k in ch for k in LEAF_ONLY):
                raise EditError("invalid", "요약 업무는 선행·이름만 고칠 수 있다")
            if "duration" in ch:
                ch["duration"] = _num(ch["duration"], "기간", 0, 1e5)
            if "start" in ch:
                ch["start"] = date.fromisoformat(str(ch["start"])).isoformat() if ch["start"] else None
            if "man_hours" in ch and ch["man_hours"] is not None:
                ch["man_hours"] = _num(ch["man_hours"], "Man Hours")
            if "fix_relax" in ch and ch["fix_relax"] is not True:
                raise EditError("invalid", "fix_relax 는 true 만")
            if "name" in ch:
                ch["name"] = nfkc(ch["name"]).strip()
                if not ch["name"] or len(ch["name"]) > 300:
                    raise EditError("invalid", "이름이 비었거나 너무 길다")
            if "preds" in ch:
                text = nfkc(ch["preds"] or "").strip()
                deps, badtok = parse_deps(text, units)
                if badtok:
                    raise EditError("invalid", "선행 표기를 읽지 못함: " + ", ".join(badtok))
                miss = [d.pred for d in deps if d.pred not in tasks or d.pred == no]
                if miss:
                    raise EditError("invalid", "없는 번호·자기 자신: " + ", ".join(map(str, miss)))
                toks = [tok.strip() for tok in re.split(r"[,;]", text) if tok.strip()]
                ch["preds"] = [(d, DEP_RE.match(tok)) for d, tok in zip(deps, toks)]
        except (EditError, ValueError) as x:
            bad.append({"uid": e["uid"], "no": no, "wbs": t.wbs, "name": t.name, "error": str(x)})
            continue
        plan.append((no, ch))
    if denied:
        raise EditError("denied", f"고칠 권한이 없는 업무 {len(denied)}개 (내 모듈이 아님)", denied)
    if bad:
        raise EditError("invalid", f"고칠 수 없는 값 {len(bad)}개", bad)
    if not plan:
        raise EditError("invalid", "바뀐 칸이 없다")
    uids = [m.wbs.uids[no] for no, _ in plan]
    q = ",".join("?" * len(uids))
    hit = st.db.execute(f"SELECT DISTINCT c.task_uid, x.id, x.at, x.role FROM change c JOIN txn x ON x.id = c.txn_id "
                        f"WHERE x.kind IN ('edit', 'rows') AND c.field != 'children' AND x.id > ? AND c.task_uid IN ({q}) ORDER BY x.id", (base_rev, *uids)).fetchall()
    if hit:
        items = [{"uid": r["task_uid"], "no": no_of[r["task_uid"]], "wbs": tasks[no_of[r["task_uid"]]].wbs,
                  "name": tasks[no_of[r["task_uid"]]].name, "txn": r["id"], "at": r["at"], "role": r["role"]} for r in hit]
        raise EditError("conflict", f"내가 본 판 뒤에 다른 저장이 같은 업무 {len({i['uid'] for i in items})}개를 고쳤다", items)

    # --- 쓰기 (한 트랜잭션) ---
    end_before = st.db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]
    try:
        txn = st._txn("edit", role.label)

        def log(uid, field, old, new):
            st.db.execute("INSERT INTO change(txn_id, task_uid, field, old, new) VALUES(?,?,?,?,?)", (txn, uid, field, _fmt(old), _fmt(new)))

        for no, ch in plan:
            r = m.rows[no]
            uid = r["uid"]
            if "duration" in ch and ch["duration"] != r["work_days"]:
                log(uid, "duration", r["work_days"], ch["duration"])
                st.db.execute("UPDATE task SET work_days = ?, duration_days = ? WHERE uid = ?", (ch["duration"], ch["duration"], uid))
            if "start" in ch and ch["start"] != r["start_floor"]:
                log(uid, "start", r["start_floor"], ch["start"])
                st.db.execute("UPDATE task SET start_floor = ? WHERE uid = ?", (ch["start"], uid))
            if "name" in ch and ch["name"] != nfkc(r["name"]).strip():
                log(uid, "name", nfkc(r["name"]).strip(), ch["name"])
                st.db.execute("UPDATE task SET name = ? WHERE uid = ?", (ch["name"], uid))
            if "man_hours" in ch and ch["man_hours"] != r["man_hours"]:
                log(uid, "man_hours", r["man_hours"], ch["man_hours"])
                st.db.execute("UPDATE task SET man_hours = ? WHERE uid = ?", (ch["man_hours"], uid))
            if "preds" in ch:
                old = m.deps.get(no, [])
                old_text = ",".join(f"{no_of[d['pred_uid']]}{d['suffix']}" for d in old)
                relax = {(d["pred_uid"], d["type"], d["lag_days"]): d["relax_days"] for d in old}
                new = [(m.wbs.uids[d.pred], d, mt) for d, mt in ch["preds"]]
                new_text = ",".join(f"{d.pred}{mt.group(0)[len(mt.group(1)):].strip()}" for _, d, mt in new)
                key = lambda xs: sorted((a, b or "FS", c or 0.0) for a, b, c in xs)
                if key([(d["pred_uid"], d["type"], d["lag_days"]) for d in old]) == key([(u, d.type, d.lag_days) for u, d, _ in new]):
                    continue
                log(uid, "preds", old_text, new_text)
                st.db.execute("DELETE FROM dep WHERE succ_uid = ?", (uid,))
                for k, (pu, d, mt) in enumerate(new):
                    st.db.execute("INSERT INTO dep(succ_uid, pred_uid, ord, type, lag_days, lag_text, suffix, relax_days) VALUES(?,?,?,?,?,?,?,?)",
                                  (uid, pu, k, d.type, d.lag_days, d.lag_text, mt.group(0)[len(mt.group(1)):].strip(),
                                   relax.get((pu, d.type, d.lag_days), 0.0)))
                gone = {d["pred_uid"] for d in old} - {pu for pu, _, _ in new}
                cur = st.db.execute("SELECT start_floor, start FROM task WHERE uid = ?", (uid,)).fetchone()
                if gone and not tasks[no].children and not cur["start_floor"] and "start" not in ch:
                    log(uid, "start", None, cur["start"])            # 지워도 저절로 당겨지지 않게
                    st.db.execute("UPDATE task SET start_floor = ? WHERE uid = ?", (cur["start"], uid))
            if ch.get("fix_relax"):                                  # 들여올 때 어겨 있던 만큼(느슨함)을 0 으로 → 날짜가 선행을 따름
                total = st.db.execute("SELECT COALESCE(SUM(relax_days), 0) FROM dep WHERE succ_uid = ? AND relax_days > 0", (uid,)).fetchone()[0]
                if total:
                    log(uid, "relax", total, 0)
                    st.db.execute("UPDATE dep SET relax_days = 0 WHERE succ_uid = ?", (uid,))
        n = st.db.execute("SELECT COUNT(*) FROM change WHERE txn_id = ?", (txn,)).fetchone()[0]
        if not n:
            raise EditError("invalid", "바뀐 칸이 없다")
        st.db.execute("UPDATE task SET rev = rev + 1 WHERE uid IN (SELECT DISTINCT task_uid FROM change WHERE txn_id = ?)", (txn,))
        m2 = st.model(cfg)
        check_cycles(cyc0, m2)                                   # 선행이 새 고리를 이루면 거절
        moved = st.recompute(cfg, m2)
        end_after = st.db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]
        edited = [r[0] for r in st.db.execute("SELECT DISTINCT task_uid FROM change WHERE txn_id = ?", (txn,))]
        st.db.execute("UPDATE txn SET note = ? WHERE id = ?", (json.dumps(
            {"tasks": len(edited), "cells": n, "moved": len(moved), "end_before": end_before, "end_after": end_after}, ensure_ascii=False), txn))
        st.db.commit()
    except BaseException:
        st.db.rollback()
        raise
    return {"txn": txn, "rev": txn, "tasks": len(edited), "cells": n, "moved": len(moved), "end_before": end_before, "end_after": end_after}


def history(st: Store, cfg: dict, uid: int | None = None, limit: int = 50) -> list[dict]:
    """최근 저장 묶음 (역할 이름만). uid 를 주면 그 업무가 고쳐진 것만."""
    m = st.model(cfg)
    no_of = {u: no for no, u in m.wbs.uids.items()}
    q = "SELECT * FROM txn WHERE kind != 'import'" + (" AND id IN (SELECT txn_id FROM change WHERE task_uid = ?)" if uid else "") + " ORDER BY id DESC LIMIT ?"
    out = []
    for x in st.db.execute(q, ((uid, limit) if uid else (limit,))):
        chs = []
        for c in st.db.execute("SELECT * FROM change WHERE txn_id = ?" + (" AND task_uid = ?" if uid else "") + " ORDER BY id",
                               ((x["id"], uid) if uid else (x["id"],))):
            no = no_of.get(c["task_uid"])
            t = m.wbs.tasks.get(no) if no else None
            chs.append({"uid": c["task_uid"], "no": no, "wbs": t.wbs if t else "", "name": t.name if t else "(지운 업무)",
                        "field": c["field"], "label": LABEL.get(c["field"], c["field"]), "old": c["old"], "new": c["new"]})
        note = json.loads(x["note"]) if x["note"] and x["note"].startswith("{") else {}
        out.append({"txn": x["id"], "at": x["at"], "role": x["role"], "kind": x["kind"], "kind_ko": KIND_KO.get(x["kind"], x["kind"]),
                    "note": note, "changes": chs})
    return out
