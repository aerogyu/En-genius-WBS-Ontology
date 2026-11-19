"""일괄 조작 (v0.6) — 일정 옮기기(쉬프트). 범위의 하위 업무를 n 평일 통째로 늦추거나 당긴다.

  범위     전체 사업 (uid 없음 — 관리자만) · 업무 하나(요약이면 그 아래 전부) — 관리자, 또는 자기 모듈 안이면 모듈 리더 (자기 Lv1 포함)
  반영     시작 하한(start_floor)만 고친다 → 날짜는 store.recompute(schedule.forward) 가 다시 짠다. 규칙은 schedule.shift_floors.
  이력     txn(kind 'edit', note.shift) + 하한이 바뀐 업무마다 change('start'). 같은 업무의 다른 저장과 겹침 검사는 칸 고치기와 같다.
  미리     dry_run = 실제로 해 보고 되돌림 (화면은 JS 가 같은 규칙으로 미리 그린다 — 저장 결과 = 미리 보기)
"""
from __future__ import annotations

import json

from .auth import Role
from .rows import EditError, check_open
from .schedule import EPS, event_values, finish_date, shift_floors, start_date
from .store import Store, _iso

MAX_DAYS = 5000


def shift(st: Store, cfg: dict, role: Role, base_rev: int, uid: int | None, days, dry_run: bool = False) -> dict:
    """→ {txn, rev, scope, days, leaves, floors, moved, held, end_before, end_after, start_before, start_after}. 안 되면 EditError."""
    if role.kind == "viewer":
        raise EditError("denied", "조회자는 저장할 수 없다")
    check_open(st, role)
    if isinstance(days, bool) or not isinstance(days, int) or days == 0 or abs(days) > MAX_DAYS:
        raise EditError("invalid", f"옮길 평일 수는 0 이 아닌 정수 (±{MAX_DAYS} 안)")
    m = st.model(cfg)
    tasks, uids = m.wbs.tasks, m.wbs.uids
    no_of = {u: n for n, u in uids.items()}
    if uid is None:
        if not role.is_admin:
            raise EditError("denied", "전체 일정 옮기기는 관리자만")
        scope_nos = list(tasks)
        scope_name = "전체"
    else:
        if isinstance(uid, bool) or not isinstance(uid, int) or uid not in no_of:
            raise EditError("invalid", f"없는 업무: {uid}")
        top = no_of[uid]
        root = top
        while tasks[root].parent is not None:
            root = tasks[root].parent
        if not role.is_admin and uids[root] != role.module_uid:
            raise EditError("denied", "내 모듈이 아닌 업무", [{"uid": uid, "no": top, "wbs": tasks[top].wbs, "name": tasks[top].name}])
        scope_nos, stack = [], [top]
        while stack:
            x = stack.pop()
            scope_nos.append(x)
            stack.extend(tasks[x].children)
        scope_name = f"{tasks[top].wbs} {tasks[top].name}"
    val0, _ = event_values(tasks)
    leaves = [no for no in scope_nos if not tasks[no].children and 2 * no in val0]
    bare = [no for no in scope_nos if tasks[no].children and 2 * no in val0 and not any(2 * c in val0 for c in tasks[no].children)]
    if not leaves and not bare:
        raise EditError("invalid", "옮길 날짜가 있는 업무가 없다")
    # 내가 본 판 뒤에 다른 저장이 범위 안 업무를 고쳤으면 거절 (미리 본 것과 달라지므로)
    scope_uids = {uids[no] for no in scope_nos}
    hit = [r for r in st.db.execute(
        "SELECT DISTINCT c.task_uid, x.id, x.at, x.role FROM change c JOIN txn x ON x.id = c.txn_id "
        "WHERE x.kind IN ('edit', 'rows') AND c.field != 'children' AND x.id > ? ORDER BY x.id", (base_rev,)) if r["task_uid"] in scope_uids]
    if hit:
        items = [{"uid": r["task_uid"], "no": no_of[r["task_uid"]], "wbs": tasks[no_of[r["task_uid"]]].wbs,
                  "name": tasks[no_of[r["task_uid"]]].name, "txn": r["id"], "at": r["at"], "role": r["role"]} for r in hit[:50]]
        raise EditError("conflict", f"내가 본 판 뒤에 다른 저장이 범위 안 업무 {len({r['task_uid'] for r in hit})}개를 고쳤다", items)

    new, val, held = shift_floors(tasks, m.edges, val0, m.floor, m.work, m.relax, leaves, days)

    def first(v):
        return _iso(start_date(min(v[2 * no] for no in leaves))) if leaves else None

    end_before = st.db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]
    try:
        txn = st._txn("edit", role.label)
        n_floor = 0
        for no in leaves:
            a, b = m.floor.get(no), new.get(no)
            if b is None or (a is not None and abs(a - b) < EPS):
                continue
            t, r = tasks[no], m.rows[no]
            d = _iso(finish_date(b) if t.is_milestone else start_date(b))
            st.db.execute("INSERT INTO change(txn_id, task_uid, field, old, new) VALUES(?,?,?,?,?)", (txn, r["uid"], "start", r["start_floor"], d))
            st.db.execute("UPDATE task SET start_floor = ?, rev = rev + 1 WHERE uid = ?", (d, r["uid"]))
            n_floor += 1
        for no in bare:                                           # 날짜 있는 하위 업무가 없는 요약 — 적힌 날짜를 그대로 옮긴다
            r = m.rows[no]
            s, f = _iso(start_date(val0[2 * no] + days)), _iso(finish_date(val0[2 * no + 1] + days))
            st.db.execute("INSERT INTO change(txn_id, task_uid, field, old, new) VALUES(?,?,?,?,?)",
                          (txn, r["uid"], "shift", f"{r['start']} ~ {r['finish']}", f"{s} ~ {f}"))
            st.db.execute("UPDATE task SET start = ?, finish = ?, rev = rev + 1 WHERE uid = ?", (s, f, r["uid"]))
        if not n_floor and not bare:
            raise EditError("invalid", "옮겨지는 업무가 없다 (범위 밖 선행에 모두 묶여 있음)")
        moved = st.recompute(cfg, st.model(cfg))
        end_after = st.db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]
        res = {"txn": txn, "rev": txn, "scope": scope_name, "days": days, "leaves": len(leaves), "floors": n_floor,
               "moved": len(moved) + len(bare), "held": len(held), "end_before": end_before, "end_after": end_after,
               "start_before": first(val0), "start_after": first(val)}
        st.db.execute("UPDATE txn SET note = ? WHERE id = ?", (json.dumps(
            {"shift": {"uid": uid, "scope": scope_name, "days": days, "leaves": len(leaves), "held": len(held)},
             "tasks": len(leaves), "cells": n_floor, "moved": res["moved"], "end_before": end_before, "end_after": end_after}, ensure_ascii=False), txn))
        if dry_run:
            st.db.rollback()
            res.update(txn=None, rev=base_rev, dry_run=True)
        else:
            st.db.commit()
    except BaseException:
        st.db.rollback()
        raise
    return res
