"""줄 편집 (v0.5 S4) — 추가 · 삭제 · 위/아래 · 들여쓰기/내어쓰기 · 풀기를 한 번에 하나, 바로 저장 (docs/05 13장).

  미리 계산  dry_run=True 면 같은 트랜잭션으로 실제로 반영·다시 계산한 뒤 되돌린다 → 확인 창 숫자 = 저장 결과
  권한       관리자: 전부 · 모듈 리더: 자기 모듈(Lv1 아래) 안에서만 (Lv1 줄을 만들거나·옮기거나·지우는 것은 관리자)
  삭제       지우지 않고 deleted = 1 (요약이면 아래 전부). 이 업무를 선행으로 쓰던 후속은 그 연결이 빠지고,
             시작 하한이 없던 하위 업무는 지금 시작을 하한으로 (날짜가 저절로 당겨지지 않게 — S3 선행 지우기와 같음).
             다른 모듈 업무가 걸려 있으면 리더는 못 지운다 (목록을 돌려줌).
  풀기       요약 줄 하나를 없애고 그 아래를 한 레벨 올린다 (Lv 줄이기용). 요약에 걸린 선행은 바로 아래 업무 각각으로,
             요약을 선행으로 쓰던 후속은 바로 아래 업무 각각을 선행으로 (FS 는 같은 뜻, SS·FF·SF 는 근사).
  겹침       내가 본 판 뒤에 다른 저장이 이 업무를 고쳤거나 같은 부모의 줄 순서를 바꿨으면 거절.
  순환       저장 뒤 선행이 새 고리를 이루면 거절 (원래 있던 고리는 그대로 둠).
번호·WBS 코드는 위치로 매번 계산되므로 바뀐다. 선행은 uid 로 이어져 있어 깨지지 않는다.
이력의 선행 글자는 '저장 전 번호' 기준이다.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import date

from .analyze import tarjan
from .auth import Role
from .config import nfkc
from .parse import DEP_RE, parse_deps
from .schedule import ONE, widx
from .store import Store, _dates, _units

OPS = ("add", "delete", "up", "down", "indent", "outdent", "ungroup")
OP_KO = {"add": "추가", "delete": "삭제", "up": "위로", "down": "아래로", "indent": "들여쓰기", "outdent": "내어쓰기", "ungroup": "풀기"}


class EditError(ValueError):
    def __init__(self, kind: str, message: str, items: list | None = None):
        super().__init__(message)
        self.kind = kind            # invalid · denied · conflict
        self.items = items or []


def max_level(cfg: dict) -> int:
    return int(cfg.get("max_level") or max(cfg["levels"]))


def cycle_uids(m) -> set[int]:
    """선행 고리(강연결 요소 크기 2 이상)에 든 업무 uid."""
    adj = defaultdict(list)
    for e in m.edges:
        adj[e.pred].append(e.succ)
    return {m.wbs.uids[n] for c in tarjan(list(m.wbs.tasks), adj) if len(c) > 1 for n in c}


def check_cycles(before: set[int], m) -> None:
    new = cycle_uids(m) - before
    if new:
        no_of = {u: n for n, u in m.wbs.uids.items()}
        items = [{"uid": u, "no": no_of[u], "wbs": m.wbs.tasks[no_of[u]].wbs, "name": m.wbs.tasks[no_of[u]].name}
                 for u in sorted(new, key=no_of.get)]
        raise EditError("invalid", f"선행이 돌고 도는 고리가 생긴다 (업무 {len(new)}개) — 저장하지 않음", items[:50])


def parse_preds(text, units: dict, tasks: dict, self_no: int | None = None) -> list[tuple]:
    """화면 번호 표기('12,15SS+5일') → [(Dep, 번호 뒤 표기)]. 못 읽으면 EditError."""
    text = nfkc(text or "").strip()
    deps, badtok = parse_deps(text, units)
    if badtok:
        raise EditError("invalid", "선행 표기를 읽지 못함: " + ", ".join(badtok))
    miss = [d.pred for d in deps if d.pred not in tasks or d.pred == self_no]
    if miss:
        raise EditError("invalid", "없는 번호·자기 자신: " + ", ".join(map(str, miss)))
    toks = [tok.strip() for tok in re.split(r"[,;]", text) if tok.strip()]
    out = []
    for d, tok in zip(deps, toks):
        mt = DEP_RE.match(tok)
        out.append((d, mt.group(0)[len(mt.group(1)):].strip()))
    return out


def apply_row_op(st: Store, cfg: dict, role: Role, base_rev: int, op: dict, dry_run: bool = False) -> dict:
    if role.kind == "viewer":
        raise EditError("denied", "조회자는 저장할 수 없다")
    if not isinstance(op, dict) or op.get("op") not in OPS:
        raise EditError("invalid", "모르는 줄 동작")
    kind = op["op"]
    m = st.model(cfg)
    tasks, uids = m.wbs.tasks, m.wbs.uids
    no_of = {u: n for n, u in uids.items()}
    uid = op.get("uid")
    if isinstance(uid, bool) or not isinstance(uid, int) or uid not in no_of:
        raise EditError("invalid", f"없는 업무: {uid}")
    no = no_of[uid]
    t = tasks[no]
    cyc0 = cycle_uids(m)

    def mod(n: int) -> int:
        while tasks[n].parent is not None:
            n = tasks[n].parent
        return uids[n]

    def info(n: int) -> dict:
        return {"uid": uids[n], "no": n, "wbs": tasks[n].wbs, "name": tasks[n].name}

    def puid(n: int):
        return uids[tasks[n].parent] if tasks[n].parent is not None else None

    def subtree(n: int) -> list[int]:
        out, stack = [], [n]
        while stack:
            x = stack.pop()
            out.append(x)
            stack.extend(tasks[x].children)
        return out

    def mine(n: int) -> bool:
        return role.is_admin or (tasks[n].level > 1 and mod(n) == role.module_uid)

    def deny(msg: str, items=None):
        raise EditError("denied", msg, items or [info(no)])

    role_mods = {r[0] for r in st.db.execute("SELECT module_uid FROM role WHERE module_uid IS NOT NULL")}
    targets, parents, warn = {uid}, set(), []
    my_mod = mod(no)

    # --- 검사 (쓰기 전에) ---
    if kind == "add":
        where = op.get("where", "after")
        if where not in ("after", "child"):
            raise EditError("invalid", "추가 위치는 after · child")
        if where == "after" and not mine(no):
            deny("이 자리에 업무를 넣을 권한이 없다 (Lv1 을 만드는 것은 관리자)" if t.level == 1 else "내 모듈이 아니다")
        if where == "child":
            if not (role.is_admin or my_mod == role.module_uid):
                deny("내 모듈이 아니다")
            if t.level + 1 > max_level(cfg):
                raise EditError("invalid", f"Lv{max_level(cfg)} 보다 깊게 넣을 수 없다 (설정 max_level)")
            if not t.children:
                warn.append(f"'{t.name}' 이(가) 요약이 된다 — 그 업무의 기간·MH·시작 하한은 일정에 쓰이지 않는다 (값은 남김)")
        name = nfkc(op.get("name") or "").strip()
        if not name or len(name) > 300:
            raise EditError("invalid", "이름이 비었거나 너무 길다")
        dur = op.get("duration", 5)
        if isinstance(dur, bool) or not isinstance(dur, (int, float)) or not 0 <= dur <= 1e5:
            raise EditError("invalid", "기간은 0 이상 숫자 (평일)")
        try:
            start = date.fromisoformat(str(op["start"])) if op.get("start") else (t.start or date.today())
        except ValueError:
            raise EditError("invalid", "시작은 YYYY-MM-DD")
        mh = op.get("man_hours")
        if mh is not None and (isinstance(mh, bool) or not isinstance(mh, (int, float)) or not 0 <= mh <= 1e7):
            raise EditError("invalid", "Man Hours 는 0 이상 숫자")
        preds = parse_preds(op.get("preds"), _units(cfg), tasks)
        new_parent = uid if where == "child" else puid(no)
        if new_parent is not None:
            parents.add(new_parent)
    elif kind == "delete":
        if not mine(no):
            deny("지울 권한이 없다 (Lv1 은 관리자, 그 밖은 자기 모듈만)")
        sub = subtree(no)
        sub_u = {uids[n] for n in sub}
        if sub_u & role_mods:
            raise EditError("invalid", "역할이 걸린 모듈(Lv1)은 지울 수 없다 — 역할을 먼저 정리할 것")
        outs = defaultdict(list)
        for d in st.db.execute(f"SELECT * FROM dep WHERE pred_uid IN ({','.join('?' * len(sub_u))})", tuple(sub_u)):
            if d["succ_uid"] not in sub_u and d["succ_uid"] in no_of:
                outs[d["succ_uid"]].append(d)
        other = [info(no_of[s]) for s in outs if mod(no_of[s]) != my_mod]
        if other and not role.is_admin:
            raise EditError("denied", f"다른 모듈 업무 {len(other)}개가 이 업무(또는 그 아래)를 선행으로 쓰고 있다 — "
                            "관리자나 그 모듈 리더가 연결을 먼저 풀 것", other)
        if other:
            warn.append(f"다른 모듈 업무 {len(other)}개의 선행 연결이 빠진다")
        if len(sub) > 1:
            warn.append(f"아래 업무 {len(sub) - 1}개도 함께 지운다")
        targets |= sub_u | set(outs)
        if puid(no) is not None:
            parents.add(puid(no))
    elif kind in ("up", "down"):
        if not mine(no):
            deny("옮길 권한이 없다")
        if puid(no) is not None:
            parents.add(puid(no))
    elif kind == "indent":
        if not mine(no):
            deny("옮길 권한이 없다")
        sib = tasks[t.parent].children if t.parent is not None else [n for n in tasks if tasks[n].parent is None]
        i = sib.index(no)
        if i == 0:
            raise EditError("invalid", "바로 위 형제가 없어 들여쓸 수 없다")
        if uid in role_mods:
            raise EditError("invalid", "역할이 걸린 모듈(Lv1)은 다른 업무 아래로 넣을 수 없다 — 역할을 먼저 정리할 것")
        deepest = max(tasks[n].level for n in subtree(no))
        if deepest + 1 > max_level(cfg):
            raise EditError("invalid", f"아래 업무가 Lv{max_level(cfg)} 보다 깊어진다 (설정 max_level)")
        prev = sib[i - 1]
        if not tasks[prev].children:
            warn.append(f"'{tasks[prev].name}' 이(가) 요약이 된다 — 그 업무의 기간·MH·시작 하한은 일정에 쓰이지 않는다 (값은 남김)")
        new_parent = uids[prev]
        parents.add(new_parent)
        if puid(no) is not None:
            parents.add(puid(no))
    elif kind == "outdent":
        if t.parent is None:
            raise EditError("invalid", "맨 위 레벨이라 내어쓸 수 없다")
        if not mine(no) or (t.level == 2 and not role.is_admin):
            deny("내어쓰면 Lv1(모듈)이 된다 — 관리자만" if t.level == 2 else "옮길 권한이 없다")
        if len(tasks[t.parent].children) == 1:
            warn.append(f"'{tasks[t.parent].name}' 은(는) 하위가 없어져 하위 업무가 된다 (기간 = 지금 범위, 시작 하한 = 지금 시작)")
        parents.add(puid(no))
        if puid(t.parent) is not None:
            parents.add(puid(t.parent))
    elif kind == "ungroup":
        if not t.children:
            raise EditError("invalid", "하위가 없는 업무는 풀 것이 없다")
        if t.level == 1:
            raise EditError("invalid", "Lv1(모듈)은 풀 수 없다")
        if not mine(no):
            deny("풀 권한이 없다")
        kids_u = [uids[c] for c in t.children]
        outs = [d for d in st.db.execute("SELECT * FROM dep WHERE pred_uid = ?", (uid,)) if d["succ_uid"] in no_of]
        other = [info(no_of[d["succ_uid"]]) for d in outs if mod(no_of[d["succ_uid"]]) != my_mod]
        if other and not role.is_admin:
            raise EditError("denied", f"다른 모듈 업무 {len(other)}개가 이 요약을 선행으로 쓰고 있다 — 관리자만 풀 수 있다", other)
        ins = [d for d in st.db.execute("SELECT * FROM dep WHERE succ_uid = ? ORDER BY ord, id", (uid,)) if d["pred_uid"] in no_of]
        if any(d["type"] != "FS" for d in ins + outs):
            warn.append("FS 가 아닌 연결이 있어 아래 업무로 옮기면 뜻이 조금 달라질 수 있다 (날짜 변화로 확인)")
        if ins or outs:
            warn.append(f"요약에 걸린 선행 {len(ins)}개 · 후속 {len(outs)}개를 바로 아래 업무 {len(kids_u)}개로 옮긴다")
        targets |= set(kids_u) | {d["succ_uid"] for d in outs}
        if puid(no) is not None:
            parents.add(puid(no))

    q1 = ",".join("?" * len(targets))
    q2 = ",".join("?" * len(parents)) or "NULL"
    hit = st.db.execute(
        "SELECT DISTINCT c.task_uid, c.field, x.id, x.at, x.role FROM change c JOIN txn x ON x.id = c.txn_id "
        f"WHERE x.kind IN ('edit', 'rows') AND x.id > ? AND (c.task_uid IN ({q1}) OR (c.field = 'children' AND c.task_uid IN ({q2}))) "
        "ORDER BY x.id", (base_rev, *targets, *parents)).fetchall()
    if hit:
        items = []
        for r in hit:
            i = info(no_of[r["task_uid"]]) if r["task_uid"] in no_of else {"uid": r["task_uid"], "no": None, "wbs": "", "name": "(지운 업무)"}
            items.append(i | {"txn": r["id"], "at": r["at"], "role": r["role"]})
        raise EditError("conflict", "내가 본 판 뒤에 다른 저장이 이 업무나 같은 묶음의 줄 순서를 바꿨다 — 새로 고친 뒤 다시", items)

    # --- 쓰기 (한 트랜잭션, 미리 계산이면 마지막에 되돌림) ---
    db = st.db
    end_before = db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]

    def siblings(p) -> list[int]:
        return [r[0] for r in db.execute("SELECT uid FROM task WHERE parent_uid IS ? AND deleted = 0 ORDER BY pos, uid", (p,))]

    def place(p, order: list[int]) -> None:
        for i, u in enumerate(order):
            db.execute("UPDATE task SET parent_uid = ?, pos = ? WHERE uid = ?", (p, i, u))

    def preds_text(s: int) -> str:
        return ",".join(f"{no_of[d['pred_uid']]}{d['suffix']}" for d in
                        db.execute("SELECT pred_uid, suffix FROM dep WHERE succ_uid = ? ORDER BY ord, id", (s,)) if d["pred_uid"] in no_of)

    def hold_start(s: int) -> None:
        """시작 하한이 없던 하위 업무는 지금 시작을 하한으로 (저절로 당겨지지 않게)."""
        r = db.execute("SELECT start_floor, start FROM task WHERE uid = ?", (s,)).fetchone()
        if not r["start_floor"] and r["start"] and not tasks[no_of[s]].children:
            log(s, "start", None, r["start"])
            db.execute("UPDATE task SET start_floor = ? WHERE uid = ?", (r["start"], s))

    def leafify(p) -> None:
        """하위가 다 빠진 요약 → 하위 업무 (지금 범위를 기간·시작 하한으로)."""
        if p is None or siblings(p):
            return
        r = db.execute("SELECT * FROM task WHERE uid = ?", (p,)).fetchone()
        if r["start"] and r["finish"]:
            s, f = date.fromisoformat(r["start"]), date.fromisoformat(r["finish"])
            wd = float(widx(f + ONE) - widx(s))
            db.execute("UPDATE task SET work_days = COALESCE(work_days, ?), duration_days = COALESCE(duration_days, ?) WHERE uid = ?",
                       (wd, wd, p))
            if not r["start_floor"]:
                log(p, "start", None, r["start"])
                db.execute("UPDATE task SET start_floor = ? WHERE uid = ?", (r["start"], p))

    try:
        txn = st._txn("rows", role.label)

        def log(u, fld, old, new):
            db.execute("INSERT INTO change(txn_id, task_uid, field, old, new) VALUES(?,?,?,?,?)",
                       (txn, u, fld, None if old is None else str(old), None if new is None else str(new)))

        result_uid = uid
        if kind == "add":
            if dur == 0:
                s0 = f0 = start
            else:
                s0, f0 = _dates(widx(start), widx(start) + dur)
            cur = db.execute("INSERT INTO task(parent_uid, pos, name, duration_days, work_days, start_floor, start, finish, man_hours) "
                             "VALUES(?,?,?,?,?,?,?,?,?)", (new_parent, -1, name, float(dur), float(dur), start.isoformat(),
                                                          s0.isoformat(), f0.isoformat(), None if mh is None else float(mh)))
            result_uid = new = cur.lastrowid
            order = [u for u in siblings(new_parent) if u != new]
            order.insert(order.index(uid) + 1 if where == "after" else len(order), new)
            place(new_parent, order)
            for k, (d, suffix) in enumerate(preds):
                db.execute("INSERT INTO dep(succ_uid, pred_uid, ord, type, lag_days, lag_text, suffix, relax_days) VALUES(?,?,?,?,?,?,?,0)",
                           (new, uids[d.pred], k, d.type, d.lag_days, d.lag_text, suffix))
            log(new, "add", None, name)
            if preds:
                log(new, "preds", None, preds_text(new))
            if new_parent is not None:
                log(new_parent, "children", None, f"+ {name}")
        elif kind == "delete":
            for s, rows in outs.items():
                old = preds_text(s)
                db.execute(f"DELETE FROM dep WHERE id IN ({','.join('?' * len(rows))})", tuple(d["id"] for d in rows))
                log(s, "preds", old, preds_text(s))
                hold_start(s)
            for n in sub:
                db.execute("UPDATE task SET deleted = 1 WHERE uid = ?", (uids[n],))
                log(uids[n], "delete", f"{tasks[n].wbs} {tasks[n].name}", None)
            p = puid(no)
            if p is not None:
                log(p, "children", f"- {t.name}", None)
                place(p, siblings(p))
                leafify(p)
        elif kind in ("up", "down"):
            order = siblings(puid(no))
            i = order.index(uid)
            j = i - 1 if kind == "up" else i + 1
            if not 0 <= j < len(order):
                raise EditError("invalid", "더 옮길 수 없다 (맨 " + ("위" if kind == "up" else "아래") + ")")
            order[i], order[j] = order[j], order[i]
            place(puid(no), order)
            if puid(no) is not None:
                log(puid(no), "children", None, f"{OP_KO[kind]}: {t.name}")
        elif kind == "indent":
            old_p = puid(no)
            place(old_p, [u for u in siblings(old_p) if u != uid])
            place(new_parent, [u for u in siblings(new_parent) if u != uid] + [uid])
            for p in (old_p, new_parent):
                if p is not None:
                    log(p, "children", None, f"{OP_KO[kind]}: {t.name}")
        elif kind == "outdent":
            old_p, gp = puid(no), puid(t.parent)
            place(old_p, [u for u in siblings(old_p) if u != uid])
            order = [u for u in siblings(gp) if u != uid]
            order.insert(order.index(old_p) + 1, uid)
            place(gp, order)
            for p in (old_p, gp):
                if p is not None:
                    log(p, "children", None, f"{OP_KO[kind]}: {t.name}")
            leafify(old_p)
        elif kind == "ungroup":
            for c in kids_u:                                     # 요약에 걸린 선행 → 바로 아래 업무 각각
                have = {(d["pred_uid"], d["type"]) for d in db.execute("SELECT pred_uid, type FROM dep WHERE succ_uid = ?", (c,))}
                add = [d for d in ins if (d["pred_uid"], d["type"]) not in have and d["pred_uid"] != c]
                if add:
                    old = preds_text(c)
                    base = db.execute("SELECT COALESCE(MAX(ord), -1) + 1 FROM dep WHERE succ_uid = ?", (c,)).fetchone()[0]
                    for k, d in enumerate(add):
                        db.execute("INSERT INTO dep(succ_uid, pred_uid, ord, type, lag_days, lag_text, suffix, relax_days) VALUES(?,?,?,?,?,?,?,?)",
                                   (c, d["pred_uid"], base + k, d["type"], d["lag_days"], d["lag_text"], d["suffix"], d["relax_days"]))
                    log(c, "preds", old, preds_text(c))
            by_succ = defaultdict(list)
            for d in outs:
                by_succ[d["succ_uid"]].append(d)
            for s, rows in by_succ.items():                      # 요약을 선행으로 쓰던 후속 → 바로 아래 업무 각각을 선행으로
                old = preds_text(s)
                for d in rows:
                    db.execute("DELETE FROM dep WHERE id = ?", (d["id"],))
                    for c in kids_u:
                        if c != s:
                            db.execute("INSERT INTO dep(succ_uid, pred_uid, ord, type, lag_days, lag_text, suffix, relax_days) "
                                       "VALUES(?,?,?,?,?,?,?,?)",
                                       (s, c, d["ord"], d["type"], d["lag_days"], d["lag_text"], d["suffix"], d["relax_days"]))
                log(s, "preds", old, preds_text(s))
            p = puid(no)
            order = siblings(p)
            i = order.index(uid)
            place(p, order[:i] + kids_u + order[i + 1:])
            db.execute("UPDATE task SET deleted = 1 WHERE uid = ?", (uid,))
            log(uid, "ungroup", f"{t.wbs} {t.name}", None)
            if p is not None:
                log(p, "children", None, f"{OP_KO[kind]}: {t.name}")

        m2 = st.model(cfg)
        check_cycles(cyc0, m2)
        no2 = {u: n for n, u in m2.wbs.uids.items()}
        renum = [u for u, n in no_of.items() if u in no2 and (no2[u] != n or m2.wbs.tasks[no2[u]].wbs != tasks[n].wbs)]
        if kind in ("up", "down", "indent", "outdent") and uid in no2:
            log(uid, "move", t.wbs, m2.wbs.tasks[no2[uid]].wbs)
        db.execute("UPDATE task SET rev = rev + 1 WHERE uid IN (SELECT DISTINCT task_uid FROM change WHERE txn_id = ?)", (txn,))
        moved = st.recompute(cfg, m2)

        def mod2(n: int) -> int:
            while m2.wbs.tasks[n].parent is not None:
                n = m2.wbs.tasks[n].parent
            return m2.wbs.uids[n]
        home = mod2(no2[result_uid]) if result_uid in no2 else my_mod
        moved_other = [u for u in moved if u in no2 and mod2(no2[u]) != home]
        end_after = db.execute("SELECT MAX(finish) FROM task WHERE deleted = 0").fetchone()[0]
        n_del = db.execute("SELECT COUNT(*) FROM change WHERE txn_id = ? AND field IN ('delete', 'ungroup')", (txn,)).fetchone()[0]
        res = {"op": kind, "op_ko": OP_KO[kind], "dry_run": dry_run, "uid": result_uid,
               "no": no2.get(result_uid), "wbs": m2.wbs.tasks[no2[result_uid]].wbs if result_uid in no2 else None,
               "renumbered": len(renum), "moved": len(moved), "moved_other": len(moved_other), "deleted": n_del,
               "end_before": end_before, "end_after": end_after, "warnings": warn}
        if dry_run:
            db.rollback()
            return res | {"txn": None, "rev": base_rev}
        db.execute("UPDATE txn SET note = ? WHERE id = ?", (json.dumps(
            {"op": kind, "tasks": 1, "renumbered": len(renum), "moved": len(moved), "end_before": end_before, "end_after": end_after},
            ensure_ascii=False), txn))
        db.commit()
        return res | {"txn": txn, "rev": txn}
    except BaseException:
        db.rollback()
        raise
