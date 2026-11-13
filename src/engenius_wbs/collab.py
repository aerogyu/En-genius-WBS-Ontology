"""수정 요청 · 편집 열림/닫힘 (v0.5 S5, docs/05 14-3·14-5).

수정 요청 — 리더 A 가 'B 모듈에 내 업무의 선행이 될 업무가 없음' 등을 남기면 B 리더(또는 관리자)가 처리·반려, A 는 거둘 수 있다.
  보내기   관리자·모듈 리더 (조회자는 못 함 — 공용 '조회자 (Windows)' 하나라 받는 쪽이 누가 보냈는지 모름). 리더는 자기 모듈로는 못 보냄.
  대상     task_uid(고쳐 달라는 업무, 받는 모듈 = 그 업무의 Lv1) 또는 to_module(Lv1) 만 · ref_uid = 요청한 쪽 업무(리더는 자기 모듈 것)
  처리     받는 모듈 리더·관리자: done(답 글 + 만든·고친 업무 result_uid) · rejected(이유 필수). 보낸 역할·관리자: withdrawn.
  요청은 일정을 바꾸지 않으므로 txn 이 아니고 판 번호를 올리지 않는다. 글은 서버 DB 에만 (엑셀·보고서·TTL 에 안 나감).
편집 열림/닫힘 — meta edit_open(기본 열림)·edit_note. 닫히면 관리자 외 저장·줄 편집 거절(rows.check_open), 요청은 계속 받음.
"""
from __future__ import annotations

import json
from datetime import datetime

from .auth import Role
from .config import nfkc
from .rows import EditError
from .store import Store, module_of

KIND_KO = {"pred": "선행이 될 업무가 없음", "fix": "고쳐 주세요", "etc": "기타"}
STATUS_KO = {"open": "열림", "done": "처리함", "rejected": "반려", "withdrawn": "거둠"}
TEXT_MAX = 500


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _text(v, what: str, required: bool = True) -> str:
    v = nfkc(v or "").strip()
    if required and not v:
        raise EditError("invalid", f"{what}을 적을 것")
    if len(v) > TEXT_MAX:
        raise EditError("invalid", f"{what}은 {TEXT_MAX}자까지")
    return v


def _uid(v, what: str):
    if v is None or v == "":
        return None
    if isinstance(v, bool) or not isinstance(v, int):
        raise EditError("invalid", f"{what} 는 업무 uid (숫자)")
    return v


def create_request(st: Store, cfg: dict, role: Role, body: dict) -> dict:
    if role.kind == "viewer":
        raise EditError("denied", "조회자는 수정 요청을 보낼 수 없다 (역할을 받을 것)")
    if not isinstance(body, dict):
        raise EditError("invalid", "요청을 읽지 못함")
    kind = body.get("kind")
    if kind not in KIND_KO:
        raise EditError("invalid", "종류는 pred·fix·etc")
    text = _text(body.get("text"), "요청 글")
    m = st.model(cfg)
    uids = m.wbs.uids
    no_of = {u: n for n, u in uids.items()}
    lv1 = module_of(m)
    task_uid, ref_uid, to_module = _uid(body.get("task_uid"), "task_uid"), _uid(body.get("ref_uid"), "ref_uid"), _uid(body.get("to_module"), "to_module")
    for u, what in ((task_uid, "고쳐 달라는 업무"), (ref_uid, "요청한 쪽 업무")):
        if u is not None and u not in no_of:
            raise EditError("invalid", f"{what}가 없다 (지웠거나 잘못된 uid {u})")
    if task_uid is not None:
        mod = uids[lv1[no_of[task_uid]]]
        if to_module is not None and to_module != mod:
            raise EditError("invalid", "받는 모듈이 그 업무의 모듈과 다르다")
        to_module = mod
    mods = {uids[n] for n in lv1.values()}
    if to_module is None or to_module not in mods:
        raise EditError("invalid", "받는 모듈(Lv1)을 고를 것")
    if role.kind == "leader":
        if to_module == role.module_uid:
            raise EditError("invalid", "내 모듈은 직접 고칠 수 있다 (요청은 다른 모듈로)")
        if ref_uid is not None and uids[lv1[no_of[ref_uid]]] != role.module_uid:
            raise EditError("invalid", "요청한 쪽 업무는 내 모듈 업무여야 한다")
    cur = st.db.execute("INSERT INTO request(at, from_role, from_module, to_module, task_uid, ref_uid, kind, text) VALUES(?,?,?,?,?,?,?,?)",
                        (_now(), role.label, role.module_uid if role.kind == "leader" else None, to_module, task_uid, ref_uid, kind, text))
    st.db.commit()
    return _one(st, cfg, role, cur.lastrowid)


def act_request(st: Store, cfg: dict, role: Role, rid: int, action: str, reply: str = "", result_uid=None) -> dict:
    r = st.db.execute("SELECT * FROM request WHERE id = ?", (rid,)).fetchone()
    if r is None:
        raise EditError("invalid", f"없는 요청: {rid}")
    if r["status"] != "open":
        raise EditError("conflict", f"이미 닫힌 요청이다 ({STATUS_KO[r['status']]} · {r['closed_role'] or ''} {r['closed_at'] or ''})")
    if action in ("done", "rejected"):
        if not (role.is_admin or (role.kind == "leader" and role.module_uid == r["to_module"])):
            raise EditError("denied", "받는 모듈의 리더나 관리자만 처리한다")
        reply = _text(reply, "반려 이유" if action == "rejected" else "답 글", required=action == "rejected")
        result_uid = _uid(result_uid, "result_uid") if action == "done" else None
        if result_uid is not None and not st.db.execute("SELECT 1 FROM task WHERE uid = ? AND deleted = 0", (result_uid,)).fetchone():
            raise EditError("invalid", f"결과 업무가 없다 (uid {result_uid})")
    elif action == "withdrawn":
        if not (role.is_admin or r["from_role"] == role.label):
            raise EditError("denied", "보낸 역할이나 관리자만 거둔다")
        reply, result_uid = _text(reply, "거두는 이유", required=False), None
    else:
        raise EditError("invalid", "action 은 done·rejected·withdrawn")
    st.db.execute("UPDATE request SET status = ?, closed_at = ?, closed_role = ?, reply = ?, result_uid = ? WHERE id = ? AND status = 'open'",
                  (action, _now(), role.label, reply or None, result_uid, rid))
    st.db.commit()
    return _one(st, cfg, role, rid)


def _visible(role: Role) -> tuple[str, tuple]:
    if role.is_admin:
        return "1", ()
    if role.kind == "leader":
        return "(to_module = ? OR from_role = ?)", (role.module_uid, role.label)
    return "0", ()


def list_requests(st: Store, cfg: dict, role: Role, rid: int | None = None) -> list[dict]:
    """이 역할이 볼 요청 (관리자 = 전부, 리더 = 받은 것·보낸 것, 조회자 = 없음). 열린 것 먼저, 최근 것 먼저."""
    where, args = _visible(role)
    if rid is not None:
        where, args = f"{where} AND id = ?", (*args, rid)
    rows = st.db.execute(f"SELECT * FROM request WHERE {where} ORDER BY status != 'open', id DESC", args).fetchall()
    if not rows:
        return []
    m = st.model(cfg)
    tasks, uids = m.wbs.tasks, m.wbs.uids
    no_of = {u: n for n, u in uids.items()}
    names = {r["uid"]: r["name"].strip() for r in st.db.execute("SELECT uid, name FROM task")}

    def info(u):
        if u is None:
            return None
        no = no_of.get(u)
        if no is None:
            return {"uid": u, "no": None, "wbs": "", "name": f"(지운 업무) {names.get(u, '')}".strip()}
        return {"uid": u, "no": no, "wbs": tasks[no].wbs, "name": tasks[no].name}

    out = []
    for r in rows:
        inbox = role.is_admin or (role.kind == "leader" and r["to_module"] == role.module_uid)
        mine = r["from_role"] == role.label
        out.append({"id": r["id"], "at": r["at"], "from_role": r["from_role"], "from_module": names.get(r["from_module"]),
                    "to_module": r["to_module"], "to_module_name": names.get(r["to_module"], ""),
                    "task": info(r["task_uid"]), "ref": info(r["ref_uid"]), "result": info(r["result_uid"]),
                    "kind": r["kind"], "kind_ko": KIND_KO.get(r["kind"], r["kind"]), "text": r["text"],
                    "status": r["status"], "status_ko": STATUS_KO[r["status"]], "closed_at": r["closed_at"], "closed_role": r["closed_role"],
                    "reply": r["reply"], "inbox": inbox, "outbox": mine,
                    "can_act": r["status"] == "open" and inbox, "can_withdraw": r["status"] == "open" and (mine or role.is_admin)})
    return out


def _one(st: Store, cfg: dict, role: Role, rid: int) -> dict:
    got = list_requests(st, cfg, role, rid)
    if got:
        return got[0]
    return {"id": rid}


def inbox_count(st: Store, role: Role) -> int:
    if role.is_admin:
        return st.db.execute("SELECT COUNT(*) FROM request WHERE status = 'open'").fetchone()[0]
    if role.kind == "leader":
        return st.db.execute("SELECT COUNT(*) FROM request WHERE status = 'open' AND to_module = ?", (role.module_uid,)).fetchone()[0]
    return 0


# --- 편집 열림/닫힘 --------------------------------------------------------------------------------
def edit_state(st: Store) -> dict:
    return {"open": bool(st.meta("edit_open", True)), "note": st.meta("edit_note", "") or ""}


def set_edit_open(st: Store, role: Role, open_: bool, note: str = "") -> dict:
    if not role.is_admin:
        raise EditError("denied", "편집 열고 닫기는 관리자만")
    if not isinstance(open_, bool):
        raise EditError("invalid", "open 은 true·false")
    note = _text(note, "안내 글", required=False)[:200]
    try:
        st.set_meta("edit_open", open_)
        st.set_meta("edit_note", "" if open_ else note)
        st._txn("lock", role.label, json.dumps({"open": open_, "note": note}, ensure_ascii=False))
        st.db.commit()
    except BaseException:
        st.db.rollback()
        raise
    return edit_state(st)
