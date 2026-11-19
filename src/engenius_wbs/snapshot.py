"""판 확정·판 비교 (v0.5 S5, docs/05 14-1).

  freeze        관리자 — 지금 업무·선행 표를 같은 DB 안 snap_task·snap_dep 에 통째로 복사해 얼린다 (MS Project '기준 계획 저장').
                한 트랜잭션: snapshot 행 + 표 복사 + 요약 + 이력(txn kind 'snapshot', 판 번호는 올리지 않음).
  open_snapshot 옛 판을 메모리 DB(Store)로 — 지금 코드(model·export_xlsx·analyze·report)가 그대로 읽는다.
  compare       두 판(또는 판 ↔ 지금)을 uid 로 맞춰 무엇이 바뀌었나 — 요약·모듈(Lv1)별·업무별.
  compare_xlsx  비교 결과 엑셀 (요약 · 모듈별 · 바뀐 업무 — 업무·항목마다 한 줄).

판은 고치지도 지우지도 않는다. 다시 들여오면 uid 가 새로 매겨져 옛 판과 맞지 않으므로 새 DB 로 옮기지 않는다 (.bak 에 남음).
"""
from __future__ import annotations

import json
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

from .analyze import analyze
from .auth import Role
from .config import nfkc
from .rows import EditError
from .schedule import widx
from .store import Model, Store, module_of

LABEL_MAX = 60
FIELD_KO = {"added": "추가", "deleted": "삭제", "name": "이름", "duration": "기간(평일)", "preds": "선행", "man_hours": "Man Hours",
            "staff": "투입인력", "deliverables": "산출물", "parent": "위치(묶음)", "start_floor": "시작 하한", "dates": "날짜"}
DIRECT = ("name", "duration", "preds", "man_hours", "staff", "deliverables", "parent", "start_floor")


def summary(m: Model, cfg: dict) -> dict:
    """판 한 장의 요약 숫자 (판 목록·비교 머리)."""
    w = m.wbs
    a = analyze(w, cfg)
    sc, wl = a.schedule, a.workload
    issues = Counter(i.code for i in w.issues + a.issues)
    return {"tasks": len(w.tasks), "deps": len(a.edges), "modules": sum(1 for t in w.tasks.values() if t.level == 1),
            "start": sc.start_date.isoformat() if sc and sc.start_date else None,
            "end": sc.end_date.isoformat() if sc and sc.end_date else None,
            "critical": len(sc.critical) if sc else 0, "near": len(sc.near) if sc else 0,
            "mh": round((wl.total_mh + wl.undated_mh) if wl else 0.0, 3), "issues": dict(sorted(issues.items()))}


def _label(label) -> str:
    label = nfkc(label or "").strip()
    if not label or len(label) > LABEL_MAX:
        raise EditError("invalid", f"판 이름은 1~{LABEL_MAX}자")
    return label


def check_label(st: Store, label) -> str:
    label = _label(label)
    if st.db.execute("SELECT 1 FROM snapshot WHERE label = ?", (label,)).fetchone():
        raise EditError("invalid", f"같은 이름의 판이 이미 있다: {label}")
    return label


def freeze(st: Store, cfg: dict, role: Role, label: str, note: str = "") -> dict:
    """→ 판 정보 (list_snapshots 한 줄과 같은 모양). 안 되면 EditError (아무것도 쓰지 않음)."""
    if not role.is_admin:
        raise EditError("denied", "판 확정은 관리자만")
    label = check_label(st, label)
    note = nfkc(note or "").strip()[:500]
    summ = summary(st.model(cfg), cfg)
    rev = st.rev()
    try:
        txn = st._txn("snapshot", role.label, json.dumps({"label": label, "rev": rev}, ensure_ascii=False))
        cur = st.db.execute("INSERT INTO snapshot(label, at, role, rev, txn_id, note, summary) VALUES(?,?,?,?,?,?,?)",
                            (label, datetime.now().isoformat(timespec="seconds"), role.label, rev, txn, note,
                             json.dumps(summ, ensure_ascii=False)))
        sid = cur.lastrowid
        for table in ("task", "dep"):
            cols = ", ".join(st.columns(table))
            st.db.execute(f"INSERT INTO snap_{table}(snapshot_id, {cols}) SELECT ?, {cols} FROM {table}", (sid,))
        st.db.commit()
    except BaseException:
        st.db.rollback()
        raise
    return next(s for s in list_snapshots(st) if s["id"] == sid)


def list_snapshots(st: Store) -> list[dict]:
    out = []
    for r in st.db.execute("SELECT * FROM snapshot ORDER BY id"):
        out.append({"id": r["id"], "label": r["label"], "at": r["at"], "role": r["role"], "rev": r["rev"], "note": r["note"] or "",
                    "summary": json.loads(r["summary"]) if r["summary"] else {}})
    return out


def find_snapshot(st: Store, key) -> int:
    """판 id 또는 이름 → id."""
    r = None
    if isinstance(key, int) or (isinstance(key, str) and key.strip().isdigit()):
        r = st.db.execute("SELECT id FROM snapshot WHERE id = ?", (int(key),)).fetchone()
    if r is None and isinstance(key, str):
        r = st.db.execute("SELECT id FROM snapshot WHERE label = ?", (nfkc(key).strip(),)).fetchone()
    if r is None:
        raise ValueError(f"없는 판: {key}")
    return r[0]


def open_snapshot(st: Store, sid: int) -> Store:
    """옛 판 → 메모리 Store (같은 meta — 서식 틀·칸 위치는 들여올 때 것 그대로)."""
    if not st.db.execute("SELECT 1 FROM snapshot WHERE id = ?", (sid,)).fetchone():
        raise ValueError(f"없는 판: {sid}")
    mem = Store.memory()
    for k, v in st.db.execute("SELECT key, value FROM meta WHERE key NOT IN ('schema_version', 'edit_open', 'edit_note')"):
        mem.db.execute("INSERT OR REPLACE INTO meta(key, value) VALUES(?, ?)", (k, v))
    for table in ("task", "dep"):
        cols = st.columns(table)
        rows = st.db.execute(f"SELECT {', '.join(cols)} FROM snap_{table} WHERE snapshot_id = ?", (sid,)).fetchall()
        mem.db.executemany(f"INSERT INTO {table}({', '.join(cols)}) VALUES({', '.join('?' * len(cols))})", [tuple(r) for r in rows])
    mem.db.commit()
    return mem


# --- 비교 ----------------------------------------------------------------------------------------
def _side(m: Model) -> dict[int, dict]:
    tasks, uids = m.wbs.tasks, m.wbs.uids
    no_of = {u: n for n, u in uids.items()}
    lv1 = module_of(m)
    out = {}
    for no, t in tasks.items():
        r = m.rows[no]
        deps = m.deps.get(no, [])
        out[uids[no]] = {
            "no": no, "wbs": t.wbs, "name": nfkc(r["name"]).strip(), "level": t.level, "leaf": not t.children,
            "parent": r["parent_uid"], "work": r["work_days"], "floor": r["start_floor"], "mh": r["man_hours"],
            "staff": r["staff"], "dl": "; ".join(t.deliverables),
            "start": t.start, "finish": t.finish, "module": uids[lv1[no]], "module_name": tasks[lv1[no]].name,
            "deps": frozenset((d["pred_uid"], d["type"], float(d["lag_days"] or 0.0)) for d in deps),
            "preds": ",".join(f"{no_of.get(d['pred_uid'], '?')}{d['suffix']}" for d in deps)}
    return out


def _fmt(v) -> str | None:
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def _shift(a: date | None, b: date | None) -> int | None:
    if a is None or b is None:
        return None
    return widx(b) - widx(a)


def compare(a: Store, b: Store, cfg: dict) -> dict:
    """a(옛 판) → b(새 판·지금). 업무는 uid 로 맞춘다."""
    ma, mb = a.model(cfg), b.model(cfg)
    A, B = _side(ma), _side(mb)
    sa, sb = summary(ma, cfg), summary(mb, cfg)
    items, counts = [], Counter()
    mods: dict[int, dict] = {}

    def mod_row(x: dict) -> dict:
        return mods.setdefault(x["module"], {"module": x["module_name"], "added": 0, "deleted": 0, "changed": 0, "dated": 0,
                                             "finish_late": 0, "finish_max": 0})

    def item(kind, x, y, changes, ss=None, fs=None):
        base = y or x
        return {"kind": kind, "uid": None, "module": base["module_name"], "no_a": x["no"] if x else None, "no_b": y["no"] if y else None,
                "wbs_a": x["wbs"] if x else None, "wbs_b": y["wbs"] if y else None, "name": base["name"], "level": base["level"],
                "changes": changes, "start_shift": ss, "finish_shift": fs}

    for u, y in B.items():
        x = A.get(u)
        if x is None:
            counts["added"] += 1
            mod_row(y)["added"] += 1
            items.append(item("added", None, y, [{"field": "added", "label": FIELD_KO["added"], "a": None,
                                                  "b": f"{_fmt(y['start'])} ~ {_fmt(y['finish'])}"}]) | {"uid": u})
            continue
        ch = []
        if x["name"] != y["name"]:
            ch.append(("name", x["name"], y["name"]))
        if x["leaf"] and y["leaf"] and (x["work"] or 0.0) != (y["work"] or 0.0):
            ch.append(("duration", _fmt(x["work"]), _fmt(y["work"])))
        if x["deps"] != y["deps"]:
            ch.append(("preds", x["preds"], y["preds"]))
        if (x["mh"] or 0.0) != (y["mh"] or 0.0) and x["leaf"] and y["leaf"]:
            ch.append(("man_hours", _fmt(x["mh"]), _fmt(y["mh"])))
        if (x["staff"] or 0.0) != (y["staff"] or 0.0):
            ch.append(("staff", _fmt(x["staff"]), _fmt(y["staff"])))
        if x["dl"] != y["dl"]:
            ch.append(("deliverables", x["dl"] or None, y["dl"] or None))
        if x["parent"] != y["parent"]:
            pa, pb = A.get(x["parent"]), B.get(y["parent"])
            ch.append(("parent", pa["wbs"] + " " + pa["name"] if pa else "(맨 위)", pb["wbs"] + " " + pb["name"] if pb else "(맨 위)"))
        if x["floor"] != y["floor"] and x["leaf"] and y["leaf"]:
            ch.append(("start_floor", x["floor"], y["floor"]))
        ss, fs = _shift(x["start"], y["start"]), _shift(x["finish"], y["finish"])
        moved = bool(ss) or bool(fs) or (x["start"] is None) != (y["start"] is None)
        if not ch and not moved:
            continue
        mr = mod_row(y)
        for f, _, _ in ch:
            counts[f] += 1
        if ch:
            counts["changed"] += 1
            mr["changed"] += 1
        if moved:
            counts["dates"] += 1
            mr["dated"] += 1
            if fs and fs > 0 and y["leaf"]:
                mr["finish_late"] += 1
                mr["finish_max"] = max(mr["finish_max"], fs)
        changes = [{"field": f, "label": FIELD_KO[f], "a": _fmt(va), "b": _fmt(vb)} for f, va, vb in ch]
        if moved:
            changes.append({"field": "dates", "label": FIELD_KO["dates"], "a": f"{_fmt(x['start'])} ~ {_fmt(x['finish'])}",
                            "b": f"{_fmt(y['start'])} ~ {_fmt(y['finish'])}"})
        items.append(item("changed" if ch else "dates", x, y, changes, ss, fs) | {"uid": u})
    for u, x in A.items():
        if u not in B:
            counts["deleted"] += 1
            mod_row(x)["deleted"] += 1
            items.append(item("deleted", x, None, [{"field": "deleted", "label": FIELD_KO["deleted"],
                                                    "a": f"{_fmt(x['start'])} ~ {_fmt(x['finish'])}", "b": None}]) | {"uid": u})
    order = {k: i for i, k in enumerate(["added", "deleted", "changed", "dates"])}
    items.sort(key=lambda i: (order[i["kind"]], i["no_b"] if i["no_b"] is not None else i["no_a"]))
    ea, eb = (date.fromisoformat(s["end"]) if s["end"] else None for s in (sa, sb))
    codes = sorted(set(sa["issues"]) | set(sb["issues"]))
    return {"a": sa, "b": sb, "end_shift": _shift(ea, eb), "counts": {k: counts.get(k, 0) for k in
            ("added", "deleted", "changed", *DIRECT, "dates")},
            "issues": [{"code": c, "a": sa["issues"].get(c, 0), "b": sb["issues"].get(c, 0)} for c in codes
                       if sa["issues"].get(c, 0) != sb["issues"].get(c, 0)],
            "modules": sorted(mods.values(), key=lambda r: r["module"]), "items": items}


def compare_xlsx(res: dict, path: Path, name_a: str, name_b: str) -> Path:
    """비교 엑셀 — 요약 · 모듈별 · 바뀐 업무(업무·항목마다 한 줄)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    sa, sb, c = res["a"], res["b"], res["counts"]
    rows = [("", name_a, name_b), ("업무 수", sa["tasks"], sb["tasks"]), ("선행 수", sa["deps"], sb["deps"]),
            ("사업 시작", sa["start"], sb["start"]), ("사업 완료", sa["end"], sb["end"]), ("완료 이동 (평일)", None, res["end_shift"]),
            ("주경로 하위 업무", sa["critical"], sb["critical"]), ("준주경로 하위 업무", sa["near"], sb["near"]), ("Man Hours 합", sa["mh"], sb["mh"]),
            (), ("추가", None, c["added"]), ("삭제", None, c["deleted"]), ("고친 업무", None, c["changed"])]
    rows += [(f"  · {FIELD_KO[k]}", None, c[k]) for k in DIRECT]
    rows += [("날짜가 움직인 업무", None, c["dates"]), ()]
    rows += [(f"검사 {i['code']}", i["a"], i["b"]) for i in res["issues"]]
    for r in rows:
        ws.append(list(r))
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    wm = wb.create_sheet("모듈별")
    wm.append(["모듈 (Lv1)", "추가", "삭제", "고친 업무", "날짜 움직임", "완료가 늦어진 하위 업무", "가장 늦어진 평일"])
    for r in res["modules"]:
        wm.append([r["module"], r["added"], r["deleted"], r["changed"], r["dated"], r["finish_late"], r["finish_max"]])
    wm.column_dimensions["A"].width = 24
    wi = wb.create_sheet("바뀐 업무")
    wi.append(["구분", "모듈", f"번호 {name_a}", f"번호 {name_b}", f"WBS {name_a}", f"WBS {name_b}", "Lv", "이름", "무엇이",
               name_a, name_b, "시작 이동(평일)", "완료 이동(평일)"])
    kind_ko = {"added": "추가", "deleted": "삭제", "changed": "고침", "dates": "날짜만"}
    for i in res["items"]:
        for ch in i["changes"]:
            wi.append([kind_ko[i["kind"]], i["module"], i["no_a"], i["no_b"], i["wbs_a"], i["wbs_b"], i["level"], i["name"], ch["label"],
                       ch["a"], ch["b"], i["start_shift"], i["finish_shift"]])
    for col, w in zip("ABCDEFGHIJKLM", (8, 18, 9, 9, 12, 12, 4, 40, 12, 28, 28, 10, 10)):
        wi.column_dimensions[col].width = w
    wi.freeze_panes = "A2"
    wi.auto_filter.ref = f"A1:M{wi.max_row}"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
