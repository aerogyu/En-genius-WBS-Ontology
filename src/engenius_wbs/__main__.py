"""python -m engenius_wbs <명령>

    build <WBS.xlsx> [--sheet 이름] [--out 폴더] [--config wbs.yaml] [--title 제목]
        → <폴더>/report.xlsx · wbs_view.html · wbs.ttl · derived.ttl · run_log.json (기본 output/<시각>/)
    synth [--out data/synthetic/wbs_synthetic.xlsx] [--rows 12827] [--seed 7]
        → 가짜 WBS 엑셀 (실데이터 없이 시험할 때)
    apply <원본 WBS.xlsx> <변경 목록.json> [--out 폴더] [--config wbs.yaml] [--force] [--no-build]
        → <원본>_수정_v01.xlsx (원본 형식 그대로, 원본은 안 건드림) + .changes.json, 이어서 그 파일로 build
    import <WBS.xlsx> [--db data/real/wbs.db] [--sheet 이름] [--config wbs.yaml] [--force]
        → 도구 DB (SQLite 파일 하나, v0.5 S1). 업무마다 고유 ID. 원본과 달라지는 칸은 <db>.import_notes.json 에
    export [--db data/real/wbs.db] [--out 파일.xlsx] [--config wbs.yaml]
        → DB 를 원본 양식 엑셀로 (기본 output/<시각>/)
    role add <로그인ID> --label 역할이름 --kind admin|leader|viewer [--module Lv1 이름·WBS 코드] [--db]   (비밀번호는 물어봄)
    role list | role passwd <로그인ID> | role off <로그인ID> | role on <로그인ID>
        → 역할 계정 (v0.5 S2). 역할 이름만 이력에 남는다 — 사람 이름을 쓰지 않는다
    account set <DOMAIN\\계정> --role <역할 로그인ID> | account list | account remove <DOMAIN\\계정>
        → Windows 계정 → 역할 (서버 DB 에만). 표에 없는 계정은 설정 windows_login.unregistered 에 따라 조회자/거절
    serve [--db data/real/wbs.db] [--host 127.0.0.1] [--port 8100] [--title 이름] [--cert 파일 --key 파일]
        → 읽기 서버 (v0.5 S2) — 브라우저로 http://<서버>:<포트>/ (starlette·uvicorn 필요)
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

from .config import REPO_ROOT, load_config


def cmd_build(a) -> int:
    from .analyze import analyze
    from .parse import read_wbs
    from .rdfout import write_ttl
    from .report import all_issues, write_report
    from .viz import write_html

    src = Path(a.xlsx)
    if not src.exists():
        print(f"[NG] 파일이 없다: {src}")
        return 2
    cfg = load_config(a.config)
    stamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    out = Path(a.out) if a.out else REPO_ROOT / "output" / stamp
    out.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    w = read_wbs(src, cfg, a.sheet)
    if not w.tasks:
        for i in w.issues:
            print(f"[NG] {i.code} {i.title} {i.detail}")
        return 2
    t1 = time.perf_counter()
    an = analyze(w, cfg)
    t2 = time.perf_counter()
    rep = write_report(w, an, cfg, out / "report.xlsx")
    view = write_html(w, an, out / "wbs_view.html", a.title or "WBS 선후 관계", cfg)
    ttl = write_ttl(w, an, cfg, out, f"WRUN-{stamp}")
    t3 = time.perf_counter()
    issues = all_issues(w, an)
    cnt = Counter(i.code for i in issues)
    log = {"run": f"WRUN-{stamp}", "source_file": src.name, "sheet": w.sheet, "tasks": len(w.tasks), "edges": len(an.edges),
           "cycles": len(an.cycles), "project_workdays": round(an.project_days, 1), "issues": dict(sorted(cnt.items())),
           "levels": dict(sorted(Counter(t.level for t in w.tasks.values()).items())),
           "milestones": sum(1 for t in w.tasks.values() if t.is_milestone),
           "man_hours": {"rows_used": len(an.workload.contributors), "by_level": an.workload.by_level,
                         "months": len(an.workload.months), "undated_rows": len(an.workload.undated),
                         "leaves_without_mh": an.workload.leaves_without_mh},
           "critical_path": {"critical": len(an.schedule.critical), "near": len(an.schedule.near),
                             "negative_gaps": sum(1 for g in an.schedule.gaps if g is not None and g < -1e-6),
                             "excluded": len(an.schedule.excluded), "workdays": round(an.schedule.end - an.schedule.start, 1),
                             "next_float": an.schedule.next_float},
           "seconds": {"read": round(t1 - t0, 2), "analyze": round(t2 - t1, 2), "write": round(t3 - t2, 2)}}
    (out / "run_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"[build] {src.name} ({w.sheet}) → {out}")
    sc = an.schedule
    print(f"  · 업무 {len(w.tasks):,} · 선행 관계 {len(an.edges):,} · 순환 {len(an.cycles)} · 사업 {sc.start_date} ~ {sc.end_date}"
          f" · 주경로 {len(sc.critical)} · 준주경로 {len(sc.near)} (여유 ≤ {sc.near_days:g} 평일)")
    print("  · 문제: " + (" · ".join(f"{k} {v}" for k, v in sorted(cnt.items())) or "없음"))
    print(f"  · 보고서 {rep.name} · 보기 {view.name} · {ttl[0].name} · {ttl[1].name}  ({t3 - t0:.1f}초)")
    return 0


def cmd_apply(a) -> int:
    from .apply import apply_changes
    xs = [Path(f) for f in a.files if Path(f).suffix.lower() == ".xlsx"]
    js = [Path(f) for f in a.files if Path(f).suffix.lower() == ".json"]
    if len(xs) != 1 or len(js) != 1:
        print("[NG] 원본 WBS 엑셀(.xlsx) 하나와 변경 목록(.json) 하나를 주세요 (순서 상관없음)")
        return 2
    for f in xs + js:
        if not f.exists():
            print(f"[NG] 파일이 없다: {f}")
            return 2
    cfg = load_config(a.config)
    try:
        r = apply_changes(xs[0], js[0], cfg, a.out, a.force)
    except ValueError as e:
        print(f"[NG] {e}")
        return 2
    print(f"[apply] {xs[0].name} + {js[0].name} → {r.out}")
    print(f"  · 고친 칸 {r.cells:,} · 날짜가 바뀐 업무 {len(r.moved):,} · 사업 완료 {r.end_before} → {r.end_after}")
    if r.kept_summaries:
        print(f"  · 엑셀 요약 날짜가 하위 범위와 달라(C11) 그대로 둔 요약 {len(r.kept_summaries)}: " + ", ".join(map(str, r.kept_summaries[:10])))
    if r.mismatches:
        print(f"  · [주의] 보기 HTML 계산과 다른 것 {len(r.mismatches)}: " + " / ".join(r.mismatches[:5]))
    elif r.checked:
        print("  · 보기 HTML 계산과 일치 ✓")
    else:
        print("  · (변경 목록에 보기 결과가 없어 대조하지 않음)")
    if a.no_build:
        return 0
    b = argparse.Namespace(xlsx=str(r.out), sheet=None, out=None, config=a.config, title=None)
    return cmd_build(b)


COL_KO = {"row_no": "번호", "wbs": "WBS", "name": "이름", "duration": "기간", "start": "시작", "finish": "완료", "preds": "선행", "succs": "후속",
          "staff": "투입인력", "man_hours": "Man Hours", "deliverable": "산출물"}


def _old_access(db: Path) -> tuple[int, int, list[dict], list[dict]]:
    """덮어쓸 DB 의 업무 수 · 저장(편집) 수 · 역할 · Windows 계정 표 — 새 DB 로 옮기려고 읽어 둔다."""
    from . import auth
    from .store import Store

    with Store(db) as st:
        auth.ensure(st.db)
        n = st.count()
        edits = st.db.execute("SELECT COUNT(*) FROM txn WHERE kind = 'edit'").fetchone()[0]
        roles = [dict(r) | {"module_name": (st.db.execute("SELECT name FROM task WHERE uid = ?", (r["module_uid"],)).fetchone() or [None])[0]}
                 for r in st.db.execute("SELECT * FROM role WHERE login != ?", (auth.WIN_VIEWER_LOGIN,))]
        win = st.db.execute("SELECT active FROM role WHERE login = ?", (auth.WIN_VIEWER_LOGIN,)).fetchone()
        roles.append({"login": auth.WIN_VIEWER_LOGIN, "active": win[0] if win else 1, "builtin": True})
        accounts = [dict(r) for r in st.db.execute("SELECT a.account, r.login FROM account a JOIN role r ON r.id = a.role_id")]
    return n, edits, roles, accounts


def _carry_access(st, roles: list[dict], accounts: list[dict]) -> list[str]:
    """역할(비밀번호 해시 그대로)·계정 표를 새 DB 에. 모듈 리더의 모듈은 Lv1 이름으로 다시 찾는다 → 못 찾은 역할 이름 목록."""
    from . import auth

    auth.ensure(st.db)
    lv1 = {}
    for r in st.db.execute("SELECT uid, name FROM task WHERE parent_uid IS NULL AND deleted = 0"):
        lv1.setdefault(r["name"].strip(), []).append(r["uid"])
    lost = []
    for r in roles:
        if r.get("builtin"):
            st.db.execute("UPDATE role SET active = ? WHERE login = ?", (r["active"], r["login"]))
            continue
        mod = None
        if r["kind"] == "leader":
            hit = lv1.get((r["module_name"] or "").strip(), [])
            mod = hit[0] if len(hit) == 1 else None
            if mod is None:
                lost.append(r["label"])
        st.db.execute("INSERT INTO role(login, label, kind, module_uid, pw, active, created) VALUES(?,?,?,?,?,?,?)",
                      (r["login"], r["label"], r["kind"], mod, r["pw"], r["active"], r["created"]))
    for x in accounts:
        rid = st.db.execute("SELECT id FROM role WHERE login = ?", (x["login"],)).fetchone()
        if rid:
            st.db.execute("INSERT INTO account(account, role_id, created) VALUES(?,?,datetime('now'))", (x["account"], rid[0]))
    st.db.commit()
    return lost


def cmd_import(a) -> int:
    from .store import Store

    src, db = Path(a.xlsx), Path(a.db)
    if not src.exists():
        print(f"[NG] 파일이 없다: {src}")
        return 2
    cfg = load_config(a.config)
    old = None
    if db.exists():
        n, edits, roles, accounts = _old_access(db)
        if n:
            print(f"[import] 이미 업무 {n:,}개가 든 DB 가 있다: {db}")
            if edits:
                print(f"  · 이 DB 에 서버에서 저장한 편집이 {edits}번 있다 — 새로 들여오면 화면에서는 사라진다 (옛 DB 파일에는 남음)")
            print("  · 덮어쓰면 옛 DB 는 .bak-<시각> 으로 이름을 바꿔 남기고, 역할·Windows 계정 표는 새 DB 로 옮긴다")
            ok = a.force
            if not ok and a.ask and sys.stdin and sys.stdin.isatty():
                ok = input("  덮어쓸까요? (y/N): ").strip().lower() in ("y", "yes", "예", "ㅛ")
            if not ok:
                print("[import] 그만둠 — 덮어쓰려면 --force (import.bat 은 물어본다)")
                return 2
            old = (roles, accounts)
    t0 = time.perf_counter()
    tmp = db.with_name(db.name + ".new")
    if tmp.exists():
        tmp.unlink()
    try:
        with Store(tmp) as st:
            rep = st.import_xlsx(src, cfg, a.sheet)
            notes = [dict(r) for r in st.db.execute(
                "SELECT n.task_uid AS uid, t.src_no AS no, t.src_row AS excel_row, n.col, n.old, n.new FROM note n JOIN task t ON t.uid = n.task_uid ORDER BY n.id")]
            lost = _carry_access(st, *old) if old else []
    except Exception as e:                                # 엑셀이 아니거나 망가진 파일 포함
        print(f"[NG] 들여오지 못함: {type(e).__name__}: {e}  (기존 DB 는 그대로)")
        if tmp.exists():
            tmp.unlink()
        return 2
    if db.exists():                                       # 새 DB 가 다 만들어진 뒤에만 옛 것을 옮긴다
        bak = db.with_name(f"{db.name}.bak-{datetime.now().strftime('%Y%m%dT%H%M%S')}")
        db.rename(bak)
        print(f"[import] 옛 DB → {bak.name}")
    tmp.rename(db)
    if old:
        print(f"  · 역할 {len([r for r in old[0] if not r.get('builtin')])} · Windows 계정 {len(old[1])} 를 새 DB 로 옮김"
              + (f" — 모듈을 못 찾은 리더: {', '.join(lost)} (role 로 다시 정할 것)" if lost else ""))
    out = db.with_name(db.stem + ".import_notes.json")
    out.write_text(json.dumps({"source": src.name, "db": db.name, "tasks": rep.tasks, "deps": rep.deps, "moved": rep.moved,
                               "dropped_columns": rep.dropped_columns, "kept_columns": rep.kept_columns, "skipped_rows": rep.skipped_rows,
                               "notes": notes}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"[import] 업무 {rep.tasks:,} · 선행 {rep.deps:,} → {db}  ({time.perf_counter() - t0:.1f}초)")
    print(f"  · 날짜가 움직인 하위 업무 {rep.moved} (0 이어야 한다) — 시작 하한 {rep.floors:,} · 느슨하게 둔 선행 {rep.relaxed:,}")
    if rep.notes:
        print("  · 원본과 달라진 칸: " + " · ".join(f"{COL_KO.get(k, k)} {v}" for k, v in rep.notes.items()) + f"  → {out.name}")
    else:
        print("  · 원본과 달라진 칸 없음")
    if rep.dropped_columns:
        print("  · 버린 칸 (설정 keep_columns 에 헤더를 적으면 값 그대로 보관): " + " · ".join(f"{h}({n})" for h, n in rep.dropped_columns.items()))
    if rep.kept_columns:
        print("  · 그대로 보관한 칸: " + " · ".join(rep.kept_columns))
    print(f"  · 인명 칸 {rep.person_columns}개는 읽지 않음 (내보낼 때 빈 칸)"
          + (f" · 읽지 못한 행 {rep.skipped_rows}" if rep.skipped_rows else "")
          + (f" · 헤더 위 글자 {rep.cleared_top_cells}칸은 서식 틀에서 지움" if rep.cleared_top_cells else ""))
    if rep.moved:
        print("[NG] 날짜가 움직였다 — 이 DB 는 쓰지 말고 알려 줄 것")
        return 1
    return 0


def cmd_export(a) -> int:
    from .store import Store

    db = Path(a.db)
    if not db.exists():
        print(f"[NG] DB 가 없다: {db}  (먼저 import)")
        return 2
    cfg = load_config(a.config)
    t0 = time.perf_counter()
    with Store(db) as st:
        if not st.count():
            print(f"[NG] 업무가 없는 DB 다: {db}")
            return 2
        stem = Path(st.meta("source_name", "wbs.xlsx")).stem
        out = Path(a.out) if a.out else REPO_ROOT / "output" / datetime.now().strftime("%Y%m%dT%H%M%S") / f"{stem}_export.xlsx"
        if out.exists():
            print(f"[NG] 이미 있는 파일은 덮어쓰지 않는다: {out}")
            return 2
        st.export_xlsx(out, cfg)
        n = st.count()
    print(f"[export] 업무 {n:,} → {out}  ({time.perf_counter() - t0:.1f}초)")
    return 0


def _module_uid(st, cfg, key: str) -> int:
    m = st.model(cfg)
    hits = [m.rows[t.no]["uid"] for t in m.wbs.tasks.values() if t.level == 1 and key.strip() in (t.wbs, t.name.strip())]
    if len(hits) != 1:
        names = " · ".join(f"{t.wbs} {t.name.strip()}" for t in m.wbs.tasks.values() if t.level == 1)
        raise ValueError(f"Lv1 '{key}' 를 {'찾지 못함' if not hits else '하나로 못 정함'} — 있는 것: {names}")
    return hits[0]


def _ask_password() -> str:
    import getpass
    p1 = getpass.getpass("비밀번호 (화면에 안 보임): ")
    p2 = getpass.getpass("한 번 더: ")
    if p1 != p2:
        raise ValueError("두 번 적은 비밀번호가 다르다")
    return p1


def cmd_role(a) -> int:
    from . import auth
    from .store import Store

    db = Path(a.db)
    if not db.exists():
        print(f"[NG] DB 가 없다: {db}  (먼저 import)")
        return 2
    try:
        with Store(db) as st:
            auth.ensure(st.db)
            if a.action == "list":
                rows = auth.list_roles(st.db, include_off=True)
                if not rows:
                    print("[role] 역할이 없다 — 예: role add admin1 --label \"관리자 1\" --kind admin")
                for r in rows:
                    mod = ""
                    if r["module_uid"]:
                        x = st.db.execute("SELECT name FROM task WHERE uid = ?", (r["module_uid"],)).fetchone()
                        mod = f" · 모듈 {x['name'].strip() if x else '(없음)'}"
                    print(f"  {r['login']:<16} {r['label']} ({auth.KINDS[r['kind']]}){mod}{'' if r['active'] else ' · 꺼짐'}")
                return 0
            if not a.login:
                print("[NG] 로그인 ID 를 주세요")
                return 2
            if a.action == "add":
                if not a.label or not a.kind:
                    print("[NG] --label 과 --kind 가 필요하다")
                    return 2
                mod = _module_uid(st, load_config(a.config), a.module) if a.kind == "leader" and a.module else None
                auth.check_new(a.login, a.label, a.kind, "x" * auth.MIN_PASSWORD)
                r = auth.add_role(st.db, a.login, a.label, a.kind, _ask_password(), mod)
                print(f"[role] 만듦: {r.login} — {r.label} ({r.kind_ko})")
            elif a.action == "passwd":
                auth.set_password(st.db, a.login, _ask_password())
                print(f"[role] 비밀번호 바꿈: {a.login} (그 역할의 로그인은 끊김)")
            else:
                auth.set_active(st.db, a.login, a.action == "on")
                print(f"[role] {a.login} {'켬' if a.action == 'on' else '끔 (로그인 끊김)'}")
    except ValueError as e:
        print(f"[NG] {e}")
        return 2
    return 0


def cmd_account(a) -> int:
    from . import auth, winauth
    from .store import Store

    db = Path(a.db)
    if not db.exists():
        print(f"[NG] DB 가 없다: {db}  (먼저 import)")
        return 2
    cfg = load_config(a.config)
    dom = winauth.default_domain((cfg.get("windows_login") or {}).get("domain") or None)
    try:
        with Store(db) as st:
            auth.ensure(st.db)
            if a.action == "list":
                rows = auth.list_accounts(st.db)
                wv = auth.windows_viewer(st.db)
                print(f"[account] 등록 {len(rows)} · 표에 없는 계정: "
                      + ("공용 조회자로 들임" if (cfg.get("windows_login") or {}).get("unregistered", "viewer") == "viewer" and wv else "거절"))
                for r in rows:
                    print(f"  {r['account']:<28} → {r['label']} ({auth.KINDS[r['kind']]}){'' if r['active'] else ' · 역할 꺼짐'}")
                return 0
            if not a.account:
                print("[NG] 계정을 주세요 (예: DOMAIN\\user)")
                return 2
            _, _, key = winauth.split_account(a.account, dom)
            if a.action == "set":
                if not a.role:
                    print("[NG] --role <역할 로그인ID> 가 필요하다 (role list 로 확인)")
                    return 2
                r = auth.set_account(st.db, key, a.role)
                print(f"[account] {key} → {r.label} ({r.kind_ko})")
            else:
                print(f"[account] {key} " + ("지움" if auth.remove_account(st.db, key) else "없음"))
    except ValueError as e:
        print(f"[NG] {e}")
        return 2
    return 0


def cmd_serve(a) -> int:
    try:
        import uvicorn
        from .server import create_app
    except ImportError as e:
        print(f"[NG] 서버 패키지가 없다 ({e.name}) — Tool 온톨로지 .venv 에 starlette·uvicorn 이 있어야 한다")
        return 2
    from . import auth
    from .store import Store

    db = Path(a.db)
    try:
        if db.exists():
            with Store(db) as st:
                auth.ensure(st.db)
                if not any(r["kind"] == "admin" and r["active"] for r in auth.list_roles(st.db)):
                    print("[serve] 관리자 역할이 아직 없다 (보기에는 지장 없음) — 만들려면: role add admin1 --label \"관리자 1\" --kind admin"
                          " 뒤 account set <내 Windows 계정> --role admin1")
        t0 = time.perf_counter()
        app = create_app(db, a.config, a.title or "WBS")
        _, _, info = app.state.wbs.data()          # 첫 화면이 기다리지 않게 미리 계산
    except ValueError as e:
        print(f"[NG] {e}")
        return 2
    print(f"[serve] {db.name} · 업무 {info['tasks']:,} · 선행 {info['edges']:,} · 판 #{info['rev']} ({time.perf_counter() - t0:.1f}초)")
    shown = "<이 PC 의 IP>" if a.host in ("0.0.0.0", "::") else a.host
    if bool(a.cert) != bool(a.key):
        print("[NG] HTTPS 는 --cert 와 --key 를 함께")
        return 2
    srv = app.state.wbs
    print("  · 로그인: " + (f"Windows 계정 (기본 도메인 {srv.win_domain}, 등록 안 된 계정 = "
                            + ("조회자" if srv.win_unregistered == "viewer" else "거절") + ") + 도구 역할 계정(비상용)"
                            if srv.win else "도구 역할 계정만"))
    if not a.cert:
        print("  · HTTP (암호화 없음) — 비밀번호가 망을 그대로 지난다. 사내 인증서가 있으면 --cert·--key 로 HTTPS")
    proto = "https" if a.cert else "http"
    print(f"  · 브라우저: {proto}://{shown}:{a.port}/   (끄려면 이 창에서 Ctrl+C)")
    uvicorn.run(app, host=a.host, port=a.port, log_level="warning", ssl_certfile=a.cert, ssl_keyfile=a.key)
    return 0


def cmd_synth(a) -> int:
    from .synth import generate
    out = Path(a.out)
    info = generate(out, a.rows, a.seed)
    print(f"[synth] {info['rows']:,}행 → {out}")
    print("  · 심은 문제: " + " · ".join(f"{k} {len(v)}" for k, v in info["injected"].items()))
    return 0


def main(argv=None) -> int:
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, "reconfigure"):
            s.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(prog="engenius_wbs", description="WBS 엑셀 → 온톨로지(TTL)·검사·보고서·보기 HTML")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="WBS 엑셀을 읽어 산출물을 만든다")
    b.add_argument("xlsx")
    b.add_argument("--sheet")
    b.add_argument("--out")
    b.add_argument("--config")
    b.add_argument("--title")
    s = sub.add_parser("synth", help="가짜 WBS 엑셀을 만든다")
    s.add_argument("--out", default=str(REPO_ROOT / "data" / "synthetic" / "wbs_synthetic.xlsx"))
    s.add_argument("--rows", type=int, default=12827)
    s.add_argument("--seed", type=int, default=7)
    p = sub.add_parser("apply", help="보기 HTML 의 변경 목록을 원본 엑셀에 반영해 새 판 엑셀을 쓴다")
    p.add_argument("files", nargs=2, help="원본 WBS .xlsx 와 변경 목록 .json (순서 상관없음)")
    p.add_argument("--out")
    p.add_argument("--config")
    p.add_argument("--force", action="store_true", help="원본 sha256 이 달라도 적용")
    p.add_argument("--no-build", action="store_true", help="새 엑셀로 build 하지 않음")
    default_db = str(REPO_ROOT / "data" / "real" / "wbs.db")
    i = sub.add_parser("import", help="WBS 엑셀을 도구 DB 로 들여온다 (v0.5 S1)")
    i.add_argument("xlsx")
    i.add_argument("--db", default=default_db)
    i.add_argument("--sheet")
    i.add_argument("--config")
    i.add_argument("--force", action="store_true", help="업무가 든 DB 를 .bak-<시각> 으로 옮기고 새로 들여온다 (묻지 않음)")
    i.add_argument("--ask", action="store_true", help="업무가 든 DB 가 있으면 덮어쓸지 묻는다 (import.bat)")
    e = sub.add_parser("export", help="도구 DB 를 원본 양식 엑셀로 내보낸다")
    e.add_argument("--db", default=default_db)
    e.add_argument("--out")
    e.add_argument("--config")
    r = sub.add_parser("role", help="역할 계정 만들기·보기·비밀번호 (v0.5 S2)")
    r.add_argument("action", choices=["add", "list", "passwd", "off", "on"])
    r.add_argument("login", nargs="?")
    r.add_argument("--label", help="역할 이름 (예: 관리자 1, FAN 리더) — 사람 이름 금지")
    r.add_argument("--kind", choices=["admin", "leader", "viewer"])
    r.add_argument("--module", help="모듈 리더의 Lv1 — 이름 또는 WBS 코드")
    r.add_argument("--db", default=default_db)
    r.add_argument("--config")
    ac = sub.add_parser("account", help="Windows 계정 → 역할 (v0.5 S2)")
    ac.add_argument("action", choices=["set", "list", "remove"])
    ac.add_argument("account", nargs="?")
    ac.add_argument("--role", help="역할 로그인 ID (role list)")
    ac.add_argument("--db", default=default_db)
    ac.add_argument("--config")
    s = sub.add_parser("serve", help="읽기 서버 (v0.5 S2)")
    s.add_argument("--db", default=default_db)
    s.add_argument("--host", default="127.0.0.1", help="다른 PC 에서 들어오게 하려면 0.0.0.0")
    s.add_argument("--port", type=int, default=8100)
    s.add_argument("--title")
    s.add_argument("--cert", help="HTTPS 인증서 파일 (PEM)")
    s.add_argument("--key", help="HTTPS 개인 키 파일 (PEM)")
    s.add_argument("--config")
    a = ap.parse_args(argv)
    return {"build": cmd_build, "synth": cmd_synth, "apply": cmd_apply, "import": cmd_import, "export": cmd_export,
            "role": cmd_role, "account": cmd_account, "serve": cmd_serve}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
