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


def cmd_import(a) -> int:
    from .store import Store

    src, db = Path(a.xlsx), Path(a.db)
    if not src.exists():
        print(f"[NG] 파일이 없다: {src}")
        return 2
    cfg = load_config(a.config)
    if db.exists():
        with Store(db) as st:
            n = st.count()
        if n and not a.force:
            print(f"[NG] 이미 업무 {n:,}개가 들어 있는 DB 다: {db}")
            print("     다른 --db 를 주거나, 새로 들여오려면 --force (옛 DB 는 .bak-<시각> 으로 이름을 바꿔 남긴다)")
            return 2
        bak = db.with_name(f"{db.name}.bak-{datetime.now().strftime('%Y%m%dT%H%M%S')}")
        db.rename(bak)
        if n:
            print(f"[import] 옛 DB → {bak.name}")
        else:
            bak.unlink()
    t0 = time.perf_counter()
    try:
        with Store(db) as st:
            rep = st.import_xlsx(src, cfg, a.sheet)
            notes = [dict(r) for r in st.db.execute(
                "SELECT n.task_uid AS uid, t.src_no AS no, t.src_row AS excel_row, n.col, n.old, n.new FROM note n JOIN task t ON t.uid = n.task_uid ORDER BY n.id")]
    except ValueError as e:
        print(f"[NG] {e}")
        if db.exists() and not db.stat().st_size:
            db.unlink()
        return 2
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
    i.add_argument("--force", action="store_true", help="업무가 든 DB 를 .bak-<시각> 으로 옮기고 새로 들여온다")
    e = sub.add_parser("export", help="도구 DB 를 원본 양식 엑셀로 내보낸다")
    e.add_argument("--db", default=default_db)
    e.add_argument("--out")
    e.add_argument("--config")
    a = ap.parse_args(argv)
    return {"build": cmd_build, "synth": cmd_synth, "apply": cmd_apply, "import": cmd_import, "export": cmd_export}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
