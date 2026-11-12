"""python -m engenius_wbs <명령>

    build <WBS.xlsx> [--sheet 이름] [--out 폴더] [--config wbs.yaml] [--title 제목]
        → <폴더>/report.xlsx · wbs_view.html · wbs.ttl · derived.ttl · run_log.json (기본 output/<시각>/)
    synth [--out data/synthetic/wbs_synthetic.xlsx] [--rows 12827] [--seed 7]
        → 가짜 WBS 엑셀 (실데이터 없이 시험할 때)
    apply <원본 WBS.xlsx> <변경 목록.json> [--out 폴더] [--config wbs.yaml] [--force] [--no-build]
        → <원본>_수정_v01.xlsx (원본 형식 그대로, 원본은 안 건드림) + .changes.json, 이어서 그 파일로 build
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
    a = ap.parse_args(argv)
    return {"build": cmd_build, "synth": cmd_synth, "apply": cmd_apply}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
