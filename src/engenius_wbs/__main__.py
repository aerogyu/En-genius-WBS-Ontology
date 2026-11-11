"""python -m engenius_wbs <명령>

    build <WBS.xlsx> [--sheet 이름] [--out 폴더] [--config wbs.yaml] [--title 제목]
        → <폴더>/report.xlsx · wbs_view.html · wbs.ttl · derived.ttl · run_log.json (기본 output/<시각>/)
    synth [--out data/synthetic/wbs_synthetic.xlsx] [--rows 12827] [--seed 7]
        → 가짜 WBS 엑셀 (실데이터 없이 시험할 때)
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
    an = analyze(w)
    t2 = time.perf_counter()
    rep = write_report(w, an, cfg, out / "report.xlsx")
    view = write_html(w, an, out / "wbs_view.html", a.title or "WBS 선후 관계")
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
           "seconds": {"read": round(t1 - t0, 2), "analyze": round(t2 - t1, 2), "write": round(t3 - t2, 2)}}
    (out / "run_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"[build] {src.name} ({w.sheet}) → {out}")
    print(f"  · 업무 {len(w.tasks):,} · 선행 관계 {len(an.edges):,} · 순환 {len(an.cycles)} · 가장 긴 사슬 {an.project_days:.0f} 근무일(근사)")
    print("  · 문제: " + (" · ".join(f"{k} {v}" for k, v in sorted(cnt.items())) or "없음"))
    print(f"  · 보고서 {rep.name} · 보기 {view.name} · {ttl[0].name} · {ttl[1].name}  ({t3 - t0:.1f}초)")
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
    a = ap.parse_args(argv)
    return cmd_build(a) if a.cmd == "build" else cmd_synth(a)


if __name__ == "__main__":
    sys.exit(main())
