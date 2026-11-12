"""TTL 내보내기 — wbs.ttl(엑셀에 적힌 것) + derived.ttl(계산한 것, provenance 붙임).

IRI: egid:WBS-<번호 6자리> (번호 = A열 — 같은 엑셀 판 안에서만 안정. 판을 넘는 이어받기는 v0.2 에서 레지스트리로),
     egid:WDEP-<선행>-<후속> · egid:WDEL-<해시> (산출물 이름) · egid:WRUN-<시각> (실행).
접두는 형제 저장소(Tool: TOOL·USE·IO·CAT·FLD·LOAD·FEED·RECON·ISO·HANDOFF / 워크플로: PLAN·DSC·…·STEP·IFC·LOOP…)와 겹치지 않는다.
"""
from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path

from rdflib import RDF, RDFS, XSD, Graph, Literal, Namespace, URIRef
from rdflib.namespace import PROV, SKOS

from .analyze import Analysis
from .parse import Wbs

EGID = Namespace("http://en-genius.local/id/")
WBS = Namespace("http://en-genius.local/onto/wbs/")
RULE_VERSION = "wbs_derive_v1"


def task_iri(no: int) -> URIRef:
    return EGID[f"WBS-{no:06d}"]


def deliverable_iri(name: str) -> URIRef:
    return EGID["WDEL-" + hashlib.sha1(name.encode("utf-8")).hexdigest()[:12]]


def schema_graph(cfg: dict) -> Graph:
    g = Graph()
    g.bind("wbs", WBS)
    g.add((WBS.WBSElement, RDF.type, RDFS.Class))
    g.add((WBS.WBSElement, RDFS.label, Literal("WBS 요소", lang="ko")))
    for lv, info in sorted(cfg["levels"].items()):
        c = WBS[info["class"]]
        g.add((c, RDF.type, RDFS.Class))
        g.add((c, RDFS.subClassOf, WBS.WBSElement))
        g.add((c, RDFS.label, Literal(f"Lv{lv} {info['label']}", lang="ko")))
    for c, label in ((WBS.Dependency, "선행 관계"), (WBS.Deliverable, "산출물"), (WBS.Run, "적재 실행")):
        g.add((c, RDF.type, RDFS.Class))
        g.add((c, RDFS.label, Literal(label, lang="ko")))
    for p, label in ((WBS.isMilestone, "마일스톤 (기간 0)"), (WBS.dateFloatDays, "총 여유 (평일, 엑셀 날짜 기준)"),
                     (WBS.onCriticalPath, "주경로"), (WBS.nearCriticalPath, "준주경로"), (WBS.gapDays, "선행 간격 (평일, 엑셀 날짜 기준)"),
                     (WBS.partOf, "상위 WBS"), (WBS.precededBy, "선행 작업"), (WBS.produces, "산출물을 만든다"),
                     (WBS.predecessor, "선행 업무"), (WBS.successor, "후속 업무")):
        g.add((p, RDF.type, RDF.Property))
        g.add((p, RDFS.label, Literal(label, lang="ko")))
    return g


def build_graphs(w: Wbs, a: Analysis, cfg: dict, run_id: str) -> tuple[Graph, Graph]:
    g = schema_graph(cfg)
    run = EGID[run_id]
    g.add((run, RDF.type, WBS.Run))
    g.add((run, WBS.sourceFile, Literal(Path(w.source).name)))
    g.add((run, WBS.sheet, Literal(w.sheet)))
    levels = cfg["levels"]
    for t in w.tasks.values():
        s = task_iri(t.no)
        g.add((s, RDF.type, WBS[levels[t.level]["class"]] if t.level in levels else WBS.WBSElement))
        g.add((s, SKOS.prefLabel, Literal(t.name, lang="ko")))
        g.add((s, WBS.rowNumber, Literal(t.no)))
        g.add((s, WBS.wbsCode, Literal(t.wbs)))
        g.add((s, WBS.level, Literal(t.level)))
        if t.is_milestone:
            g.add((s, WBS.isMilestone, Literal(True)))
        g.add((s, PROV.wasGeneratedBy, run))
        if t.parent is not None:
            g.add((s, WBS.partOf, task_iri(t.parent)))
        if t.duration_days is not None:
            g.add((s, WBS.durationDays, Literal(t.duration_days, datatype=XSD.decimal)))
        if t.estimated:
            g.add((s, WBS.durationEstimated, Literal(True)))
        if t.start:
            g.add((s, WBS.start, Literal(t.start, datatype=XSD.date)))
        if t.finish:
            g.add((s, WBS.finish, Literal(t.finish, datatype=XSD.date)))
        if t.staff is not None:
            g.add((s, WBS.staff, Literal(t.staff, datatype=XSD.decimal)))
        if t.man_hours is not None:
            g.add((s, WBS.manHours, Literal(t.man_hours, datatype=XSD.decimal)))
        for d in t.deliverables:
            o = deliverable_iri(d)
            g.add((o, RDF.type, WBS.Deliverable))
            g.add((o, RDFS.label, Literal(d, lang="ko")))
            g.add((s, WBS.produces, o))
    for e in a.edges:
        s, p = task_iri(e.succ), task_iri(e.pred)
        g.add((s, WBS.precededBy, p))
        dep = EGID[f"WDEP-{e.pred:06d}-{e.succ:06d}"]
        g.add((dep, RDF.type, WBS.Dependency))
        g.add((dep, WBS.predecessor, p))
        g.add((dep, WBS.successor, s))
        g.add((dep, WBS.depType, Literal(e.type)))
        if e.lag_text:
            g.add((dep, WBS.lagText, Literal(e.lag_text)))
        if e.lag_days:
            g.add((dep, WBS.lagDays, Literal(e.lag_days, datatype=XSD.decimal)))

    d = Graph()
    d.bind("wbs", WBS)
    d.add((run, PROV.generatedAtTime, Literal(datetime.now().replace(microsecond=0), datatype=XSD.dateTime)))
    d.add((run, WBS.ruleVersion, Literal(RULE_VERSION)))
    d.add((run, WBS.projectWorkdays, Literal(a.project_days, datatype=XSD.decimal)))
    critical = set(a.critical_chain)
    sch = a.schedule
    on_cp, near = set(sch.critical), set(sch.near)
    for e, gap in zip(a.edges, sch.gaps):
        if gap is not None:
            d.add((EGID[f"WDEP-{e.pred:06d}-{e.succ:06d}"], WBS.gapDays, Literal(gap, datatype=XSD.decimal)))
    for t in w.tasks.values():
        s = task_iri(t.no)
        d.add((s, WBS.upstreamCount, Literal(a.up_count.get(t.no, 0))))
        d.add((s, WBS.downstreamCount, Literal(a.down_count.get(t.no, 0))))
        d.add((s, WBS.earlyStartDay, Literal(a.es.get(t.no, 0.0), datatype=XSD.decimal)))
        d.add((s, WBS.totalFloatDays, Literal(a.total_float.get(t.no, 0.0), datatype=XSD.decimal)))
        if t.no in critical:
            d.add((s, WBS.onLongestChain, Literal(True)))
        if t.no in sch.float_:
            d.add((s, WBS.dateFloatDays, Literal(sch.float_[t.no], datatype=XSD.decimal)))
        if t.no in on_cp:
            d.add((s, WBS.onCriticalPath, Literal(True)))
        elif t.no in near:
            d.add((s, WBS.nearCriticalPath, Literal(True)))
        if t.discipline:
            d.add((s, WBS.disciplineName, Literal(t.discipline, lang="ko")))
    for i, ((pd, sd), n) in enumerate(sorted(a.discipline_flows.items())):
        f = EGID[f"WFLOW-{hashlib.sha1(f'{pd}→{sd}'.encode()).hexdigest()[:12]}"]
        d.add((f, RDF.type, WBS.DisciplineFlow))
        d.add((f, WBS.fromDiscipline, Literal(pd, lang="ko")))
        d.add((f, WBS.toDiscipline, Literal(sd, lang="ko")))
        d.add((f, WBS.edgeCount, Literal(n)))
        d.add((f, PROV.wasGeneratedBy, run))
    return g, d


def write_ttl(w: Wbs, a: Analysis, cfg: dict, out: Path, run_id: str) -> tuple[Path, Path]:
    g, d = build_graphs(w, a, cfg, run_id)
    p1, p2 = out / "wbs.ttl", out / "derived.ttl"
    g.serialize(p1, format="turtle", encoding="utf-8")
    d.serialize(p2, format="turtle", encoding="utf-8")
    return p1, p2
