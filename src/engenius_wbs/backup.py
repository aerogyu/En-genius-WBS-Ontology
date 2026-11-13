"""백업·복구 (v0.5 S5, docs/05 14-2). 표준 라이브러리 SQLite 온라인 백업 — 서버가 켜져 있어도 일관된 사본.

  언제   start(서버 켤 때) · daily(하루 한 번, backup.daily_at 이후 · 마지막 백업 뒤 저장이 있었을 때만) · snapshot(판 확정 직전) · manual(backup 명령·화면)
  어디   설정 backup.dir (비면 DB 옆 backup\\) · 이름 wbs_<YYYYMMDD-HHMMSS>_r<판>_<종류>.db
  보관   종류마다 최근 backup.keep 개 (기본 14). 판 확정 직전 사본(snapshot)은 지우지 않는다.
  검사   사본마다 다시 열어 PRAGMA integrity_check · 스키마 판 · 업무 수.
  복구   restore — 서버가 꺼져 있을 때만 (서버가 켜 두는 표시 파일 <db>.serving 을 지울 수 없으면 켜져 있는 것).
         사본 검사 → 지금 DB 를 .bak-<시각> 으로 → 사본을 제자리에.

사본에는 Windows 계정 → 역할 표와 로그인 세션 해시가 들어 있다 (복구하면 그대로 돌아오게) — 백업 폴더는 서버 PC 와 같은 수준으로 다룰 것.
"""
from __future__ import annotations

import os
import re
import shutil
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from .config import REPO_ROOT
from .store import DATA_KINDS

KINDS = ("start", "daily", "snapshot", "manual")
KIND_KO = {"start": "서버 켤 때", "daily": "하루 한 번", "snapshot": "판 확정 직전", "manual": "수동"}
NAME_RE = re.compile(r"^wbs_(\d{8}-\d{6})_r(\d+)_(start|daily|snapshot|manual)\.db$")
DEFAULT_KEEP = 14
DEFAULT_DAILY_AT = "18:00"


@dataclass
class BackupFile:
    path: Path
    at: datetime
    rev: int
    kind: str

    def as_dict(self) -> dict:
        return {"name": self.path.name, "at": self.at.isoformat(timespec="seconds"), "rev": self.rev, "kind": self.kind,
                "kind_ko": KIND_KO[self.kind], "bytes": self.path.stat().st_size if self.path.exists() else 0}


def backup_dir(db_path: str | Path, cfg: dict) -> Path:
    d = (cfg.get("backup") or {}).get("dir")
    if d:
        p = Path(d)
        return p if p.is_absolute() else REPO_ROOT / p
    return Path(db_path).parent / "backup"


def _rev(db: sqlite3.Connection) -> int:
    q = ",".join("?" * len(DATA_KINDS))
    return db.execute(f"SELECT COALESCE(MAX(id), 0) FROM txn WHERE kind IN ({q})", DATA_KINDS).fetchone()[0]


def _ro(path: Path) -> sqlite3.Connection:
    """읽기 전용 연결 (사본·DB 를 건드리지 않게) — 경로의 한글·띄어쓰기는 URI 로 바꿔 넘긴다."""
    return sqlite3.connect(f"file:{quote(Path(path).resolve().as_posix(), safe='/:')}?mode=ro", uri=True)


def verify(path: str | Path) -> dict:
    """사본(또는 DB) 검사 → {ok, integrity, schema, tasks, rev}."""
    path = Path(path)
    out = {"ok": False, "integrity": "", "schema": None, "tasks": 0, "rev": 0}
    if not path.exists():
        out["integrity"] = "파일 없음"
        return out
    try:
        db = _ro(path)
        try:
            out["integrity"] = db.execute("PRAGMA integrity_check").fetchone()[0]
            v = db.execute("SELECT value FROM meta WHERE key = 'schema_version'").fetchone()
            out["schema"] = int(v[0]) if v else None
            out["tasks"] = db.execute("SELECT COUNT(*) FROM task WHERE deleted = 0").fetchone()[0]
            out["rev"] = _rev(db)
        finally:
            db.close()
    except sqlite3.Error as e:
        out["integrity"] = f"{type(e).__name__}: {e}"
        return out
    out["ok"] = out["integrity"] == "ok" and out["schema"] is not None and out["tasks"] > 0
    return out


def list_backups(folder: str | Path) -> list[BackupFile]:
    folder = Path(folder)
    if not folder.is_dir():
        return []
    out = []
    for p in folder.iterdir():
        mt = NAME_RE.match(p.name)
        if mt:
            out.append(BackupFile(p, datetime.strptime(mt.group(1), "%Y%m%d-%H%M%S"), int(mt.group(2)), mt.group(3)))
    return sorted(out, key=lambda b: (b.at, b.path.name))


def prune(folder: str | Path, keep: int = DEFAULT_KEEP) -> list[Path]:
    """종류마다 최근 keep 개만 (snapshot 은 지우지 않음) → 지운 파일."""
    gone = []
    for kind in KINDS:
        if kind == "snapshot":
            continue
        files = [b for b in list_backups(folder) if b.kind == kind]
        for b in files[:max(0, len(files) - max(1, keep))]:
            b.path.unlink()
            gone.append(b.path)
    return gone


def make_backup(db_path: str | Path, cfg: dict, kind: str) -> tuple[Path, dict]:
    """온라인 백업 → (사본 경로, 검사 결과). 검사에 실패하면 사본을 남기지 않고 RuntimeError."""
    if kind not in KINDS:
        raise ValueError(f"모르는 백업 종류: {kind}")
    db_path = Path(db_path)
    if not db_path.exists():
        raise ValueError(f"DB 가 없다: {db_path}")
    folder = backup_dir(db_path, cfg)
    folder.mkdir(parents=True, exist_ok=True)
    src = sqlite3.connect(str(db_path))
    try:
        rev = _rev(src)
        now = datetime.now()
        name = f"wbs_{now.strftime('%Y%m%d-%H%M%S')}_r{rev}_{kind}.db"
        while (folder / name).exists():                 # 같은 초에 두 번
            now = datetime.fromtimestamp(now.timestamp() + 1)
            name = f"wbs_{now.strftime('%Y%m%d-%H%M%S')}_r{rev}_{kind}.db"
        tmp = folder / (name + ".part")
        if tmp.exists():
            tmp.unlink()
        dst = sqlite3.connect(str(tmp))
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    chk = verify(tmp)
    if not chk["ok"]:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f"백업 사본 검사 실패: {chk['integrity']} (업무 {chk['tasks']})")
    out = folder / name
    tmp.rename(out)
    prune(folder, int((cfg.get("backup") or {}).get("keep") or DEFAULT_KEEP))
    return out, chk


def daily_due(db_path: str | Path, cfg: dict, now: datetime | None = None) -> bool:
    """오늘 daily_at 이 지났고, 오늘 daily 사본이 없고, 마지막 사본 뒤 저장이 있었으면 True."""
    now = now or datetime.now()
    at = str((cfg.get("backup") or {}).get("daily_at") or DEFAULT_DAILY_AT)
    hh, mm = (int(x) for x in at.split(":"))
    if (now.hour, now.minute) < (hh, mm):
        return False
    files = list_backups(backup_dir(db_path, cfg))
    if any(b.kind == "daily" and b.at.date() == now.date() for b in files):
        return False
    src = _ro(Path(db_path))
    try:
        rev = _rev(src)
    finally:
        src.close()
    return not files or files[-1].rev != rev


# --- 서버 표시 파일 · 복구 ------------------------------------------------------------------------
def marker_path(db_path: str | Path) -> Path:
    db_path = Path(db_path)
    return db_path.with_name(db_path.name + ".serving")


class ServingMarker:
    """서버가 켜져 있는 동안 열어 두는 파일 — Windows 에서는 열린 파일을 다른 곳에서 지울 수 없다 (복구가 이것으로 안다)."""

    def __init__(self, db_path: str | Path):
        self.path = marker_path(db_path)
        self.f = None

    def __enter__(self):
        self.f = open(self.path, "w", encoding="utf-8")
        self.f.write(f"pid {os.getpid()} · {datetime.now().isoformat(timespec='seconds')}\n")
        self.f.flush()
        return self

    def __exit__(self, *exc):
        if self.f:
            self.f.close()
            self.f = None
        try:
            self.path.unlink()
        except OSError:
            pass


def server_running(db_path: str | Path) -> bool:
    m = marker_path(db_path)
    if not m.exists():
        return False
    try:
        m.unlink()                      # 남은 표시 파일(서버가 비정상 종료)이면 지워진다
    except PermissionError:
        return True
    return False


def restore(db_path: str | Path, backup_file: str | Path) -> tuple[Path | None, dict]:
    """→ (옛 DB 를 옮긴 .bak 경로 또는 None, 사본 검사). 서버가 켜져 있거나 사본이 망가졌으면 ValueError."""
    db_path, backup_file = Path(db_path), Path(backup_file)
    if server_running(db_path):
        raise ValueError("서버가 켜져 있다 — 서버 창을 닫은 뒤 복구할 것")
    chk = verify(backup_file)
    if not chk["ok"]:
        raise ValueError(f"사본 검사 실패: {chk['integrity']} (업무 {chk['tasks']})")
    bak = None
    if db_path.exists():
        bak = db_path.with_name(f"{db_path.name}.bak-{datetime.now().strftime('%Y%m%dT%H%M%S')}")
        db_path.rename(bak)
    tmp = db_path.with_name(db_path.name + ".restore")
    shutil.copy2(backup_file, tmp)
    tmp.rename(db_path)
    return bak, chk
