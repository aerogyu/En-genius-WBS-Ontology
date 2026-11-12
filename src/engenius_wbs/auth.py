"""역할 계정 (v0.5 S2) — 로그인·세션. 표준 라이브러리만 (hashlib PBKDF2·secrets·sqlite3).

docs/05 4장: 이력·화면·내보낸 파일에는 **역할 이름만** 남는다 ('관리자 1', 'FAN 리더'). 사람 이름은 어디에도 적지 않는다.
사내 계정 연동(S5)은 계정 ID → 역할 표를 서버 DB 에만 두고 이 모듈의 verify 자리에 어댑터를 붙인다.

  role     로그인 ID(영문) · 역할 이름 · 종류(admin 관리자 / leader 모듈 리더 / viewer 조회자) · 모듈(Lv1 업무 uid, 리더만) · 비밀번호 해시
  session  로그인 표 — 토큰은 SHA-256 해시만 DB 에 (DB 파일을 봐도 쿠키를 만들 수 없게), 만료 시각
"""
from __future__ import annotations

import hashlib
import hmac
import re
import secrets
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime

KINDS = {"admin": "관리자", "leader": "모듈 리더", "viewer": "조회자"}
LOGIN_RE = re.compile(r"^[A-Za-z0-9._-]{2,32}$")
PBKDF2_ITERS = 240_000
SESSION_HOURS = 12
MIN_PASSWORD = 8

AUTH_SCHEMA = """
CREATE TABLE IF NOT EXISTS role (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  login TEXT NOT NULL UNIQUE COLLATE NOCASE,     -- 로그인 ID (영문·숫자)
  label TEXT NOT NULL UNIQUE,                    -- 역할 이름 — 이력에 남는 것 (사람 이름 금지)
  kind TEXT NOT NULL CHECK (kind IN ('admin', 'leader', 'viewer')),
  module_uid INTEGER REFERENCES task(uid),       -- 모듈 리더의 Lv1 업무
  pw TEXT NOT NULL,                              -- pbkdf2_sha256$반복$소금$해시
  active INTEGER NOT NULL DEFAULT 1,
  created TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS session (
  token_hash TEXT PRIMARY KEY,
  role_id INTEGER NOT NULL REFERENCES role(id),
  created REAL NOT NULL, expires REAL NOT NULL
);
"""


@dataclass(frozen=True)
class Role:
    id: int
    login: str
    label: str
    kind: str
    module_uid: int | None

    @property
    def kind_ko(self) -> str:
        return KINDS[self.kind]

    @property
    def is_admin(self) -> bool:
        return self.kind == "admin"


def ensure(db: sqlite3.Connection) -> None:
    db.executescript(AUTH_SCHEMA)


def hash_password(pw: str, iters: int = PBKDF2_ITERS) -> str:
    salt = secrets.token_bytes(16)
    h = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iters)
    return f"pbkdf2_sha256${iters}${salt.hex()}${h.hex()}"


def check_password(pw: str, stored: str) -> bool:
    try:
        algo, iters, salt, want = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        got = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), int(iters))
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(got.hex(), want)


_DUMMY = hash_password("-", 1000)


def _role(r: sqlite3.Row) -> Role:
    return Role(r["id"], r["login"], r["label"], r["kind"], r["module_uid"])


def check_new(login: str, label: str, kind: str, password: str) -> None:
    if not LOGIN_RE.match(login or ""):
        raise ValueError("로그인 ID 는 영문·숫자·. _ - 로 2~32자")
    if not (label or "").strip() or len(label) > 40:
        raise ValueError("역할 이름이 비었거나 40자를 넘는다")
    if kind not in KINDS:
        raise ValueError(f"종류는 {' / '.join(KINDS)} 중 하나")
    if len(password or "") < MIN_PASSWORD:
        raise ValueError(f"비밀번호는 {MIN_PASSWORD}자 이상")


def add_role(db: sqlite3.Connection, login: str, label: str, kind: str, password: str, module_uid: int | None = None) -> Role:
    ensure(db)
    check_new(login, label, kind, password)
    if kind == "leader" and module_uid is None:
        raise ValueError("모듈 리더는 모듈(Lv1)을 정해야 한다")
    if kind != "leader":
        module_uid = None
    if db.execute("SELECT 1 FROM role WHERE login = ? OR label = ?", (login, label.strip())).fetchone():
        raise ValueError("같은 로그인 ID 나 역할 이름이 이미 있다")
    cur = db.execute("INSERT INTO role(login, label, kind, module_uid, pw, created) VALUES(?,?,?,?,?,?)",
                     (login, label.strip(), kind, module_uid, hash_password(password), datetime.now().isoformat(timespec="seconds")))
    db.commit()
    return get_role(db, cur.lastrowid)


def get_role(db: sqlite3.Connection, role_id: int) -> Role | None:
    r = db.execute("SELECT * FROM role WHERE id = ? AND active = 1", (role_id,)).fetchone()
    return _role(r) if r else None


def list_roles(db: sqlite3.Connection, include_off: bool = False) -> list[sqlite3.Row]:
    ensure(db)
    q = "SELECT * FROM role" + ("" if include_off else " WHERE active = 1") + " ORDER BY kind, label"
    return db.execute(q).fetchall()


def set_password(db: sqlite3.Connection, login: str, password: str) -> None:
    if len(password or "") < MIN_PASSWORD:
        raise ValueError(f"비밀번호는 {MIN_PASSWORD}자 이상")
    r = db.execute("SELECT id FROM role WHERE login = ?", (login,)).fetchone()
    if not r:
        raise ValueError(f"없는 로그인 ID: {login}")
    db.execute("UPDATE role SET pw = ? WHERE id = ?", (hash_password(password), r["id"]))
    db.execute("DELETE FROM session WHERE role_id = ?", (r["id"],))   # 바꾸면 기존 로그인은 끊는다
    db.commit()


def set_active(db: sqlite3.Connection, login: str, active: bool) -> None:
    r = db.execute("SELECT id FROM role WHERE login = ?", (login,)).fetchone()
    if not r:
        raise ValueError(f"없는 로그인 ID: {login}")
    db.execute("UPDATE role SET active = ? WHERE id = ?", (1 if active else 0, r["id"]))
    if not active:
        db.execute("DELETE FROM session WHERE role_id = ?", (r["id"],))
    db.commit()


def verify(db: sqlite3.Connection, login: str, password: str) -> Role | None:
    ensure(db)
    r = db.execute("SELECT * FROM role WHERE login = ? AND active = 1", (login or "",)).fetchone()
    if r is None:
        check_password(password or "", _DUMMY)          # 없는 ID 도 비슷한 시간이 걸리게
        return None
    return _role(r) if check_password(password or "", r["pw"]) else None


def _th(token: str) -> str:
    return hashlib.sha256(token.encode("ascii", "replace")).hexdigest()


def new_session(db: sqlite3.Connection, role: Role, hours: float = SESSION_HOURS) -> str:
    token = secrets.token_urlsafe(32)
    now = time.time()
    db.execute("DELETE FROM session WHERE expires < ?", (now,))
    db.execute("INSERT INTO session(token_hash, role_id, created, expires) VALUES(?,?,?,?)", (_th(token), role.id, now, now + hours * 3600))
    db.commit()
    return token


def session_role(db: sqlite3.Connection, token: str | None) -> Role | None:
    if not token or len(token) > 100:
        return None
    r = db.execute("SELECT role_id FROM session WHERE token_hash = ? AND expires > ?", (_th(token), time.time())).fetchone()
    return get_role(db, r["role_id"]) if r else None


def drop_session(db: sqlite3.Connection, token: str | None) -> None:
    if token:
        db.execute("DELETE FROM session WHERE token_hash = ?", (_th(token),))
        db.commit()
