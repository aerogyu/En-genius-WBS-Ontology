"""읽기 서버 (v0.5 S2) — 역할 계정 로그인 · 화면 = 지금 보기 + 서버 데이터 · 조회자 권한."""
from __future__ import annotations

import json
from io import BytesIO

import pytest

pytest.importorskip("starlette")
pytest.importorskip("httpx")
from openpyxl import load_workbook
from starlette.testclient import TestClient

from engenius_wbs import auth, winauth
from engenius_wbs.__main__ import main
from engenius_wbs.analyze import analyze
from engenius_wbs.server import CODE_TAG, COOKIE, FAIL_LIMIT, create_app
from engenius_wbs.store import Store
from engenius_wbs.viz import payload
from test_wbs import CFG, make_small

PW = "test-pass-1234"


@pytest.fixture()
def app_db(tmp_path):
    db = tmp_path / "wbs.db"
    with Store(db) as st:
        st.import_xlsx(make_small(tmp_path / "small.xlsx"), CFG)
        lv1 = st.db.execute("SELECT uid FROM task WHERE parent_uid IS NULL ORDER BY pos").fetchone()[0]
        auth.add_role(st.db, "admin1", "관리자 1", "admin", PW)
        auth.add_role(st.db, "fan", "FAN 리더", "leader", PW, lv1)
        auth.add_role(st.db, "view", "조회", "viewer", PW)
    return create_app(db), db


def login(c: TestClient, who: str, pw: str = PW, mode: str = "role"):
    return c.post("/login", data={"login": who, "password": pw, "mode": mode}, follow_redirects=False)


def test_login_required(app_db):
    app, _ = app_db
    c = TestClient(app)
    r = c.get("/", follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "login"
    for path in ("/data.js", "/api/status", "/api/roles", "/export.xlsx"):
        assert c.get(path).status_code == 401, path
    r = c.get("/login")
    assert r.status_code == 200 and 'name="password"' in r.text and r.headers["x-frame-options"] == "DENY"
    assert login(c, "view", "wrong-password").status_code == 401
    assert login(c, "nobody", PW).status_code == 401
    assert COOKIE not in c.cookies


def test_view_same_data_as_build(app_db):
    app, db = app_db
    c = TestClient(app)
    r = login(c, "view")
    assert r.status_code == 303 and COOKIE in r.cookies
    ck = r.headers["set-cookie"].lower()
    assert "httponly" in ck and "samesite=strict" in ck
    page = c.get("/").text
    assert "const D=window.__WBS_DATA__;" in page and "__DATA__" not in page.replace("__WBS_DATA__", "")
    assert f'<script src="data.js?r=1&amp;v={CODE_TAG}"></script>' in page and "조회 (조회자) · 판 #1" in page
    assert "저장된 판 엑셀 내려받기" in page and "ED.size" in page
    js = c.get("/data.js")
    assert js.headers["content-type"].startswith("text/javascript") and js.headers["cache-control"] == "private, no-cache"
    d = json.loads(js.text[len("window.__WBS_DATA__="):-1])
    with Store(db) as st:
        w = st.load(CFG)
    want = json.loads(json.dumps(payload(w, analyze(w, CFG), CFG)))
    want["sha256"] = ""
    d.pop("server")
    assert len(d.pop("uid")) == len(d.pop("fl")) == len(d.pop("wk")) == len(d["tasks"])
    d["edges"] = [e[:6] for e in d["edges"]]                  # 7번째 = 선행 느슨함 (서버 미리 보기용)
    assert d == want                                       # 화면 데이터 = 기존 build 의 보기 데이터 (DB 에서 만든 것)
    assert c.get("/data.js", headers={"if-none-match": js.headers["etag"]}).status_code == 304
    # 두 번째 브라우저도 같은 판
    c2 = TestClient(app)
    login(c2, "fan")
    assert c2.get("/data.js").text == js.text and "FAN 리더 · " in c2.get("/").text
    assert c2.get("/api/status").json() == {"rev": 1, "tasks": 12, "role": "FAN 리더", "kind": "leader",
                                            "edit_open": True, "edit_note": "", "inbox": 0}     # S5: 편집 열림 · 받은 요청 수


def test_viewer_rights_and_export(app_db):
    app, _ = app_db
    v, a = TestClient(app), TestClient(app)
    login(v, "view")
    login(a, "admin1")
    assert v.get("/api/roles").status_code == 403
    roles = a.get("/api/roles").json()
    assert {r["login"]: r["kind"] for r in roles if r["login"] != auth.WIN_VIEWER_LOGIN} == {"admin1": "admin", "fan": "leader", "view": "viewer"}
    assert next(r for r in roles if r["login"] == "fan")["module"]
    assert v.post("/api/status").status_code == 405       # 읽기 서버 — 쓰는 길이 없다
    x = v.get("/export.xlsx")
    assert x.status_code == 200 and "filename*=UTF-8''" in x.headers["content-disposition"]
    ws = load_workbook(BytesIO(x.content)).worksheets[0]
    assert ws.cell(4, 9).value is None and ws.cell(3, ws.max_column).value == "UID"


def test_session_store_and_logout(app_db):
    app, db = app_db
    c = TestClient(app)
    token = login(c, "admin1").cookies[COOKIE]
    with Store(db) as st:
        raw = st.db.execute("SELECT token_hash FROM session").fetchall()
        pw = st.db.execute("SELECT pw FROM role WHERE login = 'admin1'").fetchone()[0]
    assert token not in {r[0] for r in raw} and PW not in pw and pw.startswith("pbkdf2_sha256$")
    assert c.post("/logout", follow_redirects=False).status_code == 303
    c.cookies.set(COOKIE, token)
    assert c.get("/api/status").status_code == 401         # 끊긴 토큰은 다시 못 쓴다
    # 비밀번호를 바꾸거나 역할을 끄면 기존 로그인이 끊긴다
    c2 = TestClient(app)
    login(c2, "fan")
    with Store(db) as st:
        auth.set_active(st.db, "fan", False)
    assert c2.get("/api/status").status_code == 401 and login(TestClient(app), "fan").status_code == 401


def test_login_throttle(app_db):
    app, _ = app_db
    c = TestClient(app)
    for _ in range(FAIL_LIMIT):
        assert login(c, "view", "nope-nope").status_code == 401
    assert login(c, "view").status_code == 429             # 맞는 비밀번호도 잠시 막힌다


def test_rev_changes_data(app_db):
    app, db = app_db
    c = TestClient(app)
    login(c, "view")
    first = c.get("/data.js").text
    with Store(db) as st:
        uid = st.db.execute("SELECT uid FROM task WHERE duration_days > 0 AND uid NOT IN (SELECT parent_uid FROM task WHERE parent_uid IS NOT NULL) LIMIT 1").fetchone()[0]
        st._txn("edit")          # 일정을 바꾸는 저장 (DATA_KINDS)
        st.db.execute("UPDATE task SET duration_days = duration_days + 2, work_days = work_days + 2 WHERE uid = ?", (uid,))
        st.recompute(CFG)
        st.db.commit()
    assert c.get("/api/status").json()["rev"] == 2
    second = c.get("/data.js").text
    assert second != first and '"rev":2' in second


def test_role_cli(tmp_path, monkeypatch, capsys):
    db = tmp_path / "wbs.db"
    with Store(db) as st:
        st.import_xlsx(make_small(tmp_path / "small.xlsx"), CFG)
        code = st.model(CFG).wbs.tasks[1].wbs
    monkeypatch.setattr("getpass.getpass", lambda prompt="": PW)
    assert main(["role", "add", "admin1", "--label", "관리자 1", "--kind", "admin", "--db", str(db)]) == 0
    assert main(["role", "add", "fan", "--label", "FAN 리더", "--kind", "leader", "--module", code, "--db", str(db)]) == 0
    assert main(["role", "add", "x", "--label", "Y", "--kind", "leader", "--module", "없는 모듈", "--db", str(db)]) == 2
    assert main(["role", "add", "admin1", "--label", "관리자 2", "--kind", "admin", "--db", str(db)]) == 2
    monkeypatch.setattr("getpass.getpass", lambda prompt="": "short")
    assert main(["role", "passwd", "fan", "--db", str(db)]) == 2
    assert main(["role", "list", "--db", str(db)]) == 0
    out = capsys.readouterr().out
    assert "관리자 1 (관리자)" in out and "FAN 리더 (모듈 리더) · 모듈" in out and PW not in out
    with Store(db) as st:
        assert auth.verify(st.db, "FAN", PW).label == "FAN 리더"      # 로그인 ID 는 대소문자 무시


# --- Windows 계정 로그인 (확인 자체는 가짜로 바꿔 끼운다 — 개발 PC 에서 사내 계정을 쓸 수 없다) -------------------------------
WIN_PW = "windows-pass-xyz"


@pytest.fixture()
def fake_windows(monkeypatch):
    calls = []

    def check(account, password, domain=None):
        calls.append(account)
        _, _, key = winauth.split_account(account, "corp")
        return key if password == WIN_PW else None
    monkeypatch.setattr(winauth, "check_windows", check)
    monkeypatch.setattr(winauth, "available", lambda: True)
    return calls


@pytest.fixture()
def win_app(tmp_path, fake_windows):
    db = tmp_path / "wbs.db"
    with Store(db) as st:
        st.import_xlsx(make_small(tmp_path / "small.xlsx"), CFG)
        auth.add_role(st.db, "admin1", "관리자 1", "admin", PW)
        auth.set_account(st.db, "corp\\boss", "admin1")
    cfg = tmp_path / "wbs.yaml"
    return db, cfg


def _app(db, cfg, unregistered="viewer"):
    import yaml
    from engenius_wbs.config import find_config
    c = yaml.safe_load(find_config().read_text(encoding="utf-8"))
    c["windows_login"] = {"enabled": True, "domain": "corp", "unregistered": unregistered}
    cfg.write_text(yaml.safe_dump(c, allow_unicode=True), encoding="utf-8")
    return create_app(db, cfg)


def test_windows_login(win_app, fake_windows):
    db, cfg = win_app
    app = _app(db, cfg)
    page = TestClient(app).get("/login").text
    assert 'name="mode"' in page and "corp" in page and "조회자로" in page
    # 등록 안 된 계정 → 공용 조회자, 계정은 어디에도 남지 않는다
    c = TestClient(app)
    assert login(c, "someone", WIN_PW, "win").status_code == 303
    s = c.get("/api/status").json()
    assert s["role"] == auth.WIN_VIEWER_LABEL and s["kind"] == "viewer"
    assert c.get("/api/roles").status_code == 403
    with Store(db) as st:
        dump = "\n".join(str(tuple(r)) for t in ("session", "account", "role", "txn") for r in st.db.execute(f"SELECT * FROM {t}"))
    assert "someone" not in dump
    # 등록된 계정 → 그 역할 (표기는 user · DOMAIN\\user · 대소문자 상관없음)
    for who in ("boss", "CORP\\Boss"):
        a = TestClient(app)
        assert login(a, who, WIN_PW, "win").status_code == 303
        assert a.get("/api/status").json()["role"] == "관리자 1" and a.get("/api/roles").status_code == 200
    assert login(TestClient(app), "boss", "wrong", "win").status_code == 401
    # 역할 계정(비상용)은 Windows 확인을 거치지 않는다
    n = len(fake_windows)
    assert login(TestClient(app), "admin1", PW, "role").status_code == 303 and len(fake_windows) == n


def test_windows_unregistered_policy(win_app):
    db, cfg = win_app
    app = _app(db, cfg, "deny")
    assert "조회자로" not in TestClient(app).get("/login").text
    r = login(TestClient(app), "someone", WIN_PW, "win")
    assert r.status_code == 401 and "등록되지 않은" in r.text
    assert login(TestClient(app), "boss", WIN_PW, "win").status_code == 303
    # viewer 설정이어도 공용 조회자 역할을 끄면 등록된 계정만
    app = _app(db, cfg, "viewer")
    with Store(db) as st:
        auth.set_active(st.db, auth.WIN_VIEWER_LOGIN, False)
    assert login(TestClient(app), "someone", WIN_PW, "win").status_code == 401
    # 역할을 끄면 그 역할에 매인 계정도 못 들어온다
    with Store(db) as st:
        auth.set_active(st.db, "admin1", False)
    assert login(TestClient(app), "boss", WIN_PW, "win").status_code == 401


def test_windows_throttle_per_account(win_app):
    db, cfg = win_app
    app = _app(db, cfg)
    for i in range(FAIL_LIMIT):
        assert login(TestClient(app), "boss", "wrong", "win").status_code == 401
    srv = app.state.wbs
    srv.fails.pop("ip:testclient", None)                 # 다른 PC 에서 와도
    assert login(TestClient(app), "boss", WIN_PW, "win").status_code == 429   # 그 계정은 막혀 있다 (사내 계정 잠금 전에)


def test_split_account():
    assert winauth.split_account("Kim", "CORP") == ("CORP", "Kim", "corp\\kim")
    assert winauth.split_account("corp\\kim", "X")[2] == "corp\\kim"
    assert winauth.split_account("kim@corp.example.com")[2] == "corp\\kim"
    for bad in ("", "a\\b\\c", "a@b@c", "x:y", "a<b"):
        with pytest.raises(ValueError):
            winauth.split_account(bad, "D")


@pytest.mark.skipif(not winauth.available(), reason="Windows 전용")
def test_real_windows_rejects_unknown_account():
    assert winauth.check_windows("no-such-user-wbs-test-7f3a", "x-not-a-password", ".") is None
    assert winauth.check_windows("", "x") is None and winauth.check_windows("someone", "") is None


def test_account_cli(tmp_path, capsys):
    db = tmp_path / "wbs.db"
    with Store(db) as st:
        st.import_xlsx(make_small(tmp_path / "small.xlsx"), CFG)
        auth.add_role(st.db, "fan", "FAN 리더", "viewer", PW)
    assert main(["account", "set", "CORP\\Lee", "--role", "fan", "--db", str(db)]) == 0
    assert main(["account", "set", "corp\\lee", "--role", "nope", "--db", str(db)]) == 2
    assert main(["account", "list", "--db", str(db)]) == 0
    assert "corp\\lee" in capsys.readouterr().out
    assert main(["account", "remove", "CORP\\LEE", "--db", str(db)]) == 0
    with Store(db) as st:
        assert auth.list_accounts(st.db) == []


def test_reimport_keeps_roles_and_refreshes_server(tmp_path, monkeypatch, capsys):
    db = tmp_path / "wbs.db"
    src = make_small(tmp_path / "small.xlsx")
    assert main(["import", str(src), "--db", str(db)]) == 0
    with Store(db) as st:
        lv1 = st.db.execute("SELECT uid FROM task WHERE parent_uid IS NULL").fetchone()[0]
        auth.add_role(st.db, "admin1", "관리자 1", "admin", PW)
        auth.add_role(st.db, "fan", "FAN 리더", "leader", PW, lv1)
        auth.set_account(st.db, "corp\boss", "admin1")
        auth.set_active(st.db, auth.WIN_VIEWER_LOGIN, False)
    app = create_app(db)
    c = TestClient(app)
    login(c, "fan")
    first = c.get("/data.js").text
    # 묻는 모드인데 대화 창이 아니면 그만둔다 · 망가진 파일이면 기존 DB 그대로
    assert main(["import", str(src), "--db", str(db), "--ask"]) == 2
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not an excel file")
    assert main(["import", str(bad), "--db", str(db), "--force"]) == 2
    assert not list(tmp_path.glob("wbs.db.bak-*")) and not (tmp_path / "wbs.db.new").exists()
    with Store(db) as st:
        assert st.count() == 12
    # 묻는 모드 + 예 → 덮어씀, 옛 DB 는 .bak, 역할·계정·비밀번호 그대로
    monkeypatch.setattr("sys.stdin.isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", lambda prompt="": "y")
    wb = load_workbook(src)
    wb.active.cell(9, 3).value = "   " * 5 + "B 새 이름"
    wb.save(src)
    assert main(["import", str(src), "--db", str(db), "--ask"]) == 0
    assert len(list(tmp_path.glob("wbs.db.bak-*"))) == 1
    out = capsys.readouterr().out
    assert "역할 2 · Windows 계정 1 를 새 DB 로 옮김" in out and "못 찾은" not in out
    with Store(db) as st:
        assert auth.verify(st.db, "fan", PW).module_uid == st.db.execute("SELECT uid FROM task WHERE parent_uid IS NULL").fetchone()[0]
        assert auth.account_role(st.db, "corp\boss").label == "관리자 1"
        assert auth.windows_viewer(st.db) is None                 # 껐던 공용 조회자도 그대로
        assert st.db.execute("SELECT COUNT(*) FROM session").fetchone()[0] == 0
    # 서버는 판 번호가 같아도(1) 새 DB 를 알아본다 — 로그인은 다시
    assert c.get("/api/status").status_code == 401
    login(c, "fan")
    second = c.get("/data.js").text
    assert second != first and "B 새 이름" in second
