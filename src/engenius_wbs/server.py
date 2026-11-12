"""WBS 서버 (v0.5 S2 — 읽기 전용). 도구 DB 하나를 여러 사람이 사내망 브라우저로 본다 (docs/05).

  python -m engenius_wbs serve [--db data/real/wbs.db] [--host 127.0.0.1] [--port 8100]   (scripts\\serve.bat)

화면은 지금 보기 HTML(template_v3)을 그대로 쓰고 데이터만 서버에서 받는다 (data.js — DB 판이 바뀔 때만 다시 계산).
로그인은 역할 계정 (auth.py, `role add` 로 만든다). 모든 역할이 보기·내려받기, 관리자만 역할 목록.
가정 편집은 지금처럼 그 브라우저 안에서만 (저장은 S3). 쓰는 패키지: starlette·uvicorn (사내 1.3.1·0.51.0).
"""
from __future__ import annotations

import html
import json
import sqlite3
import tempfile
import threading
import time
from collections import defaultdict, deque
from pathlib import Path
from urllib.parse import parse_qs, quote

from starlette.applications import Starlette
from starlette.concurrency import run_in_threadpool
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from starlette.routing import Route

from . import auth, winauth
from .analyze import analyze
from .config import load_config
from .store import Store
from .viz import TEMPLATE, payload

COOKIE = "wbs_session"
FAIL_LIMIT, FAIL_WINDOW = 5, 600          # 같은 주소에서 10분에 5번 틀리면 10분 쉼
HEADERS = {"X-Content-Type-Options": "nosniff", "X-Frame-Options": "DENY", "Referrer-Policy": "no-referrer",
           "Content-Security-Policy": "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; "
                                      "img-src 'self' data: blob:; frame-ancestors 'none'; form-action 'self'"}


def _rev(db: sqlite3.Connection) -> int:
    return db.execute("SELECT COALESCE(MAX(id), 0) FROM txn").fetchone()[0]


class WbsServer:
    """DB 경로·설정과 계산 결과 캐시. 요청마다 연결을 새로 연다 (SQLite 연결은 스레드마다 따로)."""

    def __init__(self, db_path: str | Path, cfg_path: str | Path | None = None, title: str = "WBS"):
        self.db_path = Path(db_path)
        self.cfg = load_config(cfg_path)
        self.title = title
        self._lock = threading.Lock()
        self._cache: tuple[int, bytes, dict] | None = None
        wl = self.cfg.get("windows_login") or {}
        self.win = bool(wl.get("enabled", True)) and winauth.available()
        self.win_domain = winauth.default_domain(wl.get("domain") or None)
        self.win_unregistered = wl.get("unregistered", "viewer")
        self.fails: dict[str, deque] = defaultdict(deque)
        with self.connect() as st:
            auth.ensure(st.db)
            if not st.count():
                raise ValueError(f"업무가 없는 DB 다 (먼저 import): {self.db_path}")

    def connect(self) -> Store:
        if not self.db_path.exists():
            raise ValueError(f"DB 가 없다: {self.db_path}")
        return Store(self.db_path)

    # --- 데이터 (판이 바뀔 때만 다시 계산) ---------------------------------------------------------
    def data(self) -> tuple[int, bytes, dict]:
        with self._lock:
            with self.connect() as st:
                rev = _rev(st.db)
                if self._cache and self._cache[0] == rev:
                    return self._cache
                t0 = time.perf_counter()
                w = st.load(self.cfg)
                a = analyze(w, self.cfg)
                p = payload(w, a, self.cfg)
            p["sha256"] = ""                     # 원본 엑셀과 맞춰 보는 값 — 서버에서는 쓰지 않는다
            p["server"] = {"rev": rev, "seconds": round(time.perf_counter() - t0, 2)}
            js = ("window.__WBS_DATA__=" + json.dumps(p, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/") + ";").encode("utf-8")
            info = {"rev": rev, "tasks": len(w.tasks), "edges": len(a.edges), "source": p["source"], "seconds": p["server"]["seconds"]}
            self._cache = (rev, js, info)
            return self._cache

    # --- 로그인 -------------------------------------------------------------------------------------
    def role_of(self, request: Request) -> auth.Role | None:
        with self.connect() as st:
            return auth.session_role(st.db, request.cookies.get(COOKIE))

    def blocked(self, key: str) -> bool:
        q = self.fails[key]
        now = time.time()
        while q and q[0] < now - FAIL_WINDOW:
            q.popleft()
        return len(q) >= FAIL_LIMIT

    def module_name(self, uid: int | None) -> str:
        if uid is None:
            return ""
        with self.connect() as st:
            r = st.db.execute("SELECT name FROM task WHERE uid = ?", (uid,)).fetchone()
        return r["name"].strip() if r else f"(없는 업무 {uid})"


def _h(resp: Response, cache: str = "no-store") -> Response:
    for k, v in HEADERS.items():
        resp.headers[k] = v
    resp.headers["Cache-Control"] = cache
    return resp


def _login_page(title: str, msg: str = "", status: int = 200, win: bool = False, domain: str = "", https: bool = False,
                open_view: bool = True) -> Response:
    m = f'<p class="err">{html.escape(msg)}</p>' if msg else ""
    if win:
        head = (f'<p>Windows 계정(사내 PC 로그인)으로 — 계정만 적으면 {html.escape(domain)}\\계정.'
                + (' 등록 안 된 계정은 조회자로 들어간다.' if open_view else '') + '</p>')
        mode = ('<label for="m">로그인 방식</label><select id="m" name="mode"><option value="win">Windows 계정</option>'
                '<option value="role">도구 역할 계정 (비상용)</option></select>')
        uid = "계정"
    else:
        head, mode, uid = "<p>역할 계정으로 로그인 (관리자가 발급)</p>", '<input type="hidden" name="mode" value="role">', "로그인 ID"
    warn = "" if https else '<p class="warn">이 접속은 암호화되지 않는다 (HTTP) — 사내망 안에서만 쓸 것.</p>'
    page = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)} — 로그인</title><style>
:root{{--bg:#fafafa;--fg:#1f2328;--mut:#6b7280;--line:#d0d7de;--card:#fff;--acc:#2563eb;--hot:#dc2626}}
@media (prefers-color-scheme:dark){{:root{{--bg:#16181d;--fg:#e6e6e6;--mut:#9aa0a6;--line:#363b44;--card:#1f2228;--acc:#60a5fa;--hot:#f87171}}}}
body{{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 "Malgun Gothic","Apple SD Gothic Neo",sans-serif;display:flex;min-height:100vh;align-items:center;justify-content:center}}
form{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:24px 28px;width:min(340px,calc(100vw - 32px))}}
h1{{font-size:17px;margin:0 0 4px}}p{{margin:0 0 16px;color:var(--mut);font-size:12px}}label{{display:block;margin:10px 0 4px}}
select,input{{width:100%;box-sizing:border-box;padding:7px 9px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg);font:inherit}}
button{{margin-top:18px;width:100%;padding:8px;border:0;border-radius:6px;background:var(--acc);color:#fff;font:inherit;cursor:pointer}}
.err{{color:var(--hot);font-size:13px;margin:10px 0 0}}.warn{{margin:14px 0 0;font-size:11px}}
</style></head><body><form method="post" action="login" autocomplete="off">
<h1>{html.escape(title)}</h1>{head}{mode}
<label for="u">{uid}</label><input id="u" name="login" required maxlength="110" autofocus>
<label for="p">비밀번호</label><input id="p" name="password" type="password" required maxlength="200">
{m}<button type="submit">로그인</button>{warn}</form></body></html>"""
    return _h(HTMLResponse(page, status_code=status))


def view_page(srv: WbsServer, role: auth.Role, rev: int) -> str:
    page = TEMPLATE.replace("__TITLE__", html.escape(srv.title))
    old_d, old_h = "<script>\nconst D=__DATA__;", '<span class="mut" id="meta"></span></header>'
    if page.count(old_d) != 1 or page.count(old_h) != 1:
        raise RuntimeError("보기 틀이 바뀌었다 — server.py 의 끼워 넣는 자리를 고칠 것")
    who = role.label + (f" · {srv.module_name(role.module_uid)}" if role.module_uid else "")
    bar = ('<span class="mut" id="meta"></span><span style="flex:1"></span>'
           '<span class="mut" id="srvNew" style="display:none;color:var(--hot)">새 판이 있음 — <a href="./">새로 고침</a></span>'
           f'<span class="mut" id="srvWho" title="역할 · DB 판 번호">{html.escape(who)} ({role.kind_ko}) · 판 #{rev}</span>'
           '<a class="mut" id="srvExport" href="export.xlsx" title="서버에 저장된 판을 원본 양식 엑셀로 (인명 칸은 비어 있음). 이 브라우저의 가정 편집은 들어가지 않는다 — 저장은 다음 단계(S3)">저장된 판 엑셀 내려받기</a>'
           '<form method="post" action="logout" style="display:inline;margin:0"><button style="font:inherit;font-size:12px">로그아웃</button></form></header>')
    poll = ("\n<script>(function(){const R=" + str(rev) + ";setInterval(()=>fetch('api/status',{cache:'no-store'})"
            ".then(r=>r.ok?r.json():null).then(s=>{if(s&&s.rev!==R)document.getElementById('srvNew').style.display=''})"
            ".catch(()=>{}),30000)})();</script>\n")
    style = "<style>#edDown{display:none}</style>\n"   # 변경 목록 → apply.bat 흐름은 서버에서는 쓰지 않는다 (저장은 S3)
    poll += ("<script>document.addEventListener('click',e=>{if(e.target.id!=='srvExport')return;let n=0;try{n=ED.size}catch(_){}"
             "if(n&&!confirm('가정 편집 '+n+'건은 아직 서버에 저장되지 않아 이 엑셀에 들어가지 않습니다 (저장 기능은 다음 단계).\\n저장된 판 그대로 내려받을까요?'))e.preventDefault()});</script>\n")
    return page.replace(old_h, bar).replace(old_d, f'<script src="data.js?r={rev}"></script>{poll}{style}<script>\nconst D=window.__WBS_DATA__;')


def create_app(db_path: str | Path, cfg_path: str | Path | None = None, title: str = "WBS") -> Starlette:
    srv = WbsServer(db_path, cfg_path, title)

    async def need(request: Request, admin: bool = False):
        role = await run_in_threadpool(srv.role_of, request)
        if role is None:
            return None, _h(JSONResponse({"error": "로그인이 필요하다"}, status_code=401))
        if admin and not role.is_admin:
            return None, _h(JSONResponse({"error": "관리자만"}, status_code=403))
        return role, None

    async def index(request: Request):
        role = await run_in_threadpool(srv.role_of, request)
        if role is None:
            return _h(RedirectResponse("login", status_code=303))
        rev, _, _ = await run_in_threadpool(srv.data)
        page = await run_in_threadpool(view_page, srv, role, rev)
        return _h(HTMLResponse(page))

    async def data_js(request: Request):
        _, err = await need(request)
        if err:
            return err
        rev, js, _ = await run_in_threadpool(srv.data)
        etag = f'"r{rev}"'
        if request.headers.get("if-none-match") == etag:
            return _h(Response(status_code=304, headers={"ETag": etag}), "private, no-cache")
        return _h(Response(js, media_type="text/javascript; charset=utf-8", headers={"ETag": etag}), "private, no-cache")

    async def login(request: Request):
        https = request.url.scheme == "https"

        def page(msg: str = "", status: int = 200) -> Response:
            return _login_page(srv.title, msg, status, srv.win, srv.win_domain, https, srv.win_unregistered == "viewer")
        if request.method == "GET":
            return page()
        body = (await request.body())[:4096].decode("utf-8", "replace")
        form = {k: v[0] for k, v in parse_qs(body, keep_blank_values=True).items()}
        who = form.get("login", "").strip()
        win = srv.win and form.get("mode", "win") != "role"
        ip = "ip:" + (request.client.host if request.client else "-")
        acct = "acct:" + who.lower()                     # 계정별로도 막는다 (사내 계정 잠금 전에)
        if srv.blocked(ip) or srv.blocked(acct):
            return page("여러 번 틀려 10분 동안 막혔습니다. 잠시 뒤 다시 하세요.", 429)

        def check():
            if win:
                key = winauth.check_windows(who, form.get("password", ""), srv.win_domain)
                if key is None:
                    return None, None, "계정 또는 비밀번호가 맞지 않습니다."
                with srv.connect() as st:
                    role = auth.account_role(st.db, key, srv.win_unregistered)
                    if role is None:
                        return None, None, "이 도구에 등록되지 않은 계정입니다. 관리자에게 문의하세요."
                    return role, auth.new_session(st.db, role), ""
            with srv.connect() as st:
                role = auth.verify(st.db, who, form.get("password", ""))
                return (role, auth.new_session(st.db, role), "") if role else (None, None, "로그인 ID 또는 비밀번호가 맞지 않습니다.")
        role, token, msg = await run_in_threadpool(check)
        if role is None:
            now = time.time()
            srv.fails[ip].append(now)
            srv.fails[acct].append(now)
            return page(msg, 401)
        srv.fails.pop(ip, None)
        srv.fails.pop(acct, None)
        resp = RedirectResponse("./", status_code=303)
        resp.set_cookie(COOKIE, token, max_age=int(auth.SESSION_HOURS * 3600), path="/", httponly=True, samesite="strict",
                        secure=request.url.scheme == "https")
        return _h(resp)

    async def logout(request: Request):
        def drop():
            with srv.connect() as st:
                auth.drop_session(st.db, request.cookies.get(COOKIE))
        await run_in_threadpool(drop)
        resp = RedirectResponse("login", status_code=303)
        resp.delete_cookie(COOKIE, path="/")
        return _h(resp)

    async def status(request: Request):
        role, err = await need(request)
        if err:
            return err

        def get():
            with srv.connect() as st:
                return {"rev": _rev(st.db), "tasks": st.count()}
        s = await run_in_threadpool(get)
        return _h(JSONResponse({**s, "role": role.label, "kind": role.kind}))

    async def roles(request: Request):
        _, err = await need(request, admin=True)
        if err:
            return err

        def get():
            with srv.connect() as st:
                rows = auth.list_roles(st.db, include_off=True)
                n = dict(st.db.execute("SELECT role_id, COUNT(*) FROM account GROUP BY role_id").fetchall())
            return [{"login": r["login"], "label": r["label"], "kind": r["kind"], "module": srv.module_name(r["module_uid"]),
                     "active": bool(r["active"]), "accounts": n.get(r["id"], 0)} for r in rows]
        return _h(JSONResponse(await run_in_threadpool(get)))

    async def export(request: Request):
        _, err = await need(request)
        if err:
            return err

        def make() -> tuple[bytes, str]:
            with tempfile.TemporaryDirectory() as d:
                with srv.connect() as st:
                    stem = Path(st.meta("source_name", "wbs.xlsx")).stem
                    out = st.export_xlsx(Path(d) / "x.xlsx", srv.cfg)
                    name = f"{stem}_판{_rev(st.db)}.xlsx"
                return out.read_bytes(), name
        data, name = await run_in_threadpool(make)
        cd = 'attachment; filename="wbs_export.xlsx"; filename*=UTF-8' + "''" + quote(name)
        return _h(Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           headers={"Content-Disposition": cd}))

    app = Starlette(routes=[
        Route("/", index), Route("/data.js", data_js), Route("/login", login, methods=["GET", "POST"]),
        Route("/logout", logout, methods=["POST"]), Route("/api/status", status), Route("/api/roles", roles),
        Route("/export.xlsx", export),
    ])
    app.state.wbs = srv
    return app
