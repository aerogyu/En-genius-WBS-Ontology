"""WBS 서버 (v0.5 S2 — 읽기 전용). 도구 DB 하나를 여러 사람이 사내망 브라우저로 본다 (docs/05).

  python -m engenius_wbs serve [--db data/real/wbs.db] [--host 127.0.0.1] [--port 8100]   (scripts\\serve.bat)

화면은 지금 보기 HTML(template_v3)을 그대로 쓰고 데이터만 서버에서 받는다 (data.js — DB 판이 바뀔 때만 다시 계산).
로그인은 역할 계정 (auth.py, `role add` 로 만든다). 모든 역할이 보기·내려받기, 관리자만 역할 목록.
가정 편집 → '서버에 저장' (S3, edit.py — 권한·겹침 거절·이력). 쓰는 패키지: starlette·uvicorn (사내 1.3.1·0.51.0).
S5 (docs/05 14장): ⑤ 판·요청 탭 — 판 확정·비교(snapshot.py) · 백업(backup.py) · 수정 요청·편집 열림/닫힘(collab.py) · 모듈별 엑셀.
"""
from __future__ import annotations

import hashlib
import html
import json
import re
import sqlite3
import tempfile
import threading
import time
from collections import OrderedDict, defaultdict, deque
from pathlib import Path
from urllib.parse import parse_qs, quote

from starlette.applications import Starlette
from starlette.concurrency import run_in_threadpool
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from starlette.routing import Route

from . import auth, backup, collab, snapshot, winauth
from .analyze import analyze
from .config import load_config
from .edit import EditError, apply_edits, history
from .report import write_report
from .rows import apply_row_op
from .store import DATA_KINDS, Store
from .viz import TEMPLATE, payload

COOKIE = "wbs_session"
FAIL_LIMIT, FAIL_WINDOW = 5, 600          # 같은 주소에서 10분에 5번 틀리면 10분 쉼
HEADERS = {"X-Content-Type-Options": "nosniff", "X-Frame-Options": "DENY", "Referrer-Policy": "no-referrer",
           "Content-Security-Policy": "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; "
                                      "img-src 'self' data: blob:; frame-ancestors 'none'; form-action 'self'"}


# 프로그램 판 — 코드를 갈아 끼우면 DB 판이 같아도 브라우저가 옛 data.js 를 쓰지 않게 (캐시 열쇠에 섞는다)
CODE_TAG = hashlib.sha256(b"".join(p.read_bytes() for p in sorted(Path(__file__).parent.glob("*.py")))
                          + TEMPLATE.encode("utf-8")).hexdigest()[:10]


def _rev(db: sqlite3.Connection) -> int:
    """판 번호 = 일정을 바꾼 마지막 저장 (판 확정·편집 열고 닫기는 올리지 않음 — store.DATA_KINDS)."""
    q = ",".join("?" * len(DATA_KINDS))
    return db.execute(f"SELECT COALESCE(MAX(id), 0) FROM txn WHERE kind IN ({q})", DATA_KINDS).fetchone()[0]


class WbsServer:
    """DB 경로·설정과 계산 결과 캐시. 요청마다 연결을 새로 연다 (SQLite 연결은 스레드마다 따로)."""

    def __init__(self, db_path: str | Path, cfg_path: str | Path | None = None, title: str = "WBS"):
        self.db_path = Path(db_path)
        self.cfg = load_config(cfg_path)
        self.title = title
        self._lock = threading.Lock()
        self.write_lock = threading.Lock()          # 저장은 한 번에 하나씩
        self._cache: tuple[int, bytes, dict] | None = None
        self._stamp = None
        wl = self.cfg.get("windows_login") or {}
        self.win = bool(wl.get("enabled", True)) and winauth.available()
        self.win_domain = winauth.default_domain(wl.get("domain") or None)
        self.win_unregistered = wl.get("unregistered", "viewer")
        self.fails: dict[str, deque] = defaultdict(deque)
        self._cmp: OrderedDict = OrderedDict()      # 판 비교 결과 (최근 몇 개)
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
                stamp = (rev, st.meta("imported_at"), st.meta("source_sha256"))      # 새로 들여온 DB 면 판 번호가 같아도 다시
                if self._cache and self._stamp == stamp:
                    return self._cache
                t0 = time.perf_counter()
                m = st.model(self.cfg)
                w = m.wbs
                a = analyze(w, self.cfg)
                p = payload(w, a, self.cfg)
                order = sorted(w.tasks.values(), key=lambda t: t.excel_row)
                p["uid"] = [w.uids[t.no] for t in order]
                p["fl"] = [m.floor.get(t.no) for t in order]
                p["wk"] = [m.work.get(t.no) for t in order]
                rl = {(e.pred, e.succ, e.type, e.lag_days or 0): x for e, x in zip(m.edges, m.relax) if x}
                for e in p["edges"]:
                    e.append(rl.get((e[0], e[1], e[3] or "FS", e[4] or 0), 0))
            p["sha256"] = ""                     # 원본 엑셀과 맞춰 보는 값 — 서버에서는 쓰지 않는다
            p["server"] = {"rev": rev, "seconds": round(time.perf_counter() - t0, 2)}
            js = ("window.__WBS_DATA__=" + json.dumps(p, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/") + ";").encode("utf-8")
            info = {"rev": rev, "tasks": len(w.tasks), "edges": len(a.edges), "source": p["source"], "seconds": p["server"]["seconds"]}
            self._cache = (rev, js, info)
            self._stamp = stamp
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

    # --- 판 비교 · 백업 (S5) -------------------------------------------------------------------------
    def compare(self, a: int, b: int | None) -> dict:
        """판 a → 판 b (None = 지금). 같은 요청은 다시 계산하지 않는다 (지금과 비교면 판 번호가 열쇠에)."""
        with self.connect() as st:
            key = (a, b, st.rev() if b is None else None, st.meta("imported_at"))
            with self._lock:
                if key in self._cmp:
                    self._cmp.move_to_end(key)
                    return self._cmp[key]
            sa = snapshot.open_snapshot(st, a)
            sb = snapshot.open_snapshot(st, b) if b is not None else st
            try:
                res = snapshot.compare(sa, sb, self.cfg)
            finally:
                sa.close()
                if b is not None:
                    sb.close()
            labels = {x["id"]: x["label"] for x in snapshot.list_snapshots(st)}
        res["name_a"] = labels.get(a, f"판 {a}")
        res["name_b"] = labels.get(b, f"판 {b}") if b is not None else "지금"
        with self._lock:
            self._cmp[key] = res
            while len(self._cmp) > 6:
                self._cmp.popitem(last=False)
        return res

    def backup_now(self, kind: str) -> tuple[Path, dict]:
        with self.write_lock:
            return backup.make_backup(self.db_path, self.cfg, kind)

    def backup_tick(self, now=None) -> Path | None:
        """하루 한 번 백업 (서버 안 타이머가 10분마다 부름)."""
        if backup.daily_due(self.db_path, self.cfg, now):
            return self.backup_now("daily")[0]
        return None

    def state(self, role: auth.Role) -> dict:
        with self.connect() as st:
            e = collab.edit_state(st)
            return {"edit_open": e["open"], "edit_note": e["note"], "inbox": collab.inbox_count(st, role)}

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
           '<span class="err" id="srvLock" style="display:none;font-size:12px"></span>'
           '<span class="lnk" id="srvInbox" style="display:none;font-size:12px;cursor:pointer" title="⑤ 판·요청 탭에서 보기"></span>'
           f'<span class="mut" id="srvWho" title="역할 · DB 판 번호">{html.escape(who)} ({role.kind_ko}) · 판 #{rev}</span>'
           '<a class="mut" id="srvExport" href="export.xlsx" title="서버에 저장된 판을 원본 양식 엑셀로 (인명 칸은 비어 있음). 이 브라우저의 가정 편집은 ④ 주경로 탭에서 ‘서버에 저장’ 한 뒤에 들어간다">저장된 판 엑셀 내려받기</a>'
           '<form method="post" action="logout" style="display:inline;margin:0"><button style="font:inherit;font-size:12px">로그아웃</button></form></header>')
    # 30초마다 판 번호·편집 닫힘·받은 요청 수를 물어 머리 줄을 고친다 (⑤ 탭 COLLAB.state 도)
    poll = ("\n<script>(function(){const R=" + str(rev) + ";"
            "function show(s){if(!s)return;const L=document.getElementById('srvLock'),I=document.getElementById('srvInbox');"
            "L.style.display=s.edit_open?'none':'';L.textContent='편집 닫힘'+(s.edit_note?' — '+s.edit_note:'')+' · 관리자만 저장';"
            "I.style.display=s.inbox?'':'none';I.textContent='받은 요청 '+s.inbox;window.__WBS_STATE__=Object.assign(window.__WBS_STATE__||{},s);"
            "if(window.COLLAB)window.COLLAB.state(s)}window.__srvShow=show;"
            "show(window.__WBS_STATE__);document.getElementById('srvInbox').onclick=()=>{const b=document.querySelector('nav button[data-t=co]');if(b)b.click()};"
            "setInterval(()=>fetch('api/status',{cache:'no-store'})"
            ".then(r=>r.ok?r.json():null).then(s=>{if(s&&s.rev!==R)document.getElementById('srvNew').style.display='';show(s)})"
            ".catch(()=>{}),30000)})();</script>\n")
    style = ("<style>#edDown{display:none}</style>\n<script>window.__WBS_ROLE__=" +   # 화면은 이것으로 단추만 숨긴다 — 권한은 서버가 저장 때 검사
             json.dumps({"kind": role.kind, "module_uid": role.module_uid, "label": role.label}, ensure_ascii=False).replace("</", "<\\/") + ";"
             "window.__WBS_STATE__=" + json.dumps(srv.state(role), ensure_ascii=False).replace("</", "<\\/") + ";</script>\n")   # 변경 목록 → apply.bat 흐름은 서버에서는 쓰지 않는다 (저장은 S3)
    poll += ("<script>document.addEventListener('click',e=>{if(e.target.id!=='srvExport')return;let n=0;try{n=ED.size}catch(_){}"
             "if(n&&!confirm('가정 편집 '+n+'건은 아직 서버에 저장되지 않아 이 엑셀에 들어가지 않습니다 (④ 주경로 탭의 ‘서버에 저장’ 뒤에 들어감).\\n저장된 판 그대로 내려받을까요?'))e.preventDefault()});</script>\n")
    return page.replace(old_h, bar).replace(old_d, f'<script src="data.js?r={rev}&amp;v={CODE_TAG}"></script>{style}{poll}<script>\nconst D=window.__WBS_DATA__;')


def _safe(name: str) -> str:
    """파일 이름에 쓸 수 있게 (윈도 금지 글자·공백 → _)."""
    return re.sub(r'[\\/:*?"<>|\s]+', "_", str(name)).strip("_")[:60] or "x"


def _xlsx(data: bytes, name: str) -> Response:
    cd = 'attachment; filename="wbs_export.xlsx"; filename*=UTF-8' + "''" + quote(name)
    return _h(Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       headers={"Content-Disposition": cd}))


def _edit_err(e: EditError) -> Response:
    code = {"denied": 403, "conflict": 409, "locked": 423}.get(e.kind, 400)
    return _h(JSONResponse({"error": str(e), "kind": e.kind, "items": e.items}, status_code=code))


async def _json(request: Request, limit: int = 200_000) -> tuple[dict, Response | None]:
    if not request.headers.get("content-type", "").startswith("application/json"):
        return {}, _h(JSONResponse({"error": "JSON 으로 보낼 것"}, status_code=415))
    try:
        body = json.loads((await request.body())[:limit] or b"{}")
    except ValueError:
        return {}, _h(JSONResponse({"error": "요청을 읽지 못함"}, status_code=400))
    if not isinstance(body, dict):
        return {}, _h(JSONResponse({"error": "요청을 읽지 못함"}, status_code=400))
    return body, None


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
        etag = f'"r{rev}-{CODE_TAG}-{hashlib.sha256(js).hexdigest()[:8]}"'
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
        s.update(await run_in_threadpool(srv.state, role))
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

    def int_param(request: Request, key: str) -> int | None:
        v = request.query_params.get(key)
        if v in (None, "", "now"):
            return None
        if not v.isdigit():
            raise ValueError(f"{key} 는 숫자")
        return int(v)

    def export_from(st: Store, out: Path, module: int | None, tag: str) -> tuple[bytes, str]:
        stem = Path(st.meta("source_name", "wbs.xlsx")).stem
        st.export_xlsx(out, srv.cfg, module)
        mod = ""
        if module is not None:
            r = st.db.execute("SELECT name FROM task WHERE uid = ?", (module,)).fetchone()
            mod = "_" + _safe(r["name"] if r else str(module))
        return out.read_bytes(), f"{stem}_{tag}{mod}.xlsx"

    async def export(request: Request):
        _, err = await need(request)
        if err:
            return err
        try:
            module = int_param(request, "module")

            def make() -> tuple[bytes, str]:
                with tempfile.TemporaryDirectory() as d, srv.connect() as st:
                    return export_from(st, Path(d) / "x.xlsx", module, f"판{_rev(st.db)}")
            data, name = await run_in_threadpool(make)
        except ValueError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=400))
        return _xlsx(data, name)

    # --- S5: 판 확정·비교 ---------------------------------------------------------------------------
    async def snapshots(request: Request):
        role, err = await need(request, admin=request.method == "POST")
        if err:
            return err
        if request.method == "GET":
            def get():
                with srv.connect() as st:
                    return snapshot.list_snapshots(st)
            return _h(JSONResponse(await run_in_threadpool(get)))
        body, bad = await _json(request)
        if bad:
            return bad

        def run():
            with srv.connect() as st:                    # 이름 검사는 백업 전에 (헛 백업을 남기지 않게)
                snapshot.check_label(st, body.get("label"))
            p, _ = srv.backup_now("snapshot")            # 확정 직전 백업 (실패하면 확정하지 않음)
            with srv.write_lock, srv.connect() as st:
                return snapshot.freeze(st, srv.cfg, role, body.get("label"), body.get("note") or "") | {"backup": p.name}
        try:
            return _h(JSONResponse(await run_in_threadpool(run)))
        except EditError as e:
            return _edit_err(e)
        except RuntimeError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=500))

    async def compare_api(request: Request):
        _, err = await need(request)
        if err:
            return err
        try:
            a, b = int_param(request, "a"), int_param(request, "b")
            if a is None:
                raise ValueError("a(비교 기준 판)를 줄 것")
            res = await run_in_threadpool(srv.compare, a, b)
        except ValueError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=400))
        out = {k: v for k, v in res.items() if k != "items"}
        out["items"] = res["items"][:3000]
        out["items_total"] = len(res["items"])
        return _h(JSONResponse(out))

    async def compare_xlsx(request: Request):
        _, err = await need(request)
        if err:
            return err
        try:
            a, b = int_param(request, "a"), int_param(request, "b")
            if a is None:
                raise ValueError("a(비교 기준 판)를 줄 것")

            def make():
                res = srv.compare(a, b)
                with tempfile.TemporaryDirectory() as d:
                    out = snapshot.compare_xlsx(res, Path(d) / "c.xlsx", res["name_a"], res["name_b"])
                    return out.read_bytes(), f"비교_{_safe(res['name_a'])}_{_safe(res['name_b'])}.xlsx"
            data, name = await run_in_threadpool(make)
        except ValueError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=400))
        return _xlsx(data, name)

    def snap_label(st: Store, sid: int) -> str:
        label = next((x["label"] for x in snapshot.list_snapshots(st) if x["id"] == sid), None)
        if label is None:
            raise ValueError(f"없는 판: {sid}")
        return label

    async def snapshot_xlsx(request: Request):
        _, err = await need(request)
        if err:
            return err
        sid = request.path_params["sid"]
        try:
            module = int_param(request, "module")

            def make():
                with tempfile.TemporaryDirectory() as d, srv.connect() as st:
                    label = snap_label(st, sid)
                    mem = snapshot.open_snapshot(st, sid)
                    try:
                        return export_from(mem, Path(d) / "x.xlsx", module, _safe(label))
                    finally:
                        mem.close()
            data, name = await run_in_threadpool(make)
        except ValueError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=400))
        return _xlsx(data, name)

    async def report_xlsx(request: Request):
        _, err = await need(request)
        if err:
            return err
        try:
            sid = int_param(request, "snapshot")

            def make():
                with tempfile.TemporaryDirectory() as d, srv.connect() as st:
                    if sid is None:
                        src, tag = st, f"판{_rev(st.db)}"
                    else:
                        tag = _safe(snap_label(st, sid))
                        src = snapshot.open_snapshot(st, sid)
                    try:
                        w = src.load(srv.cfg)
                        out = write_report(w, analyze(w, srv.cfg), srv.cfg, Path(d) / "r.xlsx")
                        return out.read_bytes(), f"{Path(st.meta('source_name', 'wbs.xlsx')).stem}_{tag}_보고서.xlsx"
                    finally:
                        if src is not st:
                            src.close()
            data, name = await run_in_threadpool(make)
        except ValueError as e:
            return _h(JSONResponse({"error": str(e)}, status_code=400))
        return _xlsx(data, name)

    # --- S5: 수정 요청 · 편집 열림/닫힘 · 백업 ------------------------------------------------------
    async def requests_api(request: Request):
        role, err = await need(request)
        if err:
            return err
        if request.method == "GET":
            def get():
                with srv.connect() as st:
                    return collab.list_requests(st, srv.cfg, role)
            return _h(JSONResponse(await run_in_threadpool(get)))
        body, bad = await _json(request)
        if bad:
            return bad

        def run():
            with srv.write_lock, srv.connect() as st:
                return collab.create_request(st, srv.cfg, role, body)
        try:
            return _h(JSONResponse(await run_in_threadpool(run)))
        except EditError as e:
            return _edit_err(e)

    async def request_act(request: Request):
        role, err = await need(request)
        if err:
            return err
        body, bad = await _json(request)
        if bad:
            return bad

        def run():
            with srv.write_lock, srv.connect() as st:
                return collab.act_request(st, srv.cfg, role, request.path_params["rid"], body.get("action"),
                                          body.get("reply") or "", body.get("result_uid"))
        try:
            return _h(JSONResponse(await run_in_threadpool(run)))
        except EditError as e:
            return _edit_err(e)

    async def lock_api(request: Request):
        role, err = await need(request, admin=True)
        if err:
            return err
        body, bad = await _json(request)
        if bad:
            return bad

        def run():
            with srv.write_lock, srv.connect() as st:
                return collab.set_edit_open(st, role, body.get("open"), body.get("note") or "")
        try:
            return _h(JSONResponse(await run_in_threadpool(run)))
        except EditError as e:
            return _edit_err(e)

    async def backups_api(request: Request):
        _, err = await need(request, admin=True)
        if err:
            return err
        folder = backup.backup_dir(srv.db_path, srv.cfg)
        if request.method == "POST":
            try:
                p, chk = await run_in_threadpool(srv.backup_now, "manual")
            except RuntimeError as e:
                return _h(JSONResponse({"error": str(e)}, status_code=500))
            return _h(JSONResponse({"name": p.name, "tasks": chk["tasks"], "rev": chk["rev"], "folder": str(folder)}))
        files = await run_in_threadpool(backup.list_backups, folder)
        return _h(JSONResponse({"folder": str(folder), "files": [b.as_dict() for b in reversed(files)]}))

    async def save(request: Request):
        role, err = await need(request)
        if err:
            return err
        if not request.headers.get("content-type", "").startswith("application/json"):
            return _h(JSONResponse({"error": "JSON 으로 보낼 것"}, status_code=415))
        try:
            body = json.loads((await request.body())[:2_000_000] or b"{}")
            base, edits = int(body.get("base_rev", -1)), body.get("edits")
        except (ValueError, TypeError, AttributeError):
            return _h(JSONResponse({"error": "요청을 읽지 못함"}, status_code=400))

        def run():
            with srv.write_lock, srv.connect() as st:
                return apply_edits(st, srv.cfg, role, base, edits)
        try:
            res = await run_in_threadpool(run)
        except EditError as e:
            return _edit_err(e)
        return _h(JSONResponse(res))

    async def rows_op(request: Request):
        """줄 동작 하나 (S4) — dry_run 이면 계산만 하고 되돌림 (확인 창 숫자)."""
        role, err = await need(request)
        if err:
            return err
        if not request.headers.get("content-type", "").startswith("application/json"):
            return _h(JSONResponse({"error": "JSON 으로 보낼 것"}, status_code=415))
        try:
            body = json.loads((await request.body())[:200_000] or b"{}")
            base, op, dry = int(body.get("base_rev", -1)), body.get("op"), bool(body.get("dry_run"))
        except (ValueError, TypeError, AttributeError):
            return _h(JSONResponse({"error": "요청을 읽지 못함"}, status_code=400))

        def run():
            with srv.write_lock, srv.connect() as st:
                return apply_row_op(st, srv.cfg, role, base, op, dry)
        try:
            res = await run_in_threadpool(run)
        except EditError as e:
            return _edit_err(e)
        return _h(JSONResponse(res))

    async def hist(request: Request):
        _, err = await need(request)
        if err:
            return err
        try:
            uid = int(request.query_params["uid"]) if request.query_params.get("uid") else None
            limit = max(1, min(int(request.query_params.get("limit", 50)), 500))
        except ValueError:
            return _h(JSONResponse({"error": "uid·limit 는 숫자"}, status_code=400))

        def get():
            with srv.connect() as st:
                return history(st, srv.cfg, uid, limit)
        return _h(JSONResponse(await run_in_threadpool(get)))

    app = Starlette(routes=[
        Route("/", index), Route("/data.js", data_js), Route("/login", login, methods=["GET", "POST"]),
        Route("/logout", logout, methods=["POST"]), Route("/api/status", status), Route("/api/roles", roles),
        Route("/export.xlsx", export), Route("/api/save", save, methods=["POST"]), Route("/api/rows", rows_op, methods=["POST"]), Route("/api/history", hist),
        Route("/api/snapshots", snapshots, methods=["GET", "POST"]), Route("/api/compare", compare_api), Route("/compare.xlsx", compare_xlsx),
        Route("/snapshot/{sid:int}.xlsx", snapshot_xlsx), Route("/report.xlsx", report_xlsx),
        Route("/api/requests", requests_api, methods=["GET", "POST"]), Route("/api/requests/{rid:int}", request_act, methods=["POST"]),
        Route("/api/lock", lock_api, methods=["POST"]), Route("/api/backups", backups_api, methods=["GET", "POST"]),
    ])
    app.state.wbs = srv
    return app
