"""Windows 계정 확인 (v0.5 S2) — 서버 PC 의 Windows 에게 '이 계정·비밀번호가 맞나' 만 묻는다. 표준 라이브러리 ctypes (새 반입 0).

LogonUserW(LOGON32_LOGON_NETWORK) — 사내 PC 로그인과 같은 확인. 서버 PC 가 사내 도메인에 가입돼 있어야 다른 사람 계정을 확인할 수 있다.
비밀번호는 저장하지 않는다. 확인이 끝나면 받은 토큰을 바로 닫는다. Windows 가 아니면 늘 실패 (None).
틀린 비밀번호는 사내 계정 잠금 횟수에 같이 세어진다 — 서버의 로그인 차단(주소별·계정별)이 그 앞에서 막는다.

계정 표기: 'user' (기본 도메인 = 설정 windows_login.domain, 비면 서버를 띄운 사용자의 도메인) · 'DOMAIN\\user' · 'user@domain'.
돌려주는 값은 비교용 표기 'domain\\user' (소문자) — 계정 → 역할 표의 열쇠.
"""
from __future__ import annotations

import os
import re
import sys

LOGON32_LOGON_NETWORK = 3
LOGON32_PROVIDER_DEFAULT = 0
_ACCOUNT_RE = re.compile(r"^[^\x00-\x1f\"/\[\]:;|=,+*?<>]{1,104}$")


def available() -> bool:
    return sys.platform == "win32"


def default_domain(cfg_domain: str | None = None) -> str:
    return (cfg_domain or os.environ.get("USERDOMAIN") or ".").strip()


def split_account(account: str, domain: str | None = None) -> tuple[str | None, str, str]:
    """(LogonUser 에 줄 도메인, 사용자, 비교용 'domain\\user' 소문자). 형식이 이상하면 ValueError."""
    a = (account or "").strip()
    if not a or not _ACCOUNT_RE.match(a.replace("\\", "").replace("@", "")) or a.count("\\") > 1 or a.count("@") > 1:
        raise ValueError("계정 형식이 이상하다")
    if "\\" in a:
        dom, user = a.split("\\", 1)
    elif "@" in a:
        user, dom = a.split("@", 1)
        dom = dom.split(".")[0]                 # 'user@corp.example' → 비교용은 corp\user (NetBIOS 이름과 같다고 가정 — 사내 확인)
        if not user or not dom:
            raise ValueError("계정 형식이 이상하다")
        return None, a, f"{dom}\\{user}".lower()
    else:
        dom, user = default_domain(domain), a
    if not dom or not user:
        raise ValueError("계정 형식이 이상하다")
    return dom, user, f"{dom}\\{user}".lower()


def check_windows(account: str, password: str, domain: str | None = None) -> str | None:
    """맞으면 비교용 계정 표기, 아니면 None."""
    if not available() or not password:
        return None
    try:
        dom, user, key = split_account(account, domain)
    except ValueError:
        return None
    import ctypes
    from ctypes import wintypes

    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    logon = advapi32.LogonUserW
    logon.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(wintypes.HANDLE)]
    logon.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    token = wintypes.HANDLE()
    ok = logon(user, dom, password, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, ctypes.byref(token))
    if not ok:
        return None
    kernel32.CloseHandle(token)
    return key
