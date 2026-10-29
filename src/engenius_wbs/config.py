"""설정 읽기 — config/wbs.yaml (사내 조정본 data/real/config/wbs.yaml 이 있으면 그쪽)."""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = REPO_ROOT / "config"
REAL_CONFIG_DIR = REPO_ROOT / "data" / "real" / "config"


def find_config(name: str = "wbs.yaml", explicit: str | Path | None = None) -> Path:
    if explicit:
        return Path(explicit)
    real = REAL_CONFIG_DIR / name
    return real if real.exists() else CONFIG_DIR / name


def load_config(path: str | Path | None = None) -> dict:
    with open(find_config(explicit=path), encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    cfg["levels"] = {int(k): v for k, v in cfg["levels"].items()}
    return cfg


def nfkc(s) -> str:
    return unicodedata.normalize("NFKC", "" if s is None else str(s))


def norm_header(h) -> str:
    return re.sub(r"[\s()\[\]*<>]", "", nfkc(h)).lower()
