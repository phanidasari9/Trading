"""
Persistent watchlists (named ticker universes), JSON file next to this package.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

LIST_NAME_PATTERN = re.compile(r"^[\w\-. ]{1,64}$")
STORE_FILENAME = "watchlists.json"


def store_path() -> Path:
    return Path(__file__).resolve().parent / STORE_FILENAME


def _normalize_name(name: str) -> str:
    return name.strip()


def validate_list_name(name: str) -> str | None:
    n = _normalize_name(name)
    if not n:
        return "Name cannot be empty."
    if not LIST_NAME_PATTERN.match(n):
        return "Use 1–64 characters: letters, numbers, spaces, hyphen, underscore, dot."
    return None


def load_store() -> dict[str, Any]:
    path = store_path()
    if not path.exists():
        return {"lists": {}}
    try:
        with path.open(encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {"lists": {}}
    if not isinstance(data, dict):
        return {"lists": {}}
    lists_raw = data.get("lists")
    if not isinstance(lists_raw, dict):
        lists_raw = {}
    lists: dict[str, list[str]] = {}
    for key, val in lists_raw.items():
        if not isinstance(key, str):
            continue
        kn = _normalize_name(key)
        if not kn:
            continue
        if isinstance(val, list):
            sym = [str(x).strip().upper() for x in val if str(x).strip()]
            lists[kn] = list(dict.fromkeys(sym))
        elif val is None:
            lists[kn] = []
    return {"lists": lists}


def save_store(lists: dict[str, list[str]]) -> None:
    path = store_path()
    payload = {"lists": {k: lists[k] for k in sorted(lists.keys())}}
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
        f.write("\n")


def ensure_default_universe(default_tickers: list[str]) -> dict[str, list[str]]:
    """If store has no Default list, seed it from ``default_tickers`` and persist."""
    data = load_store()
    lists: dict[str, list[str]] = dict(data["lists"])
    if "Default" not in lists:
        lists["Default"] = list(dict.fromkeys(s.upper() for s in default_tickers if s.strip()))
        save_store(lists)
    return lists


def list_names(lists: dict[str, list[str]]) -> list[str]:
    return sorted(lists.keys())


def upsert_watchlist(lists: dict[str, list[str]], name: str, tickers: list[str]) -> dict[str, list[str]]:
    out = {**lists, name: list(dict.fromkeys(tickers))}
    save_store(out)
    return out


def delete_watchlist(lists: dict[str, list[str]], name: str) -> dict[str, list[str]] | None:
    if name not in lists:
        return lists
    if len(lists) <= 1:
        return None
    out = {k: v for k, v in lists.items() if k != name}
    save_store(out)
    return out
