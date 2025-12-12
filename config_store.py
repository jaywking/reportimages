"""Lightweight persistence for UI settings."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

CONFIG_PATH = Path(__file__).resolve().parent / "settings.json"


def load_settings() -> Dict[str, Any]:
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as fp:
            return json.load(fp)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}


def save_settings(data: Dict[str, Any]) -> None:
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as fp:
            json.dump(data, fp, indent=2)
    except OSError:
        # Silently ignore persistence failures to avoid breaking the UI.
        return
