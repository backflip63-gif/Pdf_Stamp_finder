from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .models import PlacementSettings


CONFIG_FILE = Path.home() / ".pdf_stamper_config.json"


def save_settings(settings: PlacementSettings) -> None:
    CONFIG_FILE.write_text(json.dumps(asdict(settings), indent=2, ensure_ascii=False), encoding="utf-8")


def load_settings() -> PlacementSettings:
    if not CONFIG_FILE.exists():
        return PlacementSettings()

    data: dict[str, Any] = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return PlacementSettings(**data)
