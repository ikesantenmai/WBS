"""定義ファイルの雛形。"""

from __future__ import annotations

import shutil
from pathlib import Path

TEMPLATE_DIR = Path(__file__).parent
TEMPLATE_NAMES = ("standard", "minimal")


def template_path(name: str) -> Path:
    path = TEMPLATE_DIR / f"{name}.yaml"
    if not path.exists():
        raise ValueError(f"雛形が見つかりません: {name} ({'/'.join(TEMPLATE_NAMES)})")
    return path


def write_template(name: str, destination) -> Path:
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path(name), destination)
    return destination
