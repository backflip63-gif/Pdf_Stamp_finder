from __future__ import annotations

from pathlib import Path

MM_TO_PT = 72.0 / 25.4


def mm_to_pt(mm: float) -> float:
    return mm * MM_TO_PT


def pt_to_mm(pt: float) -> float:
    return pt / MM_TO_PT


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def build_output_path(input_pdf: Path, output_dir: Path, suffix: str) -> Path:
    return output_dir / f"{input_pdf.stem}{suffix}{input_pdf.suffix}"
