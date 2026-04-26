from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional


@dataclass
class PlacementSettings:
    stamp_width_mm: float = 60.0
    stamp_height_mm: float = 30.0
    grid_step_mm: float = 5.0
    page_margin_mm: float = 8.0
    render_dpi: int = 180
    whiteness_threshold: int = 245
    max_occupancy_ratio: float = 0.03
    dilation_px: int = 3
    allow_scale_down_to: float = 0.85
    scale_step: float = 0.05
    process_mode: str = "all"  # all | first | last


@dataclass
class StampTemplateConfig:
    template_pdf: Path
    filled_output_pdf: Path
    field_values: Dict[str, str] = field(default_factory=dict)


@dataclass
class BatchJobConfig:
    input_dir: Path
    output_dir: Path
    stamp_pdf: Path
    settings: PlacementSettings
    output_suffix: str = "_gestempelt"


@dataclass
class PlacementResult:
    page_index: int
    rect: Optional[tuple[float, float, float, float]]
    scale: float
    occupancy_ratio: float
    status: str
    note: str = ""


@dataclass
class FileProcessResult:
    input_file: Path
    output_file: Optional[Path]
    page_results: List[PlacementResult] = field(default_factory=list)
    success: bool = False
    error: str = ""
