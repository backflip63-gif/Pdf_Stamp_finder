from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional

import fitz
import numpy as np

from .models import PlacementSettings
from .utils import mm_to_pt


@dataclass
class Candidate:
    rect: fitz.Rect
    occupancy_ratio: float
    score: float
    scale: float


class StampPlacer:
    def __init__(self, settings: PlacementSettings):
        self.settings = settings

    def find_position(
        self,
        page: fitz.Page,
        occupancy_mask: np.ndarray,
        zoom: float,
        stamp_width_pt: float,
        stamp_height_pt: float,
    ) -> Optional[Candidate]:
        margin_pt = mm_to_pt(self.settings.page_margin_mm)
        step_pt = mm_to_pt(self.settings.grid_step_mm)

        scale = 1.0
        best: Optional[Candidate] = None

        while scale >= self.settings.allow_scale_down_to - 1e-9:
            w = stamp_width_pt * scale
            h = stamp_height_pt * scale
            for rect in self._candidate_rects(page.rect, w, h, step_pt, margin_pt):
                occ = self._occupancy_ratio(rect, occupancy_mask, zoom)
                if occ > self.settings.max_occupancy_ratio:
                    continue

                score = self._score(page.rect, rect, occ)
                cand = Candidate(rect=rect, occupancy_ratio=occ, score=score, scale=scale)
                if best is None or cand.score < best.score:
                    best = cand

            if best is not None:
                return best
            scale -= self.settings.scale_step

        return None

    def _candidate_rects(
        self,
        page_rect: fitz.Rect,
        stamp_w_pt: float,
        stamp_h_pt: float,
        step_pt: float,
        margin_pt: float,
    ) -> Iterable[fitz.Rect]:
        x = margin_pt
        while x + stamp_w_pt <= page_rect.width - margin_pt:
            y = margin_pt
            while y + stamp_h_pt <= page_rect.height - margin_pt:
                yield fitz.Rect(x, y, x + stamp_w_pt, y + stamp_h_pt)
                y += step_pt
            x += step_pt

    def _occupancy_ratio(self, rect: fitz.Rect, mask: np.ndarray, zoom: float) -> float:
        h, w = mask.shape
        x0 = max(0, min(w, int(rect.x0 * zoom)))
        y0 = max(0, min(h, int(rect.y0 * zoom)))
        x1 = max(0, min(w, int(rect.x1 * zoom)))
        y1 = max(0, min(h, int(rect.y1 * zoom)))
        if x1 <= x0 or y1 <= y0:
            return 1.0
        area = mask[y0:y1, x0:x1]
        if area.size == 0:
            return 1.0
        return float(area.mean())

    def _score(self, page_rect: fitz.Rect, rect: fitz.Rect, occ: float) -> float:
        # Bonus für Randnähe, aber ohne blind auf unten rechts zu gehen.
        left = rect.x0
        top = rect.y0
        right = page_rect.width - rect.x1
        bottom = page_rect.height - rect.y1
        edge_distance = min(left, top, right, bottom)

        # leichte Präferenz für rechts unten
        right_bottom_bias = right * 0.02 + bottom * 0.02
        return occ * 1000.0 + edge_distance * 0.05 + right_bottom_bias
