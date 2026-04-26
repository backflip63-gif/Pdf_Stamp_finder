from __future__ import annotations

from dataclasses import dataclass
from typing import List, Tuple

import fitz
import numpy as np
from PIL import Image, ImageFilter


@dataclass
class PageAnalysis:
    gray: np.ndarray
    occupancy_mask: np.ndarray
    zoom: float
    object_rects: List[fitz.Rect]


class PageAnalyzer:
    def __init__(self, render_dpi: int = 180, whiteness_threshold: int = 245, dilation_px: int = 3):
        self.render_dpi = render_dpi
        self.whiteness_threshold = whiteness_threshold
        self.dilation_px = dilation_px

    def analyze(self, page: fitz.Page) -> PageAnalysis:
        gray, zoom = self._render_page_gray(page)
        mask = self._build_mask(gray)
        object_rects = self._extract_object_rects(page)
        mask = self._paint_object_rects(mask, object_rects, zoom)
        mask = self._dilate(mask, self.dilation_px)
        return PageAnalysis(gray=gray, occupancy_mask=mask, zoom=zoom, object_rects=object_rects)

    def _render_page_gray(self, page: fitz.Page) -> Tuple[np.ndarray, float]:
        zoom = self.render_dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("L")
        return np.array(img), zoom

    def _build_mask(self, gray: np.ndarray) -> np.ndarray:
        mask = (gray <= self.whiteness_threshold).astype(np.uint8)
        return mask

    def _extract_object_rects(self, page: fitz.Page) -> List[fitz.Rect]:
        rects: List[fitz.Rect] = []

        blocks = page.get_text("blocks")
        for block in blocks:
            x0, y0, x1, y1 = block[:4]
            rects.append(fitz.Rect(x0, y0, x1, y1))

        try:
            images = page.get_image_info(xrefs=True)
            for info in images:
                bbox = info.get("bbox")
                if bbox:
                    rects.append(fitz.Rect(bbox))
        except Exception:
            pass

        try:
            drawings = page.get_drawings()
            for d in drawings:
                rect = d.get("rect")
                if rect:
                    rects.append(rect)
        except Exception:
            pass

        return rects

    def _paint_object_rects(self, mask: np.ndarray, rects: List[fitz.Rect], zoom: float) -> np.ndarray:
        h, w = mask.shape
        out = mask.copy()
        for rect in rects:
            x0 = max(0, min(w, int(rect.x0 * zoom)))
            y0 = max(0, min(h, int(rect.y0 * zoom)))
            x1 = max(0, min(w, int(rect.x1 * zoom)))
            y1 = max(0, min(h, int(rect.y1 * zoom)))
            if x1 > x0 and y1 > y0:
                out[y0:y1, x0:x1] = 1
        return out

    def _dilate(self, mask: np.ndarray, iterations: int) -> np.ndarray:
        if iterations <= 0:
            return mask

        img = Image.fromarray((mask * 255).astype(np.uint8), mode="L")
        for _ in range(iterations):
            img = img.filter(ImageFilter.MaxFilter(3))
        return (np.array(img) > 0).astype(np.uint8)
