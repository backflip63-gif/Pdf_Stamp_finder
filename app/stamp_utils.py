from __future__ import annotations

from pathlib import Path
from typing import Tuple

import fitz


class StampError(RuntimeError):
    pass


def get_stamp_page_size(stamp_pdf: Path) -> Tuple[float, float]:
    doc = fitz.open(stamp_pdf)
    try:
        if len(doc) == 0:
            raise StampError("Stempel-PDF enthält keine Seiten.")
        rect = doc[0].rect
        return rect.width, rect.height
    finally:
        doc.close()


def place_stamp_pdf(page: fitz.Page, stamp_doc: fitz.Document, rect: fitz.Rect) -> None:
    page.show_pdf_page(rect, stamp_doc, 0, overlay=True)
