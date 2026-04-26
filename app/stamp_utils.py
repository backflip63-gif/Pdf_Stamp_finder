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
    # Stempel explizit als PDF-Annotation anlegen (nicht einbetten).
    _ = stamp_doc  # aktuell nur für API-Kompatibilität / zukünftige Erweiterung
    annot = page.add_stamp_annot(rect)
    annot.set_info(
        title="PDF-Stamper",
        content="Stempel wurde automatisch platziert.",
    )
    annot.update()
