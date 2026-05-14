from __future__ import annotations

from pathlib import Path
from typing import Any, Tuple

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


def place_stamp_pdf(page: fitz.Page, stamp_source: Any, rect: fitz.Rect) -> None:
    # Stempel explizit als PDF-Annotation anlegen (nicht einbetten).
    # Für ein benutzerdefiniertes Aussehen wird ein Bild/Pixmap als Appearance verwendet.
    annot = page.add_stamp_annot(rect, stamp=stamp_source)
    annot.set_info(
        title="PDF-Stamper",
        content="Stempel wurde automatisch platziert.",
    )
    annot.update()
