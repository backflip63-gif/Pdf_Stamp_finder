from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import fitz


class FormFillerError(RuntimeError):
    pass


class PDFFormFiller:
    def list_fields(self, pdf_path: Path) -> List[dict]:
        doc = fitz.open(pdf_path)
        try:
            fields: List[dict] = []
            for page_index in range(len(doc)):
                page = doc[page_index]
                widgets = page.widgets()
                if not widgets:
                    continue
                for widget in widgets:
                    fields.append(
                        {
                            "page": page_index,
                            "name": widget.field_name,
                            "label": widget.field_label or widget.field_name,
                            "value": widget.field_value or "",
                            "type": str(widget.field_type),
                        }
                    )
            return fields
        finally:
            doc.close()

    def fill_form(self, template_pdf: Path, output_pdf: Path, field_values: Dict[str, str]) -> None:
        doc = fitz.open(template_pdf)
        try:
            found_names: set[str] = set()
            for page in doc:
                widgets = page.widgets()
                if not widgets:
                    continue
                for widget in widgets:
                    if widget.field_name in field_values:
                        widget.field_value = field_values[widget.field_name]
                        widget.update()
                        found_names.add(widget.field_name)

            missing = set(field_values.keys()) - found_names
            if missing:
                raise FormFillerError(f"Formularfelder nicht gefunden: {', '.join(sorted(missing))}")

            # Vereinfachter Flatten-Schritt: speichern und danach nochmals öffnen.
            # In vielen Fällen reicht das für die Nutzung als Stempelvorlage.
            temp_bytes = doc.tobytes(garbage=4, deflate=True)
        finally:
            doc.close()

        out = fitz.open("pdf", temp_bytes)
        try:
            for page in out:
                widgets = page.widgets()
                if widgets:
                    for widget in widgets:
                        try:
                            page.delete_widget(widget)
                        except Exception:
                            pass
            out.save(output_pdf, garbage=4, deflate=True)
        finally:
            out.close()
