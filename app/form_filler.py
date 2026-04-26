from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import fitz

PDF_FIELD_FLAG_MULTILINE = 1 << 12


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
                            "is_multiline": bool((getattr(widget, "field_flags", 0) or 0) & PDF_FIELD_FLAG_MULTILINE),
                        }
                    )
            return fields
        finally:
            doc.close()

    def fill_form(self, template_pdf: Path, output_pdf: Path, field_values: Dict[str, str], flatten: bool = True) -> None:
        doc = fitz.open(template_pdf)
        try:
            found_names: set[str] = set()
            for page in doc:
                widgets = page.widgets()
                if not widgets:
                    continue
                for widget in widgets:
                    if widget.field_name in field_values:
                        value = field_values[widget.field_name].replace("\r\n", "\n").replace("\r", "\n")
                        if "\n" in value:
                            flags = getattr(widget, "field_flags", 0) or 0
                            widget.field_flags = flags | PDF_FIELD_FLAG_MULTILINE
                        widget.field_value = value
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
            if flatten:
                for page in out:
                    widgets = page.widgets()
                    if not widgets:
                        continue
                    for widget in widgets:
                        try:
                            page.delete_widget(widget)
                        except Exception:
                            continue
            out.save(output_pdf, garbage=4, deflate=True)
        finally:
            out.close()
