from __future__ import annotations

from pathlib import Path
import os
import shutil
import subprocess
import tempfile
from typing import Dict

import fitz
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QIcon, QImage, QPainter, QPixmap
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSpinBox,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QGraphicsPixmapItem,
    QGraphicsScene,
    QGraphicsView,
    QListWidget,
    QListWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QSplitter,
)

from .batch_processor import BatchProcessor
from .config import load_settings, save_settings
from .form_filler import PDFFormFiller
from .models import BatchJobConfig, PlacementSettings
from .stamp_utils import get_stamp_page_size
from .utils import ensure_dir, pt_to_mm


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PDF-Stamper")
        self.resize(1100, 800)

        self.settings = load_settings()
        self.form_filler = PDFFormFiller()
        self.form_inputs: Dict[str, QWidget] = {}

        self.template_pdf_path: Path | None = None
        self.filled_stamp_pdf_path: Path | None = None
        self.input_dir_path: Path | None = None
        self.selected_input_file: Path | None = None

        self._build_ui()
        self._load_settings_into_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)

        layout.addWidget(self._build_template_group())
        layout.addWidget(self._build_processing_group())
        layout.addWidget(self._build_log_group(), stretch=1)

        self.setCentralWidget(root)

    def _build_template_group(self) -> QGroupBox:
        group = QGroupBox("1. Stempelauswahl")
        layout = QVBoxLayout(group)

        row = QHBoxLayout()
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        btn_select_template = QPushButton("Stempel auswählen")
        btn_select_template.clicked.connect(self.select_template_pdf)
        row.addWidget(self.template_path_edit, stretch=1)
        row.addWidget(btn_select_template)
        layout.addLayout(row)

        self.fields_area = QScrollArea()
        self.fields_area.setWidgetResizable(True)
        self.fields_area.setFixedHeight(180)
        self.fields_area.setMaximumWidth(620)
        self.fields_container = QWidget()
        self.fields_form = QFormLayout(self.fields_container)
        self.fields_area.setWidget(self.fields_container)
        layout.addWidget(self.fields_area)

        self.stamp_output_label = QLabel("Noch kein erzeugtes Stempel-PDF.")
        layout.addWidget(self.stamp_output_label)

        return group

    def _build_processing_group(self) -> QGroupBox:
        group = QGroupBox("2. Stapelverarbeitung")
        layout = QGridLayout(group)

        self.input_dir_edit = QLineEdit()
        self.input_dir_edit.setReadOnly(True)

        self.stamp_width_spin = QDoubleSpinBox()
        self.stamp_width_spin.setRange(1.0, 500.0)
        self.stamp_width_spin.setSuffix(" mm")
        self.stamp_width_spin.setDecimals(1)

        self.stamp_height_spin = QDoubleSpinBox()
        self.stamp_height_spin.setRange(1.0, 500.0)
        self.stamp_height_spin.setSuffix(" mm")
        self.stamp_height_spin.setDecimals(1)

        self.grid_step_spin = QDoubleSpinBox()
        self.grid_step_spin.setRange(1.0, 100.0)
        self.grid_step_spin.setSuffix(" mm")
        self.grid_step_spin.setDecimals(1)

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(72, 600)

        self.threshold_spin = QSpinBox()
        self.threshold_spin.setRange(0, 255)

        self.max_occ_spin = QDoubleSpinBox()
        self.max_occ_spin.setRange(0.0, 1.0)
        self.max_occ_spin.setSingleStep(0.01)
        self.max_occ_spin.setDecimals(3)

        self.dilation_spin = QSpinBox()
        self.dilation_spin.setRange(0, 20)

        self.scale_down_spin = QDoubleSpinBox()
        self.scale_down_spin.setRange(0.1, 1.0)
        self.scale_down_spin.setSingleStep(0.05)
        self.scale_down_spin.setDecimals(2)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["all", "first", "last"])

        self.advanced_values = {
            "grid_step_mm": self.settings.grid_step_mm,
            "render_dpi": self.settings.render_dpi,
            "whiteness_threshold": self.settings.whiteness_threshold,
            "max_occupancy_ratio": self.settings.max_occupancy_ratio,
            "dilation_px": self.settings.dilation_px,
            "allow_scale_down_to": self.settings.allow_scale_down_to,
            "process_mode": self.settings.process_mode,
        }

        btn_run = QPushButton("Stapelverarbeitung starten")
        btn_run.clicked.connect(self.run_batch)
        btn_select_file = QPushButton("Datei auswählen")
        btn_select_file.clicked.connect(self.select_input_file)
        btn_select_folder = QPushButton("Ordner auswählen")
        btn_select_folder.clicked.connect(self.select_input_folder)
        btn_settings = QPushButton("Einstellungen")
        btn_settings.clicked.connect(self.open_settings_dialog)

        layout.addWidget(QLabel("Eingabe"), 0, 0)
        layout.addWidget(self.input_dir_edit, 0, 1)
        layout.addWidget(btn_select_file, 0, 2)
        layout.addWidget(btn_select_folder, 1, 2)

        layout.addWidget(QLabel("Stempelbreite"), 2, 0)
        layout.addWidget(self.stamp_width_spin, 2, 1)
        layout.addWidget(QLabel("Stempelhöhe"), 3, 0)
        layout.addWidget(self.stamp_height_spin, 3, 1)

        layout.addWidget(btn_settings, 3, 2)
        layout.addWidget(btn_run, 4, 0, 1, 3)

        return group

    def _build_log_group(self) -> QGroupBox:
        group = QGroupBox("3. Protokoll")
        layout = QVBoxLayout(group)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        layout.addWidget(self.log_edit)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        return group

    def _load_settings_into_ui(self) -> None:
        s = self.settings
        self.stamp_width_spin.setValue(s.stamp_width_mm)
        self.stamp_height_spin.setValue(s.stamp_height_mm)
        self.grid_step_spin.setValue(s.grid_step_mm)

    def _collect_settings(self) -> PlacementSettings:
        s = PlacementSettings(
            stamp_width_mm=self.stamp_width_spin.value(),
            stamp_height_mm=self.stamp_height_spin.value(),
            grid_step_mm=float(self.advanced_values["grid_step_mm"]),
            page_margin_mm=self.settings.page_margin_mm,
            render_dpi=int(self.advanced_values["render_dpi"]),
            whiteness_threshold=int(self.advanced_values["whiteness_threshold"]),
            max_occupancy_ratio=float(self.advanced_values["max_occupancy_ratio"]),
            dilation_px=int(self.advanced_values["dilation_px"]),
            allow_scale_down_to=float(self.advanced_values["allow_scale_down_to"]),
            process_mode=str(self.advanced_values["process_mode"]),
            preferred_anchor="bottom_right",
        )
        save_settings(s)
        return s

    def select_template_pdf(self) -> None:
        stamps_root = Path.cwd() / "stamps"
        if stamps_root.exists():
            dlg = StampSelectionDialog(self, stamps_root)
            if dlg.exec():
                selected = dlg.selected_pdf_path
                if selected is None:
                    return
                selected_page = dlg.selected_page_index
                prepared = self._prepare_selected_stamp_page(selected, selected_page)
                self.template_pdf_path = prepared
                self.template_path_edit.setText(f"{selected} (Seite {selected_page + 1})")
                self._invalidate_generated_stamp()
                self.load_form_fields()
                return

        file_name, _ = QFileDialog.getOpenFileName(self, "Formular-PDF wählen", "", "PDF (*.pdf)")
        if not file_name:
            return
        self.template_pdf_path = Path(file_name)
        self.template_path_edit.setText(file_name)
        self._invalidate_generated_stamp()
        self.load_form_fields()

    def _invalidate_generated_stamp(self) -> None:
        self.filled_stamp_pdf_path = None
        self.stamp_output_label.setText("Noch kein erzeugtes Stempel-PDF.")
        self.log("Stempelvorlage gewechselt: Bitte Stempel-PDF neu erzeugen.")

    def _prepare_selected_stamp_page(self, pdf_path: Path, page_index: int) -> Path:
        if page_index <= 0:
            return pdf_path
        out_path = self._temp_stamp_output_path().with_name("stamp_selected_page.pdf")
        src = fitz.open(pdf_path)
        dst = fitz.open()
        try:
            dst.insert_pdf(src, from_page=page_index, to_page=page_index)
            dst.save(out_path)
        finally:
            dst.close()
            src.close()
        return out_path

    def load_form_fields(self) -> None:
        if self.template_pdf_path is None:
            return

        self._clear_form_fields()

        try:
            fields = self.form_filler.list_fields(self.template_pdf_path)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Formular konnte nicht gelesen werden:\n{exc}")
            return

        if not fields:
            QMessageBox.warning(self, "Hinweis", "Keine Formularfelder gefunden.")
            return

        for field in fields:
            field_id = str(field.get("id") or field["name"])
            name = field["name"]
            label = field["label"]
            is_multiline = bool(field.get("is_multiline"))
            if is_multiline:
                edit = QTextEdit()
                edit.setFixedHeight(64)
                edit.setPlainText(str(field.get("value", "")))
                self.fields_form.addRow(f"{label} ({name})", edit)
                self.form_inputs[field_id] = edit
            else:
                edit = QLineEdit(str(field.get("value", "")))
                self.fields_form.addRow(f"{label} ({name})", edit)
                self.form_inputs[field_id] = edit

        self.log(f"{len(fields)} Formularfelder geladen.")

    def _clear_form_fields(self) -> None:
        while self.fields_form.rowCount() > 0:
            self.fields_form.removeRow(0)
        self.form_inputs.clear()

    def generate_stamp_pdf(self) -> None:
        if self.template_pdf_path is None:
            QMessageBox.warning(self, "Hinweis", "Bitte zuerst ein Formular-PDF auswählen.")
            return

        target_file = self._temp_stamp_output_path()

        field_values = {name: self._read_field_value(widget) for name, widget in self.form_inputs.items()}
        try:
            # Für die spätere Stempel-Annotation muss die Feld-Visualisierung erhalten bleiben.
            # Ein aggressives Flattening kann die sichtbaren Feldwerte entfernen.
            self.form_filler.fill_form(self.template_pdf_path, target_file, field_values, flatten=False)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Stempel-PDF konnte nicht erzeugt werden:\n{exc}")
            return

        self.filled_stamp_pdf_path = target_file
        self.stamp_output_label.setText(f"Stempel-PDF (temp): {target_file}")
        self._apply_stamp_size_defaults(self.filled_stamp_pdf_path)
        self.log(f"Stempel-PDF erzeugt: {target_file}")

    def select_input_file(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Datei auswählen", "", "PDF (*.pdf)")
        if file_name:
            path = Path(file_name)
            self.selected_input_file = path
            self.input_dir_path = path.parent
            self.input_dir_edit.setText(str(path))

    def select_input_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Ordner auswählen")
        if folder:
            self.selected_input_file = None
            self.input_dir_path = Path(folder)
            self.input_dir_edit.setText(folder)

    def run_batch(self) -> None:
        if self.filled_stamp_pdf_path is None:
            self.generate_stamp_pdf()
            if self.filled_stamp_pdf_path is None:
                QMessageBox.warning(self, "Hinweis", "Stempel-PDF konnte nicht automatisch erzeugt werden.")
                return
        if self.input_dir_path is None:
            QMessageBox.warning(self, "Hinweis", "Bitte Datei oder Ordner auswählen.")
            return

        settings = self._collect_settings()
        output_dir = self.input_dir_path / "gestempelt"
        ensure_dir(output_dir)
        config = BatchJobConfig(
            input_dir=self.input_dir_path,
            output_dir=output_dir,
            stamp_pdf=self.filled_stamp_pdf_path,
            settings=settings,
            output_suffix="",
        )
        processor = BatchProcessor(config)

        self.log("Stapelverarbeitung gestartet...")
        self.progress_bar.setValue(0)
        if self.selected_input_file is not None:
            file_result = processor.process_file(self.selected_input_file)
            results = [file_result]
            self._on_batch_progress(1, 1, self.selected_input_file, file_result)
        else:
            results = processor.process_all(progress_callback=self._on_batch_progress)
        if not results:
            self.log("Keine PDF-Dateien im Eingabeordner gefunden.")
            self.progress_bar.setValue(0)
            return

        success_count = 0
        reduced_count = 0
        no_position_count = 0
        manual_tasks: list[tuple[Path, Path, int, str]] = []
        for file_result in results:
            if file_result.success:
                success_count += 1
                self.log(f"OK: {file_result.input_file.name}")
                for pr in file_result.page_results:
                    if pr.status in {"no_position", "manual_required"}:
                        manual_tasks.append(
                            (
                                file_result.input_file,
                                file_result.output_file or file_result.input_file,
                                pr.page_index,
                                pr.note,
                            )
                        )
                    if pr.status == "no_position":
                        no_position_count += 1
                    if pr.scale < 0.999:
                        reduced_count += 1
                    self.log(f"  Seite {pr.page_index + 1}: {pr.status}, scale={pr.scale:.2f}")
            else:
                self.log(f"FEHLER: {file_result.input_file.name}: {file_result.error}")

        unstamped_dir = self.input_dir_path / "nicht_gestempelt"
        ensure_dir(unstamped_dir)
        unstamped_files: list[Path] = []
        for file_result in results:
            has_blocked = any(pr.status in {"no_position", "manual_required"} for pr in file_result.page_results)
            if has_blocked or getattr(file_result, "skipped", False):
                dst = unstamped_dir / file_result.input_file.name
                shutil.copy2(file_result.input_file, dst)
                unstamped_files.append(dst)
                self.log(f"WARNUNG: Nicht gestempelt kopiert -> {dst.name}")

        rotated_review_files: list[Path] = []
        for file_result in results:
            if not file_result.output_file:
                continue
            needs_review = any("mittig platziert" in (pr.note or "") for pr in file_result.page_results)
            if needs_review:
                rotated_review_files.append(file_result.output_file)
                self.log(f"Hinweis: Rotierte Seiten erkannt, bitte prüfen -> {file_result.output_file.name}")

        self.log(
            "Zusammenfassung: "
            f"{success_count}/{len(results)} Dateien erfolgreich, "
            f"{no_position_count} Seiten ohne freie Position, "
            f"{reduced_count} Platzierungen mit Skalierungsreduktion."
        )
        if unstamped_files:
            self.log(f"Warnung: {len(unstamped_files)} PDF(s) konnten nicht gestempelt werden und liegen in: {unstamped_dir}")
        self.log(f"Fertig. {success_count}/{len(results)} Dateien erfolgreich verarbeitet.")
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Fertig", f"{success_count}/{len(results)} Dateien erfolgreich verarbeitet.")

        if unstamped_files:
            self._open_unstamped_files(unstamped_files)
        if rotated_review_files:
            self._open_unstamped_files(rotated_review_files)

    def _open_unstamped_files(self, files: list[Path]) -> None:
        for file_path in files:
            try:
                if os.name == "nt":
                    os.startfile(str(file_path))  # type: ignore[attr-defined]
                elif os.name == "posix":
                    subprocess.Popen(["xdg-open", str(file_path)])
            except Exception as exc:
                self.log(f"Hinweis: Datei konnte nicht automatisch geöffnet werden ({file_path.name}): {exc}")

    def log(self, message: str) -> None:
        self.log_edit.append(message)

    def _read_field_value(self, widget: QWidget) -> str:
        if isinstance(widget, QTextEdit):
            return widget.toPlainText()
        if isinstance(widget, QLineEdit):
            return widget.text()
        return ""

    def _apply_stamp_size_defaults(self, stamp_pdf: Path) -> None:
        try:
            width_pt, height_pt = get_stamp_page_size(stamp_pdf)
        except Exception as exc:
            self.log(f"Hinweis: Stempelgröße konnte nicht aus Datei gelesen werden ({exc}).")
            return

        width_mm = pt_to_mm(width_pt)
        height_mm = pt_to_mm(height_pt)
        self.stamp_width_spin.setValue(width_mm)
        self.stamp_height_spin.setValue(height_mm)
        self.log(f"Standardgröße aus Stempel-PDF übernommen: {width_mm:.1f} x {height_mm:.1f} mm")

    def _label_with_info(self, text: str, info: str) -> QWidget:
        wrap = QWidget()
        row = QHBoxLayout(wrap)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)
        row.addWidget(QLabel(text))
        button = QToolButton()
        button.setText("i")
        button.setToolTip("Info")
        button.clicked.connect(lambda: QMessageBox.information(self, text, info))
        row.addWidget(button)
        row.addStretch(1)
        return wrap

    def _on_batch_progress(self, done: int, total: int, pdf_path: Path, file_result: object) -> None:
        percentage = int((done / total) * 100) if total else 0
        self.progress_bar.setValue(percentage)
        status = "OK" if getattr(file_result, "success", False) else "FEHLER"
        self.log(f"Fortschritt: {done}/{total} ({percentage} %) - {status} - {pdf_path.name}")
        QApplication.processEvents()

    def _temp_stamp_output_path(self) -> Path:
        temp_root = Path(tempfile.gettempdir()) / "pdf_stamper"
        ensure_dir(temp_root)
        return temp_root / "stamp_filled.pdf"

    def open_settings_dialog(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Einstellungen")
        layout = QFormLayout(dlg)

        grid_step = QDoubleSpinBox()
        grid_step.setRange(1.0, 100.0)
        grid_step.setSuffix(" mm")
        grid_step.setValue(float(self.advanced_values["grid_step_mm"]))
        dpi = QSpinBox()
        dpi.setRange(72, 600)
        dpi.setValue(int(self.advanced_values["render_dpi"]))
        threshold = QSpinBox()
        threshold.setRange(0, 255)
        threshold.setValue(int(self.advanced_values["whiteness_threshold"]))
        max_occ = QDoubleSpinBox()
        max_occ.setRange(0.0, 1.0)
        max_occ.setSingleStep(0.01)
        max_occ.setDecimals(3)
        max_occ.setValue(float(self.advanced_values["max_occupancy_ratio"]))
        dilation = QSpinBox()
        dilation.setRange(0, 20)
        dilation.setValue(int(self.advanced_values["dilation_px"]))
        scale = QDoubleSpinBox()
        scale.setRange(0.1, 1.0)
        scale.setSingleStep(0.05)
        scale.setDecimals(2)
        scale.setValue(float(self.advanced_values["allow_scale_down_to"]))
        mode = QComboBox()
        mode.addItems(["all", "first", "last"])
        mode.setCurrentText(str(self.advanced_values["process_mode"]))

        layout.addRow("Raster-Schritt (kleiner = genauer, langsamer)", grid_step)
        layout.addRow("Render-DPI (höher = präziser, langsamer)", dpi)
        layout.addRow("Weiß-Schwelle (höher = mehr als frei erkannt)", threshold)
        layout.addRow("Max. Belegungsquote (kleiner = strenger)", max_occ)
        layout.addRow("Dilation (Sicherheitsabstand in Pixeln)", dilation)
        layout.addRow("Min. Skalierung (bei Platzmangel)", scale)
        layout.addRow("Seitenmodus", mode)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addRow(buttons)

        if dlg.exec():
            self.advanced_values = {
                "grid_step_mm": grid_step.value(),
                "render_dpi": dpi.value(),
                "whiteness_threshold": threshold.value(),
                "max_occupancy_ratio": max_occ.value(),
                "dilation_px": dilation.value(),
                "allow_scale_down_to": scale.value(),
                "process_mode": mode.currentText(),
            }

    def _open_manual_review_popup(self, source_pdf: Path, target_pdf: Path, page_index: int, note: str) -> None:
        render_dpi = 220
        render_zoom = render_dpi / 72.0
        try:
            doc = fitz.open(source_pdf)
            try:
                page = doc[page_index]
                pix = page.get_pixmap(dpi=render_dpi, alpha=False)
            finally:
                doc.close()
        except Exception as exc:
            self.log(f"Hinweis: Vorschau konnte nicht geöffnet werden: {exc}")
            return

        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        qpix = QPixmap.fromImage(img)
        stamp_pixmap = self._load_stamp_preview_pixmap()

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Manuelle Prüfung: {source_pdf.name} - Seite {page_index + 1}")
        dlg.resize(1200, 900)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(note or "Manuelle Prüfung erforderlich."))

        controls = QHBoxLayout()
        btn_minus = QPushButton("Zoom -")
        btn_plus = QPushButton("Zoom +")
        btn_place = QPushButton("Stempel manuell platzieren")
        btn_save = QPushButton("Manuelle Position speichern")
        btn_editor = QPushButton("Im PDF-Editor öffnen")
        controls.addWidget(btn_minus)
        controls.addWidget(btn_plus)
        controls.addWidget(btn_place)
        controls.addWidget(btn_save)
        controls.addWidget(btn_editor)
        controls.addStretch(1)
        layout.addLayout(controls)

        scene = QGraphicsScene()
        item = QGraphicsPixmapItem(qpix)
        scene.addItem(item)
        view = QGraphicsView(scene)
        view.setDragMode(QGraphicsView.ScrollHandDrag)
        view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        view.setRenderHint(QPainter.SmoothPixmapTransform, True)
        layout.addWidget(view, stretch=1)
        view.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)

        stamp_item: QGraphicsPixmapItem | None = None
        place_mode = {"active": False}

        def zoom(factor: float) -> None:
            view.scale(factor, factor)

        btn_plus.clicked.connect(lambda: zoom(1.25))
        btn_minus.clicked.connect(lambda: zoom(0.8))
        btn_place.clicked.connect(lambda: place_mode.update({"active": True}))

        original_mouse_press = view.mousePressEvent

        def mouse_press(event) -> None:  # type: ignore[no-untyped-def]
            nonlocal stamp_item
            if place_mode["active"] and stamp_pixmap is not None and event.button() == Qt.LeftButton:
                scene_pos = view.mapToScene(event.pos())
                if stamp_item is None:
                    stamp_item = QGraphicsPixmapItem(stamp_pixmap)
                    stamp_item.setFlag(QGraphicsPixmapItem.ItemIsMovable, True)
                    stamp_item.setFlag(QGraphicsPixmapItem.ItemIsSelectable, True)
                    stamp_item.setOpacity(0.85)
                    scene.addItem(stamp_item)
                stamp_item.setPos(
                    scene_pos.x() - stamp_item.boundingRect().width() * 0.5,
                    scene_pos.y() - stamp_item.boundingRect().height() * 0.5,
                )
                place_mode["active"] = False
                return
            original_mouse_press(event)

        view.mousePressEvent = mouse_press  # type: ignore[assignment]

        def save_manual_position() -> None:
            if stamp_item is None or self.filled_stamp_pdf_path is None:
                return
            stamp_rect_px = stamp_item.sceneBoundingRect()
            rect_rot = fitz.Rect(
                stamp_rect_px.left() / render_zoom,
                stamp_rect_px.top() / render_zoom,
                stamp_rect_px.right() / render_zoom,
                stamp_rect_px.bottom() / render_zoom,
            )
            try:
                out_doc = fitz.open(target_pdf)
                stamp_doc = fitz.open(self.filled_stamp_pdf_path)
                try:
                    target_page = out_doc[page_index]
                    derot = target_page.derotation_matrix
                    p0 = fitz.Point(rect_rot.x0, rect_rot.y0) * derot
                    p1 = fitz.Point(rect_rot.x1, rect_rot.y1) * derot
                    rect_pt = fitz.Rect(min(p0.x, p1.x), min(p0.y, p1.y), max(p0.x, p1.x), max(p0.y, p1.y))
                    stamp_src = stamp_doc[0].get_pixmap(dpi=600, alpha=True)
                    annot = target_page.add_stamp_annot(rect_pt, stamp=stamp_src)
                    annot.set_rotation(int(target_page.rotation) % 360)
                    annot.update()
                    out_doc.save(target_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
                finally:
                    stamp_doc.close()
                    out_doc.close()
                self.log(f"Manuell gespeichert: {target_pdf.name}, Seite {page_index + 1}")
                QMessageBox.information(dlg, "Gespeichert", "Manuelle Stempelposition wurde gespeichert.")
            except Exception as exc:
                QMessageBox.critical(dlg, "Fehler", f"Manuelle Position konnte nicht gespeichert werden:\n{exc}")

        btn_save.clicked.connect(save_manual_position)
        btn_editor.clicked.connect(lambda: self._open_pdf_editor(source_pdf, target_pdf, page_index))
        dlg.exec()

    def _load_stamp_preview_pixmap(self) -> QPixmap | None:
        if self.filled_stamp_pdf_path is None:
            return None
        try:
            stamp_doc = fitz.open(self.filled_stamp_pdf_path)
            try:
                pix = stamp_doc[0].get_pixmap(dpi=150, alpha=True)
            finally:
                stamp_doc.close()
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888).copy()
            return QPixmap.fromImage(img)
        except Exception as exc:
            self.log(f"Hinweis: Stempelvorschau konnte nicht geladen werden: {exc}")
            return None

    def _open_pdf_editor(
        self,
        source_pdf: Path,
        target_pdf: Path,
        page_index: int,
        note: str,
        queue_index: int,
        queue_total: int,
    ) -> None:
        editor = ManualPdfEditorDialog(
            self,
            source_pdf,
            target_pdf,
            self.filled_stamp_pdf_path,
            page_index,
            note,
            queue_index,
            queue_total,
        )
        editor.exec()


class StampSelectionDialog(QDialog):
    def __init__(self, parent: QWidget, stamps_root: Path) -> None:
        super().__init__(parent)
        self.stamps_root = stamps_root
        self.selected_pdf_path: Path | None = None
        self.selected_page_index: int = 0
        self.setWindowTitle("Stempelpakete auswählen")
        self.resize(1200, 750)

        root = QVBoxLayout(self)
        root.addWidget(QLabel(f"Stempelverzeichnis: {stamps_root}"))
        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, stretch=1)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Stempelpakete / Dateien"])
        splitter.addWidget(self.tree)

        self.preview = QListWidget()
        self.preview.setViewMode(QListWidget.IconMode)
        self.preview.setResizeMode(QListWidget.Adjust)
        self.preview.setIconSize(QPixmap(220, 220).size())
        self.preview.setSpacing(12)
        splitter.addWidget(self.preview)
        splitter.setSizes([350, 850])

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        root.addWidget(buttons)
        buttons.accepted.connect(self._accept_selection)
        buttons.rejected.connect(self.reject)

        self.tree.itemSelectionChanged.connect(self._on_tree_selection_changed)
        self.preview.itemDoubleClicked.connect(lambda _: self._accept_selection())
        self._build_tree()

    def _build_tree(self) -> None:
        self.tree.clear()
        root_item = QTreeWidgetItem([self.stamps_root.name])
        root_item.setData(0, Qt.UserRole, str(self.stamps_root))
        self.tree.addTopLevelItem(root_item)
        self._add_dir_children(root_item, self.stamps_root)
        root_item.setExpanded(True)

    def _add_dir_children(self, parent_item: QTreeWidgetItem, directory: Path) -> None:
        for sub in sorted([p for p in directory.iterdir() if p.is_dir()], key=lambda x: x.name.lower()):
            d_item = QTreeWidgetItem([sub.name])
            d_item.setData(0, Qt.UserRole, str(sub))
            parent_item.addChild(d_item)
            self._add_dir_children(d_item, sub)

        for pdf in sorted(directory.glob("*.pdf"), key=lambda x: x.name.lower()):
            p_item = QTreeWidgetItem([pdf.name])
            p_item.setData(0, Qt.UserRole, str(pdf))
            parent_item.addChild(p_item)

    def _on_tree_selection_changed(self) -> None:
        selected = self.tree.selectedItems()
        if not selected:
            return
        path = Path(selected[0].data(0, Qt.UserRole))
        if path.is_file() and path.suffix.lower() == ".pdf":
            self._load_pdf_previews(path)
        else:
            self.preview.clear()

    def _load_pdf_previews(self, pdf_path: Path) -> None:
        self.preview.clear()
        try:
            doc = fitz.open(pdf_path)
            try:
                for page_idx in range(len(doc)):
                    pix = doc[page_idx].get_pixmap(dpi=96, alpha=False)
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
                    pm = QPixmap.fromImage(img).scaled(220, 220, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    item = QListWidgetItem(f"Seite {page_idx + 1}")
                    item.setIcon(QIcon(pm))
                    item.setData(Qt.UserRole, (str(pdf_path), page_idx))
                    self.preview.addItem(item)
            finally:
                doc.close()
        except Exception:
            pass

    def _accept_selection(self) -> None:
        current = self.preview.currentItem()
        if current is None:
            QMessageBox.warning(self, "Hinweis", "Bitte eine Stempelseite auswählen.")
            return
        raw = current.data(Qt.UserRole)
        if not raw:
            return
        pdf_path, page_idx = raw
        self.selected_pdf_path = Path(pdf_path)
        self.selected_page_index = int(page_idx)
        self.accept()


class ManualPdfEditorDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        source_pdf: Path,
        target_pdf: Path,
        stamp_pdf: Path | None,
        page_index: int,
        note: str,
        queue_index: int,
        queue_total: int,
    ) -> None:
        super().__init__(parent)
        self.source_pdf = source_pdf
        self.target_pdf = target_pdf
        self.stamp_pdf = stamp_pdf
        self.page_index = page_index
        self.render_dpi = 150
        self.page_pixmap: QPixmap | None = None
        self.stamp_preview: QPixmap | None = None
        self.annotation_items: list[dict] = []
        self.selected_index: int | None = None

        self.setWindowTitle(f"PDF-Editor: {source_pdf.name}")
        self.resize(1500, 1000)

        root = QVBoxLayout(self)
        self.lbl_status = QLabel(
            f"Seite {page_index + 1} | Manuell {queue_index}/{queue_total} | {note or 'Manuelle Platzierung erforderlich.'}"
        )
        root.addWidget(self.lbl_status)

        props = QHBoxLayout()
        self.x_spin = QDoubleSpinBox(); self.x_spin.setRange(-2000, 2000); self.x_spin.setSuffix(' pt')
        self.y_spin = QDoubleSpinBox(); self.y_spin.setRange(-2000, 2000); self.y_spin.setSuffix(' pt')
        self.w_spin = QDoubleSpinBox(); self.w_spin.setRange(1, 5000); self.w_spin.setSuffix(' pt')
        self.h_spin = QDoubleSpinBox(); self.h_spin.setRange(1, 5000); self.h_spin.setSuffix(' pt')
        self.rot_spin = QSpinBox(); self.rot_spin.setRange(0, 359); self.rot_spin.setSuffix('°')
        self.btn_apply = QPushButton('Übernehmen')
        self.btn_add = QPushButton('Neue Stempel-Annotation')
        self.btn_delete = QPushButton('Löschen')
        self.btn_save = QPushButton('PDF speichern')
        for t,w in [('X',self.x_spin),('Y',self.y_spin),('Breite',self.w_spin),('Höhe',self.h_spin),('Rotation',self.rot_spin)]:
            props.addWidget(QLabel(t)); props.addWidget(w)
        props.addWidget(self.btn_apply); props.addWidget(self.btn_add); props.addWidget(self.btn_delete); props.addWidget(self.btn_save)
        root.addLayout(props)

        body = QHBoxLayout()
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.SmoothPixmapTransform, True)
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)
        body.addWidget(self.view, stretch=1)

        side = QVBoxLayout()
        side.addWidget(QLabel('Stempel-Annotationen'))
        from PySide6.QtWidgets import QListWidget
        self.annot_list = QListWidget()
        side.addWidget(self.annot_list, stretch=1)
        body.addLayout(side)
        root.addLayout(body, stretch=1)

        self.btn_add.clicked.connect(self._add_annotation)
        self.btn_delete.clicked.connect(self._delete_selected)
        self.btn_apply.clicked.connect(self._apply_properties)
        self.btn_save.clicked.connect(self._save_pdf)
        self.annot_list.currentRowChanged.connect(self._select_index)

        self._load_page_and_annotations()

    def _load_stamp_preview(self) -> None:
        if self.stamp_pdf is None:
            self.stamp_preview = None
            return
        doc = fitz.open(self.stamp_pdf)
        try:
            pix = doc[0].get_pixmap(dpi=120, alpha=True)
        finally:
            doc.close()
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888).copy()
        self.stamp_preview = QPixmap.fromImage(img)

    def _load_page_and_annotations(self) -> None:
        self._load_stamp_preview()
        doc = fitz.open(self.target_pdf)
        try:
            page = doc[self.page_index]
            pix = page.get_pixmap(dpi=self.render_dpi, alpha=False)
            annots = []
            a = page.first_annot
            while a is not None:
                if a.type[0] == fitz.PDF_ANNOT_STAMP:
                    annots.append((fitz.Rect(a.rect), int(a.rotation or 0)))
                a = a.next
        finally:
            doc.close()

        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        self.page_pixmap = QPixmap.fromImage(img)
        self.scene.clear()
        self.annotation_items.clear()
        self.annot_list.clear()
        page_item = QGraphicsPixmapItem(self.page_pixmap)
        self.scene.addItem(page_item)

        for idx, (rect, rot) in enumerate(annots):
            self._create_item_from_pdf_rect(rect, rot, f'Annot {idx+1}')
        self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)

    def _scale(self) -> float:
        return self.render_dpi / 72.0

    def _create_item_from_pdf_rect(self, rect: fitz.Rect, rot: int, name: str) -> None:
        if self.stamp_preview is None:
            return
        s = self._scale()
        x, y, w, h = rect.x0 * s, rect.y0 * s, rect.width * s, rect.height * s
        pix = self.stamp_preview.scaled(max(10, int(w)), max(10, int(h)), Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
        item = QGraphicsPixmapItem(pix)
        item.setPos(x, y)
        item.setRotation(rot)
        item.setFlag(QGraphicsPixmapItem.ItemIsMovable, True)
        item.setFlag(QGraphicsPixmapItem.ItemIsSelectable, True)
        item.setOpacity(0.85)
        self.scene.addItem(item)
        self.annotation_items.append({'item': item, 'rotation': rot})
        self.annot_list.addItem(name)

    def _add_annotation(self) -> None:
        if self.stamp_preview is None:
            QMessageBox.warning(self, 'Hinweis', 'Kein Stempel-PDF verfügbar.')
            return
        w, h = self.stamp_preview.width(), self.stamp_preview.height()
        item = QGraphicsPixmapItem(self.stamp_preview)
        item.setPos(50, 50)
        item.setFlag(QGraphicsPixmapItem.ItemIsMovable, True)
        item.setFlag(QGraphicsPixmapItem.ItemIsSelectable, True)
        item.setOpacity(0.85)
        self.scene.addItem(item)
        self.annotation_items.append({'item': item, 'rotation': 0})
        self.annot_list.addItem(f'Annot {len(self.annotation_items)}')
        self.annot_list.setCurrentRow(len(self.annotation_items)-1)

    def _select_index(self, idx: int) -> None:
        self.selected_index = idx if idx >= 0 and idx < len(self.annotation_items) else None
        if self.selected_index is None:
            return
        rec = self.annotation_items[self.selected_index]
        item = rec['item']
        s = self._scale()
        self.x_spin.setValue(item.pos().x()/s)
        self.y_spin.setValue(item.pos().y()/s)
        self.w_spin.setValue(item.pixmap().width()/s)
        self.h_spin.setValue(item.pixmap().height()/s)
        self.rot_spin.setValue(int(rec['rotation'])%360)

    def _apply_properties(self) -> None:
        if self.selected_index is None:
            return
        rec = self.annotation_items[self.selected_index]
        item = rec['item']
        s = self._scale()
        w_px = max(10, int(self.w_spin.value()*s))
        h_px = max(10, int(self.h_spin.value()*s))
        if self.stamp_preview is not None:
            item.setPixmap(self.stamp_preview.scaled(w_px, h_px, Qt.IgnoreAspectRatio, Qt.SmoothTransformation))
        item.setPos(self.x_spin.value()*s, self.y_spin.value()*s)
        rec['rotation'] = int(self.rot_spin.value())
        item.setRotation(rec['rotation'])

    def _delete_selected(self) -> None:
        if self.selected_index is None:
            return
        rec = self.annotation_items.pop(self.selected_index)
        self.scene.removeItem(rec['item'])
        self.annot_list.takeItem(self.selected_index)
        self.selected_index = None

    def _save_pdf(self) -> None:
        if self.stamp_pdf is None:
            QMessageBox.warning(self, 'Hinweis', 'Kein Stempel-PDF verfügbar.')
            return
        out_doc = fitz.open(self.target_pdf)
        stamp_doc = fitz.open(self.stamp_pdf)
        try:
            page = out_doc[self.page_index]
            a = page.first_annot
            to_delete = []
            while a is not None:
                if a.type[0] == fitz.PDF_ANNOT_STAMP:
                    to_delete.append(a)
                a = a.next
            for annot in to_delete:
                page.delete_annot(annot)

            s = self._scale()
            for rec in self.annotation_items:
                item = rec['item']
                rect = fitz.Rect(item.pos().x()/s, item.pos().y()/s, (item.pos().x()+item.pixmap().width())/s, (item.pos().y()+item.pixmap().height())/s)
                pix = stamp_doc[0].get_pixmap(dpi=600, alpha=True)
                annot = page.add_stamp_annot(rect, stamp=pix)
                annot.set_rotation(int(rec['rotation'])%360)
                annot.update()
            out_doc.save(self.target_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        finally:
            stamp_doc.close(); out_doc.close()
        QMessageBox.information(self, 'Gespeichert', 'Stempel-Annotationen gespeichert.')
        self.accept()

