from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Dict

import fitz
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QImage, QPainter, QPixmap
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
        file_name, _ = QFileDialog.getOpenFileName(self, "Formular-PDF wählen", "", "PDF (*.pdf)")
        if not file_name:
            return
        self.template_pdf_path = Path(file_name)
        self.template_path_edit.setText(file_name)
        self.load_form_fields()

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
        for file_result in results:
            if file_result.success:
                success_count += 1
                self.log(f"OK: {file_result.input_file.name}")
                for pr in file_result.page_results:
                    if pr.status in {"no_position", "manual_required"}:
                        self._open_pdf_editor(
                            file_result.input_file,
                            file_result.output_file or file_result.input_file,
                            pr.page_index,
                        )
                    if pr.status == "no_position":
                        no_position_count += 1
                    if pr.scale < 0.999:
                        reduced_count += 1
                    self.log(f"  Seite {pr.page_index + 1}: {pr.status}, scale={pr.scale:.2f}")
            else:
                self.log(f"FEHLER: {file_result.input_file.name}: {file_result.error}")

        self.log(
            "Zusammenfassung: "
            f"{success_count}/{len(results)} Dateien erfolgreich, "
            f"{no_position_count} Seiten ohne freie Position, "
            f"{reduced_count} Platzierungen mit Skalierungsreduktion."
        )
        self.log(f"Fertig. {success_count}/{len(results)} Dateien erfolgreich verarbeitet.")
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Fertig", f"{success_count}/{len(results)} Dateien erfolgreich verarbeitet.")

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
                    stamp_src = stamp_doc[0].get_pixmap(dpi=300, alpha=True)
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

    def _open_pdf_editor(self, source_pdf: Path, target_pdf: Path, page_index: int) -> None:
        editor = ManualPdfEditorDialog(self, source_pdf, target_pdf, self.filled_stamp_pdf_path, page_index)
        editor.exec()


class ManualPdfEditorDialog(QDialog):
    def __init__(self, parent: QWidget, source_pdf: Path, target_pdf: Path, stamp_pdf: Path | None, page_index: int) -> None:
        super().__init__(parent)
        self.source_pdf = source_pdf
        self.target_pdf = target_pdf
        self.stamp_pdf = stamp_pdf
        self.page_index = page_index
        self.stamp_label: QLabel | None = None
        self.stamp_size_px: tuple[int, int] = (0, 0)
        self.drag_active = False
        self.drag_offset = QPointF(0.0, 0.0)

        self.setWindowTitle(f"PDF-Editor: {source_pdf.name}")
        self.resize(1400, 950)

        root = QVBoxLayout(self)
        ctrl = QHBoxLayout()
        self.btn_zoom_out = QPushButton("Zoom -")
        self.btn_zoom_in = QPushButton("Zoom +")
        self.btn_insert = QPushButton("Stempel platzieren")
        self.btn_place = QPushButton("Position speichern")
        self.info = QLabel("Mit Strg+Mausrad zoomen, normal scrollen.")
        ctrl.addWidget(self.btn_zoom_out)
        ctrl.addWidget(self.btn_zoom_in)
        ctrl.addWidget(self.btn_insert)
        ctrl.addWidget(self.btn_place)
        ctrl.addWidget(self.info)
        ctrl.addStretch(1)
        root.addLayout(ctrl)

        self.pdf_doc = QPdfDocument(self)
        self.pdf_doc.load(str(source_pdf))
        self.pdf_view = QPdfView(self)
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_view.setPageMode(QPdfView.PageMode.SinglePage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.pdf_view.pageNavigator().jump(page_index, QPointF(), 0)
        self.pdf_view.viewport().installEventFilter(self)
        root.addWidget(self.pdf_view, stretch=1)

        self.btn_zoom_in.clicked.connect(lambda: self.pdf_view.setZoomFactor(self.pdf_view.zoomFactor() * 1.2))
        self.btn_zoom_out.clicked.connect(lambda: self.pdf_view.setZoomFactor(self.pdf_view.zoomFactor() / 1.2))
        self.btn_insert.clicked.connect(self._insert_stamp_center)
        self.btn_place.clicked.connect(self._save_stamp_at_click)

    def eventFilter(self, obj: object, event: object) -> bool:
        if obj is self.pdf_view.viewport():
            etype = getattr(event, "type", lambda: None)()
            if etype == 31:  # Wheel
                mods = event.modifiers()
                if mods & Qt.ControlModifier:
                    delta = event.angleDelta().y()
                    factor = 1.15 if delta > 0 else 0.87
                    self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
                    self.pdf_view.setZoomFactor(max(0.1, min(8.0, self.pdf_view.zoomFactor() * factor)))
                    return True
            if etype == 2 and self.stamp_label is not None:  # MouseButtonPress
                local = event.position()
                if self.stamp_label.geometry().contains(int(local.x()), int(local.y())):
                    self.drag_active = True
                    self.drag_offset = QPointF(local.x() - self.stamp_label.x(), local.y() - self.stamp_label.y())
                    return True
            if etype == 5 and self.drag_active and self.stamp_label is not None:  # MouseMove
                local = event.position()
                nx = int(local.x() - self.drag_offset.x())
                ny = int(local.y() - self.drag_offset.y())
                vp = self.pdf_view.viewport().rect()
                nx = max(0, min(vp.width() - self.stamp_label.width(), nx))
                ny = max(0, min(vp.height() - self.stamp_label.height(), ny))
                self.stamp_label.move(nx, ny)
                return True
            if etype == 3 and self.drag_active:  # MouseButtonRelease
                self.drag_active = False
                return True
        return super().eventFilter(obj, event)

    def _insert_stamp_center(self) -> None:
        if self.stamp_pdf is None:
            return
        if self.stamp_label is not None:
            return
        try:
            stamp_doc = fitz.open(self.stamp_pdf)
            try:
                stamp_w_pt, stamp_h_pt = stamp_doc[0].rect.width, stamp_doc[0].rect.height
                pix = stamp_doc[0].get_pixmap(dpi=150, alpha=True)
            finally:
                stamp_doc.close()
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Stempel konnte nicht geladen werden:\n{exc}")
            return
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888).copy()
        qpix = QPixmap.fromImage(img)
        page_w, page_h = self._current_page_size_pt()
        if page_w <= 0 or page_h <= 0:
            return
        page_rect = self._page_display_rect(page_w, page_h)
        scale = page_rect.width() / page_w
        stamp_w_px = max(12, int(stamp_w_pt * scale))
        stamp_h_px = max(12, int(stamp_h_pt * scale))
        scaled = qpix.scaled(stamp_w_px, stamp_h_px, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        self.stamp_label = QLabel(self.pdf_view.viewport())
        self.stamp_label.setPixmap(scaled)
        self.stamp_label.setWindowOpacity(0.85)
        self.stamp_label.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.stamp_label.resize(scaled.size())
        self.stamp_size_px = (scaled.width(), scaled.height())
        cx = int(page_rect.center().x() - scaled.width() * 0.5)
        cy = int(page_rect.center().y() - scaled.height() * 0.5)
        self.stamp_label.move(cx, cy)
        self.stamp_label.show()

    def _save_stamp_at_click(self) -> None:
        if self.stamp_pdf is None or self.stamp_label is None:
            QMessageBox.warning(self, "Hinweis", "Bitte zuerst 'Stempel platzieren' klicken.")
            return
        try:
            out_doc = fitz.open(self.target_pdf)
            stamp_doc = fitz.open(self.stamp_pdf)
            try:
                page = out_doc[self.page_index]
                sw, sh = stamp_doc[0].rect.width, stamp_doc[0].rect.height
                page_rect = self._page_display_rect(page.rect.width, page.rect.height)
                center_x = self.stamp_label.x() + self.stamp_label.width() * 0.5
                center_y = self.stamp_label.y() + self.stamp_label.height() * 0.5
                rx = (center_x - page_rect.left()) / max(1.0, page_rect.width())
                ry = (center_y - page_rect.top()) / max(1.0, page_rect.height())
                rx = max(0.0, min(1.0, rx))
                ry = max(0.0, min(1.0, ry))
                w, h = page.rect.width, page.rect.height
                x0 = max(0.0, min(w - sw, rx * w - sw * 0.5))
                y0 = max(0.0, min(h - sh, ry * h - sh * 0.5))
                rect = fitz.Rect(x0, y0, x0 + sw, y0 + sh)
                pix = stamp_doc[0].get_pixmap(dpi=300, alpha=True)
                annot = page.add_stamp_annot(rect, stamp=pix)
                annot.set_rotation(int(page.rotation) % 360)
                annot.update()
                out_doc.save(self.target_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
            finally:
                stamp_doc.close()
                out_doc.close()
            QMessageBox.information(self, "Gespeichert", "Stempel wurde im Ziel-PDF gespeichert.")
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Speichern fehlgeschlagen:\n{exc}")

    def _current_page_size_pt(self) -> tuple[float, float]:
        try:
            doc = fitz.open(self.source_pdf)
            try:
                rect = doc[self.page_index].rect
                return rect.width, rect.height
            finally:
                doc.close()
        except Exception:
            return (0.0, 0.0)

    def _page_display_rect(self, page_w: float, page_h: float) -> QRectF:
        vp = self.pdf_view.viewport().rect()
        if page_w <= 0 or page_h <= 0:
            return QRectF(vp)
        if self.pdf_view.zoomMode() == QPdfView.ZoomMode.FitToWidth:
            disp_w = float(vp.width())
            disp_h = disp_w * (page_h / page_w)
        else:
            zoom = float(self.pdf_view.zoomFactor())
            disp_w = page_w * zoom
            disp_h = page_h * zoom

        hbar = self.pdf_view.horizontalScrollBar()
        vbar = self.pdf_view.verticalScrollBar()
        x = -float(hbar.value())
        y = -float(vbar.value())

        if disp_w < vp.width():
            x = (vp.width() - disp_w) * 0.5
        if disp_h < vp.height():
            y = (vp.height() - disp_h) * 0.5
        return QRectF(x, y, disp_w, disp_h)
