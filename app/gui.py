from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Dict

from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
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
        self.output_dir_path: Path | None = None

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
        group = QGroupBox("1. Stempelvorlage")
        layout = QVBoxLayout(group)

        row = QHBoxLayout()
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        btn_select_template = QPushButton("Formular-PDF wählen")
        btn_select_template.clicked.connect(self.select_template_pdf)
        row.addWidget(self.template_path_edit, stretch=1)
        row.addWidget(btn_select_template)
        layout.addLayout(row)

        self.fields_area = QScrollArea()
        self.fields_area.setWidgetResizable(True)
        self.fields_container = QWidget()
        self.fields_form = QFormLayout(self.fields_container)
        self.fields_area.setWidget(self.fields_container)
        layout.addWidget(self.fields_area)

        btn_fill = QPushButton("Stempel-PDF erzeugen")
        btn_fill.clicked.connect(self.generate_stamp_pdf)
        layout.addWidget(btn_fill)

        self.stamp_output_label = QLabel("Noch kein erzeugtes Stempel-PDF.")
        layout.addWidget(self.stamp_output_label)

        return group

    def _build_processing_group(self) -> QGroupBox:
        group = QGroupBox("2. Stapelverarbeitung")
        layout = QGridLayout(group)

        self.input_dir_edit = QLineEdit()
        self.input_dir_edit.setReadOnly(True)
        btn_input = QPushButton("Eingabeordner")
        btn_input.clicked.connect(self.select_input_dir)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        btn_output = QPushButton("Ausgabeordner")
        btn_output.clicked.connect(self.select_output_dir)

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

        self.margin_spin = QDoubleSpinBox()
        self.margin_spin.setRange(0.0, 100.0)
        self.margin_spin.setSuffix(" mm")
        self.margin_spin.setDecimals(1)

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

        btn_run = QPushButton("Stapelverarbeitung starten")
        btn_run.clicked.connect(self.run_batch)

        step_label = self._label_with_info(
            "Raster-Schritt",
            "Abstand zwischen geprüften Kandidat-Positionen. Kleiner = genauer, aber langsamer.",
        )
        dpi_label = self._label_with_info(
            "Render-DPI",
            "Auflösung für die Seitendarstellung zur Freiflächenanalyse. Höher = präziser, aber langsamer/speicherintensiver.",
        )
        scale_label = self._label_with_info(
            "Min. Skalierung",
            "Wenn kein Platz gefunden wird, wird der Stempel schrittweise bis zu diesem Faktor verkleinert.",
        )
        occ_label = self._label_with_info(
            "Max. Belegungsquote",
            "Maximal erlaubter Anteil belegter Pixel im Zielrechteck. Kleiner = strenger bei Überlappungen.",
        )
        dilation_label = self._label_with_info(
            "Dilation",
            "Erweitert erkannte belegte Bereiche um Sicherheitsabstand (in Pixeln). Höher = konservativer.",
        )

        layout.addWidget(QLabel("Eingabeordner"), 0, 0)
        layout.addWidget(self.input_dir_edit, 0, 1)
        layout.addWidget(btn_input, 0, 2)

        layout.addWidget(QLabel("Ausgabeordner"), 1, 0)
        layout.addWidget(self.output_dir_edit, 1, 1)
        layout.addWidget(btn_output, 1, 2)

        layout.addWidget(QLabel("Stempelbreite"), 2, 0)
        layout.addWidget(self.stamp_width_spin, 2, 1)
        layout.addWidget(QLabel("Stempelhöhe"), 2, 2)
        layout.addWidget(self.stamp_height_spin, 2, 3)

        layout.addWidget(step_label, 3, 0)
        layout.addWidget(self.grid_step_spin, 3, 1)
        layout.addWidget(QLabel("Seitenrand"), 3, 2)
        layout.addWidget(self.margin_spin, 3, 3)

        layout.addWidget(dpi_label, 4, 0)
        layout.addWidget(self.dpi_spin, 4, 1)
        layout.addWidget(QLabel("Weiß-Schwelle"), 4, 2)
        layout.addWidget(self.threshold_spin, 4, 3)

        layout.addWidget(occ_label, 5, 0)
        layout.addWidget(self.max_occ_spin, 5, 1)
        layout.addWidget(dilation_label, 5, 2)
        layout.addWidget(self.dilation_spin, 5, 3)

        layout.addWidget(scale_label, 6, 0)
        layout.addWidget(self.scale_down_spin, 6, 1)
        layout.addWidget(QLabel("Seitenmodus"), 6, 2)
        layout.addWidget(self.mode_combo, 6, 3)

        layout.addWidget(btn_run, 7, 0, 1, 4)

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
        self.margin_spin.setValue(s.page_margin_mm)
        self.dpi_spin.setValue(s.render_dpi)
        self.threshold_spin.setValue(s.whiteness_threshold)
        self.max_occ_spin.setValue(s.max_occupancy_ratio)
        self.dilation_spin.setValue(s.dilation_px)
        self.scale_down_spin.setValue(s.allow_scale_down_to)
        self.mode_combo.setCurrentText(s.process_mode)

    def _collect_settings(self) -> PlacementSettings:
        s = PlacementSettings(
            stamp_width_mm=self.stamp_width_spin.value(),
            stamp_height_mm=self.stamp_height_spin.value(),
            grid_step_mm=self.grid_step_spin.value(),
            page_margin_mm=self.margin_spin.value(),
            render_dpi=self.dpi_spin.value(),
            whiteness_threshold=self.threshold_spin.value(),
            max_occupancy_ratio=self.max_occ_spin.value(),
            dilation_px=self.dilation_spin.value(),
            allow_scale_down_to=self.scale_down_spin.value(),
            process_mode=self.mode_combo.currentText(),
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
            self.form_filler.fill_form(self.template_pdf_path, target_file, field_values, flatten=True)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Stempel-PDF konnte nicht erzeugt werden:\n{exc}")
            return

        self.filled_stamp_pdf_path = target_file
        self.stamp_output_label.setText(f"Stempel-PDF (temp): {target_file}")
        self._apply_stamp_size_defaults(self.filled_stamp_pdf_path)
        self.log(f"Stempel-PDF erzeugt: {target_file}")

    def select_input_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Eingabeordner wählen")
        if folder:
            self.input_dir_path = Path(folder)
            self.input_dir_edit.setText(folder)

    def select_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Ausgabeordner wählen")
        if folder:
            self.output_dir_path = Path(folder)
            self.output_dir_edit.setText(folder)

    def run_batch(self) -> None:
        if self.filled_stamp_pdf_path is None:
            self.generate_stamp_pdf()
            if self.filled_stamp_pdf_path is None:
                QMessageBox.warning(self, "Hinweis", "Stempel-PDF konnte nicht automatisch erzeugt werden.")
                return
        if self.input_dir_path is None or self.output_dir_path is None:
            QMessageBox.warning(self, "Hinweis", "Bitte Eingabe- und Ausgabeordner wählen.")
            return

        settings = self._collect_settings()
        config = BatchJobConfig(
            input_dir=self.input_dir_path,
            output_dir=self.output_dir_path,
            stamp_pdf=self.filled_stamp_pdf_path,
            settings=settings,
        )
        processor = BatchProcessor(config)

        self.log("Stapelverarbeitung gestartet...")
        self.progress_bar.setValue(0)
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
                self.log(f"OK: {file_result.input_file.name} -> {file_result.output_file}")
                for pr in file_result.page_results:
                    if pr.status == "no_position":
                        no_position_count += 1
                    if pr.scale < 0.999:
                        reduced_count += 1
                    self.log(
                        f"  Seite {pr.page_index + 1}: {pr.status}, scale={pr.scale:.2f}, occ={pr.occupancy_ratio:.4f}, rect={pr.rect}"
                    )
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
