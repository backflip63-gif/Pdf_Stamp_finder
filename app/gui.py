from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Dict

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
