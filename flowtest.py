import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QGroupBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QLineEdit, QTextEdit,
    QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDoubleValidator
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from datetime import datetime


class FlowmeterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowmeter Quality Assurance Test (v1.7) @kadirsahin")
        self.setFixedHeight(600)
        self.setFixedWidth(1200)

        self.df = None
        self.meter_multiplier = "1"
        self.master_device_id = "-"
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # === Üst 3 sütun ===
        upper_layout = QHBoxLayout()

        # Flowmeter
        self.flowmeter_group = QGroupBox("Flowmeter")
        flowmeter_layout = QVBoxLayout()
        flowmeter_layout.setAlignment(Qt.AlignTop)
        flowmeter_layout.setSpacing(15)

        self.load_button = QPushButton("Select Flowmeter Data (xlsx)")
        self.load_button.clicked.connect(self.load_xlsx)
        flowmeter_layout.addWidget(self.load_button)

        self.data_table = QTableWidget()
        self.data_table.setColumnCount(2)
        self.data_table.setHorizontalHeaderLabels(["timestamp", "volume"])
        self.data_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.setSelectionMode(QAbstractItemView.MultiSelection)
        self.data_table.itemSelectionChanged.connect(self.update_summary)
        flowmeter_layout.addWidget(self.data_table)
        flowmeter_layout.addSpacing(10)

        self.flowmeter_group.setLayout(flowmeter_layout)
        upper_layout.addWidget(self.flowmeter_group, 1)

        # Water Meter
        self.watermeter_group = QGroupBox("Water Meter")
        watermeter_layout = QVBoxLayout()
        watermeter_layout.setAlignment(Qt.AlignTop)

        self.meter_type_label = QLabel("Pulse to Liter Coefficent:")
        self.meter_multiplier_input = QLineEdit("")
        self.meter_multiplier_input.textChanged.connect(self.update_summary)

        watermeter_layout.addWidget(self.meter_type_label)
        watermeter_layout.addWidget(self.meter_multiplier_input)

        # Test Start
        self.test_start_group = QGroupBox("Test Start")
        test_start_layout = QVBoxLayout()
        self.start_input = QLineEdit()
        self.start_input.setPlaceholderText("örn: 1,59")
        test_start_layout.addWidget(QLabel("Start Value"))
        test_start_layout.addWidget(self.start_input)
        self.test_start_group.setLayout(test_start_layout)
        watermeter_layout.addWidget(self.test_start_group)

        # Test End
        self.test_end_group = QGroupBox("Test End")
        test_end_layout = QVBoxLayout()
        self.end_input = QLineEdit()
        self.end_input.setPlaceholderText("örn: 3,25")
        test_end_layout.addWidget(QLabel("End Value"))
        test_end_layout.addWidget(self.end_input)
        self.test_end_group.setLayout(test_end_layout)
        watermeter_layout.addWidget(self.test_end_group)

        # Validator
        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)
        self.start_input.setValidator(validator)
        self.end_input.setValidator(validator)

        self.watermeter_group.setLayout(watermeter_layout)
        upper_layout.addWidget(self.watermeter_group, 1)

        # Summary
        self.report_group = QGroupBox("Summary")
        report_layout = QVBoxLayout()
        report_layout.setAlignment(Qt.AlignTop)
        report_layout.setSpacing(10)

        self.device_id_label = QLabel("Device ID: -")
        self.multiplier_label = QLabel(f"Multiplier: {self.meter_multiplier}")
        self.meter_result_label = QLabel("Water Meter Count: 0 lt")
        self.total_label = QLabel("Flowmeter Count: 0 lt")
        self.error_label = QLabel("Relative Error: -")
        self.test_approval_label = QLabel("Test Approval: -")

        for label in [
            self.device_id_label, self.multiplier_label,
            self.meter_result_label, self.total_label, self.error_label, self.test_approval_label
        ]:
            report_layout.addWidget(label)

        self.report_group.setLayout(report_layout)
        upper_layout.addWidget(self.report_group, 1)
        main_layout.addLayout(upper_layout, 1)

        # === Notes + Actions ===
        bottom_row = QHBoxLayout()
        notes_group = QGroupBox("Notes")
        notes_layout = QVBoxLayout()
        self.note_edit = QTextEdit()
        self.note_edit.setPlaceholderText("Please take notes here...")
        self.note_edit.setMinimumHeight(90)
        notes_layout.addWidget(self.note_edit)
        notes_group.setLayout(notes_layout)

        actions_group = QGroupBox("Actions")
        actions_layout = QVBoxLayout()
        actions_layout.setAlignment(Qt.AlignTop)
        self.export_button = QPushButton("Export Report")
        self.export_button.clicked.connect(self.export_report)
        actions_layout.addWidget(self.export_button)
        actions_layout.addStretch()
        actions_group.setLayout(actions_layout)

        bottom_row.addWidget(notes_group, 2)
        bottom_row.addWidget(actions_group, 1)
        main_layout.addLayout(bottom_row)

        # --- Canlı güncelleme için sinyaller ---
        self.start_input.textChanged.connect(self.update_summary)
        self.end_input.textChanged.connect(self.update_summary)

    def _parse_lineedit_float(self, le: QLineEdit) -> float:
        """QLineEdit içeriğini güvenli biçimde sayıya çevirir (TR ondalık desteği, boşluk temizliği)."""
        if le is None:
            return 0.0
        txt = (le.text() or "").strip()
        txt = txt.replace("\u00A0", "").replace(" ", "")
        txt = txt.replace(",", ".")
        if txt == "":
            return 0.0
        try:
            return float(txt)
        except Exception:
            return 0.0

    def _parse_float_cell(self, item: QTableWidgetItem) -> float:
        """Flow Counter (lt) hücresini güvenli biçimde sayıya çevirir."""
        if item is None:
            return 0.0
        txt = (item.text() or "").strip().replace(" ", "").replace("\u00A0", "")
        txt = txt.replace(",", ".")
        try:
            return float(txt)
        except Exception:
            return 0.0

    def load_xlsx(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open XLSX File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            self.df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "File Error", f"Excel couldn't loaded:\n{e}")
            return

        flow_col = "volume"

        ts_candidates = ["timestamp", "telemetryDate", "readingTime"]
        ts_col = next((c for c in ts_candidates if c in self.df.columns), None)

        if flow_col not in self.df.columns or ts_col is None:
            QMessageBox.warning(
                self,
                "Missing Columns",
                f"Required columns are not found.\n"
                f"Needed: '{flow_col}' and one of {ts_candidates}\n"
                f"Found: {list(self.df.columns)}"
            )
            return

        id_candidates = ["Production ID", "Master Device ID", "Asset Id"]
        id_col = next((c for c in id_candidates if c in self.df.columns), None)

        cols_to_take = [flow_col, ts_col] + ([id_col] if id_col else [])
        self.df = self.df[cols_to_take].dropna(subset=[flow_col, ts_col])

        if id_col and len(self.df) > 0:
            self.master_device_id = str(self.df[id_col].iloc[0])
        else:
            self.master_device_id = "-"
        self.device_id_label.setText(f"Device ID: {self.master_device_id}")

        self.data_table.setRowCount(len(self.df))
        for row in range(len(self.df)):
            time_item = QTableWidgetItem(str(self.df.iloc[row][ts_col]))
            flow_item = QTableWidgetItem(str(self.df.iloc[row][flow_col]))

            self.data_table.setItem(row, 0, time_item)
            self.data_table.setItem(row, 1, flow_item)

        self.update_summary()

    def get_meter_multiplier(self) -> float:
        txt = (self.meter_multiplier_input.text() or "").strip()
        txt = txt.replace(",", ".")
        try:
            val = float(txt)
            return val if val > 0 else 1.0
        except Exception:
            return 1.0

    def calculate_meter_volume(self):
        try:
            start = self._parse_lineedit_float(self.start_input)
            end = self._parse_lineedit_float(self.end_input)
            delta = max(0.0, end - start)
            return delta * self.get_meter_multiplier()
        except Exception:
            return 0.0

    def update_summary(self):
        selected = self.data_table.selectionModel().selectedRows() if self.data_table.selectionModel() else []
        rows_to_sum = [idx.row() for idx in selected]

        total = 0.0
        for r in rows_to_sum:
            total += self._parse_float_cell(self.data_table.item(r, 1))

        start = self._parse_lineedit_float(self.start_input)
        end = self._parse_lineedit_float(self.end_input)
        mult = self.get_meter_multiplier()

        delta = end - start

        self.multiplier_label.setText(f"Multiplier: {mult:g}")
        self.total_label.setText(f"Flowmeter Count: {total:.2f} lt")

        # --- HATALI DURUM: END < START ---
        if delta < 0:
            self.meter_result_label.setText(
                f"Water Meter Count: Invalid"
            )
            self.error_label.setText("Relative Error: -")
            self.test_approval_label.setText("Test Approval: NOT OK")
            self.test_approval_label.setStyleSheet("color: red;")
            return

        meter_volume = delta * mult

        self.meter_result_label.setText(
            f"Water Meter Count: {meter_volume:.2f} lt")

        # --- DEĞER YOK DURUMU ---
        if meter_volume == 0 or total == 0:
            self.error_label.setText("Relative Error: -")
            self.test_approval_label.setText("Test Approval: NOT OK")
            self.test_approval_label.setStyleSheet("color: red;")
            return

        # --- NORMAL HESAP ---
        error = abs(meter_volume - total) / meter_volume * 100
        self.error_label.setText(f"Relative Error: {error:.3f}%")

        if error < 1.0:
            self.test_approval_label.setText("Test Approval: OK")
            self.test_approval_label.setStyleSheet("color: green;")
        else:
            self.test_approval_label.setText("Test Approval: NOT OK")
            self.test_approval_label.setStyleSheet("color: red;")

    def normalize_text(self, text: str) -> str:
        mapping = {
            "ı": "i", "İ": "I",
            "ö": "o", "Ö": "O",
            "ü": "u", "Ü": "U",
            "ş": "s", "Ş": "S",
            "ç": "c", "Ç": "C",
            "ğ": "g", "Ğ": "G",
        }
        return "".join(mapping.get(ch, ch) for ch in text)

    def export_report(self):
        if self.df is None:
            QMessageBox.warning(self, "No Data", "Lütfen önce bir Excel dosyası yükleyin.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export PDF",
            f"report_{self.master_device_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        doc = SimpleDocTemplate(file_path, pagesize=A4,
                                rightMargin=2 * cm, leftMargin=2 * cm,
                                topMargin=2 * cm, bottomMargin=2 * cm)
        elements = []
        styles = getSampleStyleSheet()

        try:
            start_val = self._parse_lineedit_float(self.start_input)
            end_val = self._parse_lineedit_float(self.end_input)
            delta_val = max(0.0, end_val - start_val)
            meter_litre = self.calculate_meter_volume()
        except Exception:
            start_val = end_val = delta_val = meter_litre = 0.0

        elements.append(Paragraph("Doktar Flowmeter Quality Assurance Form", styles["Title"]))
        elements.append(Spacer(1, 1 * cm))

        approval_text = self.test_approval_label.text()
        approval_color = colors.red if "NOT" in approval_text.upper() else colors.green

        summary_data = [
            [f"{self.device_id_label.text()}"],
            [f"{self.multiplier_label.text()}"],
            [f"Water Meter Start Value: {start_val:.2f}"],
            [f"Water Meter End Value: {end_val:.2f}"],
            [f"Water Meter Measurement: {delta_val:.2f}"],
            [f"Water Meter Count: {meter_litre:.2f} lt"],
            [f"{self.total_label.text()}"],
            [f"{self.error_label.text()} < 1%"],
            [f"{approval_text}"],
        ]

        summary_table = Table(summary_data, colWidths=[doc.width])
        summary_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('FONTNAME', (0, 8), (0, 8), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 8), (0, 8), approval_color),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 1 * cm))

        selected = self.data_table.selectionModel().selectedRows() if self.data_table.selectionModel() else []
        rows_to_use = [idx.row() for idx in selected] if selected else list(range(self.data_table.rowCount()))

        if rows_to_use:
            data_rows = [["#", "Timestamp", "Volume (lt)"]]
            for i, r in enumerate(rows_to_use, start=1):
                ts = self.data_table.item(r, 0).text() if self.data_table.item(r, 0) else ""
                flow = self.data_table.item(r, 1).text() if self.data_table.item(r, 1) else ""
                data_rows.append([str(i), ts, flow])

            data_table = Table(data_rows, colWidths=[2 * cm, doc.width - 7 * cm, 5 * cm])
            data_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
            ]))
            elements.append(Paragraph("Flowmeter Datas", styles["Title"]))
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(data_table)
            elements.append(Spacer(1, 0.7 * cm))
        else:
            elements.append(Paragraph("No flowmeter data selected.", styles["Normal"]))
            elements.append(Spacer(1, 0.7 * cm))

        note_txt = self.note_edit.toPlainText().strip()
        if note_txt:
            note_txt = self.normalize_text(note_txt)
            note_txt = note_txt.replace("\n", "<br/>")

            elements.append(Paragraph("Notes", styles["Title"]))
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(Paragraph(note_txt, styles["Normal"]))
            elements.append(Spacer(1, 0.5 * cm))

        doc.build(elements)
        QMessageBox.information(self, "Export Completed", "PDF report exported.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowmeterApp()
    window.show()
    sys.exit(app.exec_())