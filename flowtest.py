import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QGroupBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QComboBox, QLineEdit, QSpinBox, QTextEdit,
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
from reportlab.platypus import Paragraph


class FlowmeterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowmeter Quality Assurance Test (v1.5)")
        self.setFixedHeight(800)
        self.setFixedWidth(1000)

        self.df = None
        self.meter_multiplier = 1
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
        self.data_table.setHorizontalHeaderLabels(["Flow Counter (lt)", "Device TS Date"])
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

        self.meter_type_label = QLabel("Select Meter Multiplier:")
        self.meter_type_combo = QComboBox()
        self.meter_type_combo.addItems([
            "x1",
            "x10",
            "x100",
            "x1,000",
            "x10,000",
            "x100,000",
            "x1,000,000",
        ])

        self.meter_type_combo.currentIndexChanged.connect(self.update_meter_info)

        self.meter_multiplier_map = {
            "x1": 1,
            "x10": 10,
            "x100": 100,
            "x1,000": 1000,
            "x10,000": 10000,
            "x100,000": 100000,
            "x1,000,000": 1000000,
        }

        watermeter_layout.addWidget(self.meter_type_label)
        watermeter_layout.addWidget(self.meter_type_combo)

        # Test Start
        self.test_start_group = QGroupBox("Test Start")
        test_start_layout = QVBoxLayout()
        self.start_x1 = QLineEdit(); self.start_x1.setPlaceholderText("")
        self.start_x01 = QSpinBox(); self.start_x01.setRange(0, 9)
        self.start_x001 = QSpinBox(); self.start_x001.setRange(0, 9)
        test_start_layout.addWidget(QLabel("x1"));    test_start_layout.addWidget(self.start_x1)
        test_start_layout.addWidget(QLabel("x0.1"));  test_start_layout.addWidget(self.start_x01)
        test_start_layout.addWidget(QLabel("x0.01")); test_start_layout.addWidget(self.start_x001)
        self.test_start_group.setLayout(test_start_layout)
        watermeter_layout.addWidget(self.test_start_group)

        # Test End
        self.test_end_group = QGroupBox("Test End")
        test_end_layout = QVBoxLayout()
        self.end_x1 = QLineEdit(); self.end_x1.setPlaceholderText("")
        self.end_x01 = QSpinBox(); self.end_x01.setRange(0, 9)
        self.end_x001 = QSpinBox(); self.end_x001.setRange(0, 9)
        test_end_layout.addWidget(QLabel("x1"));    test_end_layout.addWidget(self.end_x1)
        test_end_layout.addWidget(QLabel("x0.1"));  test_end_layout.addWidget(self.end_x01)
        test_end_layout.addWidget(QLabel("x0.01")); test_end_layout.addWidget(self.end_x001)
        self.test_end_group.setLayout(test_end_layout)
        watermeter_layout.addWidget(self.test_end_group)

        # --- QLineEdit doğrulayıcıları (ondalık girişler için) ---
        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)
        self.start_x1.setValidator(validator)
        self.end_x1.setValidator(validator)

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
        actions_layout = QVBoxLayout(); actions_layout.setAlignment(Qt.AlignTop)
        self.export_button = QPushButton("Export Report")
        self.export_button.clicked.connect(self.export_report)
        actions_layout.addWidget(self.export_button)
        actions_layout.addStretch()
        actions_group.setLayout(actions_layout)

        bottom_row.addWidget(notes_group, 2)
        bottom_row.addWidget(actions_group, 1)
        main_layout.addLayout(bottom_row)

        # --- Canlı güncelleme için sinyaller ---
        self.start_x1.textChanged.connect(self.update_summary)
        self.end_x1.textChanged.connect(self.update_summary)
        self.start_x01.valueChanged.connect(self.update_summary)
        self.start_x001.valueChanged.connect(self.update_summary)
        self.end_x01.valueChanged.connect(self.update_summary)
        self.end_x001.valueChanged.connect(self.update_summary)
        self.meter_type_combo.currentIndexChanged.connect(self.update_summary)

    # --- Güvenli parse yardımcıları ---
    def _parse_lineedit_float(self, le: QLineEdit) -> float:
        """QLineEdit içeriğini güvenli biçimde sayıya çevirir (TR ondalık desteği, boşluk temizliği)."""
        if le is None:
            return 0.0
        txt = (le.text() or "").strip()
        # normal ve non-breaking space temizliği
        txt = txt.replace("\u00A0", "").replace(" ", "")
        # TR ondalık virgül desteği
        txt = txt.replace(",", ".")
        if txt == "":
            return 0.0
        try:
            return int(txt)
        except Exception:
            return 0.0

    def _parse_float_cell(self, item: QTableWidgetItem) -> float:
        """Flow Counter (lt) hücresini güvenli biçimde sayıya çevirir."""
        if item is None:
            return 0.0
        txt = (item.text() or "").strip().replace(" ", "").replace("\u00A0", "")
        txt = txt.replace(",", ".")  # TR ondalık desteği
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

        required_cols = {"Flow Counter (lt)", "Device TS Date", "Master Device ID"}
        if not required_cols.issubset(self.df.columns):
            QMessageBox.warning(self, "Missing Columns", "Columns are not found.")
            return

        self.df = self.df[["Flow Counter (lt)", "Device TS Date", "Master Device ID"]].dropna()
        self.master_device_id = str(self.df["Master Device ID"].iloc[0])
        self.device_id_label.setText(f"Device ID: {self.master_device_id}")

        self.data_table.setRowCount(len(self.df))
        for row in range(len(self.df)):
            flow_item = QTableWidgetItem(str(self.df.iloc[row]["Flow Counter (lt)"]))
            time_item = QTableWidgetItem(str(self.df.iloc[row]["Device TS Date"]))
            self.data_table.setItem(row, 0, flow_item)
            self.data_table.setItem(row, 1, time_item)

        # tablo yüklenince toplamı hemen göster
        self.update_summary()

    def update_meter_info(self):
        meter_type = self.meter_type_combo.currentText()
        self.meter_multiplier = self.meter_multiplier_map.get(meter_type, 1.0)
        self.multiplier_label.setText(f"Multiplier: {self.meter_multiplier}")

    def calculate_meter_volume(self):
        try:
            start = self._parse_lineedit_float(self.start_x1) \
                    + self.start_x01.value() / 10 + self.start_x001.value() / 100
            end = self._parse_lineedit_float(self.end_x1) \
                  + self.end_x01.value() / 10 + self.end_x001.value() / 100
            delta = max(0.0, end - start)
            return delta * self.meter_multiplier
        except Exception:
            return 0.0

    def update_summary(self):
        selected = self.data_table.selectionModel().selectedRows() if self.data_table.selectionModel() else []
        rows_to_sum = [idx.row() for idx in selected] if selected else list(range(self.data_table.rowCount()))

        total = 0.0
        for r in rows_to_sum:
            total += self._parse_float_cell(self.data_table.item(r, 0))

        meter_volume = self.calculate_meter_volume()

        self.total_label.setText(f"Flowmeter Count: {total:.0f} lt")
        self.meter_result_label.setText(f"Water Meter Count: {meter_volume:.0f} lt")
        self.multiplier_label.setText(f"Multiplier: {self.meter_multiplier}")

        if meter_volume == 0 or total == 0:
            self.error_label.setText("Relative Error: -")
            self.test_approval_label.setText("Test Approval: -")
            self.test_approval_label.setStyleSheet("color: black;")
            return

        error = abs(meter_volume - total) / meter_volume * 100
        self.error_label.setText(f"Relative Error: {error:.2f}%")
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

        # Sayaç değerleri (güvenli parse ile)
        try:
            start_val = self._parse_lineedit_float(self.start_x1) \
                        + self.start_x01.value() / 10 + self.start_x001.value() / 100
            end_val = self._parse_lineedit_float(self.end_x1) \
                      + self.end_x01.value() / 10 + self.end_x001.value() / 100
            delta_val = max(0.0, end_val - start_val)

            # burası artık direkt fonksiyondan:
            meter_litre = self.calculate_meter_volume()
        except Exception:
            start_val = end_val = delta_val = meter_litre = 0.0

        # Başlık
        elements.append(Paragraph("Doktar Flowmeter Quality Assurance Form", styles["Title"]))
        elements.append(Spacer(1, 1 * cm))

        # Özet
        approval_text = self.test_approval_label.text()
        approval_color = colors.red if "NOT" in approval_text.upper() else colors.green
        summary_data = [
            [f"{self.device_id_label.text()}"],
            [f"{self.multiplier_label.text()}"],
            [f"Water Meter Start Value: {start_val:.2f}"],
            [f"Water Meter End Value: {end_val:.2f}"],
            [f"Water Meter Measurement: {delta_val:.2f}"],
            [f"Water Meter Count: {meter_litre:.0f} lt"],
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
            ('FONTNAME', (0, 10), (0, 10), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 10), (0, 10), approval_color),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 1 * cm))

        # Seçilen veriler
        selected = self.data_table.selectionModel().selectedRows() if self.data_table.selectionModel() else []
        rows_to_use = [idx.row() for idx in selected] if selected else list(range(self.data_table.rowCount()))
        if rows_to_use:
            data_rows = [["#", "Water Count (lt)", "Timestamp"]]
            for i, r in enumerate(rows_to_use, start=1):
                flow = (self.data_table.item(r, 0).text() if self.data_table.item(r, 0) else "")
                ts = (self.data_table.item(r, 1).text() if self.data_table.item(r, 1) else "")
                data_rows.append([str(i), flow, ts])

            data_table = Table(data_rows, colWidths=[2 * cm, 5 * cm, doc.width - 7 * cm])
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

        # Notes
        note_txt = self.note_edit.toPlainText().strip()
        if note_txt:
            note_txt = self.normalize_text(note_txt)  # Türkçe → İngilizce harf
            # Satır sonlarını <br/> ile değiştir
            note_txt = note_txt.replace("\n", "<br/>")

            elements.append(Paragraph("Notes", styles["Title"]))
            elements.append(Spacer(1, 0.3 * cm))
            # Paragraph ile yaz
            elements.append(Paragraph(note_txt, styles["Normal"]))
            elements.append(Spacer(1, 0.5 * cm))

        # PDF
        doc.build(elements)
        QMessageBox.information(self, "Export Completed", "PDF report exported.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowmeterApp()
    window.show()
    sys.exit(app.exec_())
