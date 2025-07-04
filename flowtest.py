import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QGroupBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QComboBox, QLineEdit, QSpinBox, QMessageBox
)
from PyQt5.QtCore import Qt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from datetime import datetime


class FlowmeterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowmeter Quality Assurance Test")
        self.setFixedHeight(500)
        self.setFixedWidth(1000)

        self.df = None
        self.meter_multiplier = 10
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # === Upper Part ===
        upper_layout = QHBoxLayout()

        # === Flowmeter Group ===
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
        flowmeter_layout.addWidget(self.data_table)
        flowmeter_layout.addSpacing(10)

        self.flowmeter_group.setLayout(flowmeter_layout)
        upper_layout.addWidget(self.flowmeter_group, 1)

        # === Water Meter Group ===
        self.watermeter_group = QGroupBox("Water Meter")
        watermeter_layout = QVBoxLayout()
        watermeter_layout.setAlignment(Qt.AlignTop)

        self.meter_type_label = QLabel("Select Meter Type:")
        self.meter_type_combo = QComboBox()
        self.meter_type_combo.addItems([
            "Klepsan Woltman DN50",
            "Klepsan Woltman DN65",
            "Klepsan Woltman DN80",
            "Klepsan Woltman DN100",
            "Klepsan Woltman DN150"
        ])
        self.meter_type_combo.currentIndexChanged.connect(self.update_meter_info)
        self.meter_multiplier_map = {
            "Klepsan Woltman DN50": 10,
            "Klepsan Woltman DN65": 10,
            "Klepsan Woltman DN80": 10,
            "Klepsan Woltman DN100": 10,
            "Klepsan Woltman DN150": 100
        }

        watermeter_layout.addWidget(self.meter_type_label)
        watermeter_layout.addWidget(self.meter_type_combo)


        # === Test Start ===
        self.test_start_group = QGroupBox("Test Start")
        test_start_layout = QVBoxLayout()
        self.start_x1 = QLineEdit()
        self.start_x1.setPlaceholderText("")
        test_start_layout.addWidget(QLabel("x1"))
        test_start_layout.addWidget(self.start_x1)

        self.start_x01 = QSpinBox()
        self.start_x01.setRange(0, 9)
        test_start_layout.addWidget(QLabel("x0.1"))
        test_start_layout.addWidget(self.start_x01)

        self.start_x001 = QSpinBox()
        self.start_x001.setRange(0, 9)
        test_start_layout.addWidget(QLabel("x0.01"))
        test_start_layout.addWidget(self.start_x001)
        self.test_start_group.setLayout(test_start_layout)
        watermeter_layout.addWidget(self.test_start_group)

        # === Test End ===
        self.test_end_group = QGroupBox("Test End")
        test_end_layout = QVBoxLayout()
        self.end_x1 = QLineEdit()
        self.end_x1.setPlaceholderText("")
        self.end_x01 = QSpinBox()
        self.end_x01.setRange(0, 9)
        self.end_x001 = QSpinBox()
        self.end_x001.setRange(0, 9)
        test_end_layout.addWidget(QLabel("x1"))
        test_end_layout.addWidget(self.end_x1)
        test_end_layout.addWidget(QLabel("x0.1"))
        test_end_layout.addWidget(self.end_x01)
        test_end_layout.addWidget(QLabel("x0.01"))
        test_end_layout.addWidget(self.end_x001)
        self.test_end_group.setLayout(test_end_layout)
        watermeter_layout.addWidget(self.test_end_group)

        self.watermeter_group.setLayout(watermeter_layout)
        upper_layout.addWidget(self.watermeter_group, 1)

        # === Summary & Report Group ===
        self.report_group = QGroupBox("Summary and Report")
        report_layout = QVBoxLayout()
        report_layout.setAlignment(Qt.AlignTop)
        report_layout.setSpacing(10)

        self.device_id_label = QLabel("Device ID: -")
        self.meter_type_display = QLabel("Selected Meter: Klepsan Woltman DN50")
        self.multiplier_label = QLabel("Multiplier: 10.00")
        self.meter_result_label = QLabel("Water Meter Count: -")
        self.total_label = QLabel("Flowmeter Count: -")
        self.error_label = QLabel("Relative Error: -")
        self.test_approval_label = QLabel("Test Approval: -")

        self.perform_button = QPushButton("Perform Test")
        self.perform_button.clicked.connect(self.update_summary)
        self.export_button = QPushButton("Export Report")
        self.export_button.clicked.connect(self.export_report)

        for label in [
            self.device_id_label, self.meter_type_display, self.multiplier_label,
            self.meter_result_label, self.total_label, self.error_label, self.test_approval_label
        ]:
            report_layout.addWidget(label)

        report_layout.addStretch()
        report_layout.addWidget(self.perform_button)
        report_layout.addWidget(self.export_button)

        self.report_group.setLayout(report_layout)
        upper_layout.addWidget(self.report_group, 1)

        main_layout.addLayout(upper_layout, 1)

    def load_xlsx(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open XLSX File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "File Error", f"Excel yüklenemedi:\n{e}")
            return

        required_cols = {"Flow Counter (lt)", "Device TS Date", "Master Device ID"}
        if not required_cols.issubset(self.df.columns):
            QMessageBox.warning(self, "Missing Columns", "Gerekli sütunlar bulunamadı.")
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

    def update_meter_info(self):
        meter_type = self.meter_type_combo.currentText()
        self.meter_multiplier = self.meter_multiplier_map.get(meter_type, 1.0)
        self.meter_type_display.setText(f"Selected Meter: {meter_type}")
        self.multiplier_label.setText(f"Multiplier: {self.meter_multiplier:.2f}")

    def calculate_meter_volume(self):
        try:
            start = float(self.start_x1.text() or 0) + self.start_x01.value() / 10 + self.start_x001.value() / 100
            end = float(self.end_x1.text() or 0) + self.end_x01.value() / 10 + self.end_x001.value() / 100
            delta = max(0, end - start)
            result = delta * self.meter_multiplier
            return result
        except:
            return 0.0

    def update_summary(self):
        total = 0.0
        selected_rows = self.data_table.selectionModel().selectedRows()
        for index in selected_rows:
            try:
                value = float(self.data_table.item(index.row(), 0).text())
                total += value
            except:
                continue

        meter_volume = self.calculate_meter_volume()

        self.total_label.setText(f"Flowmeter Count: {total:.2f} liter")
        self.meter_result_label.setText(f"Water Meter Count: {meter_volume:.2f} liter")
        self.multiplier_label.setText(f"Multiplier: {self.meter_multiplier:.2f}")

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

    def export_report(self):
        if self.df is None:
            QMessageBox.warning(self, "No Data", "Lütfen önce bir Excel dosyası yükleyin.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Export PDF",
                                                   f"report_{self.master_device_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                                   "PDF Files (*.pdf)")
        if not file_path:
            return

        c = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4
        y = height - 2 * cm

        # Sayaç değerleri
        try:
            start_val = float(self.start_x1.text() or 0) + self.start_x01.value() / 10 + self.start_x001.value() / 100
            end_val = float(self.end_x1.text() or 0) + self.end_x01.value() / 10 + self.end_x001.value() / 100
            delta_val = max(0, end_val - start_val)
            meter_litre = delta_val * self.meter_multiplier
        except:
            start_val = end_val = delta_val = meter_litre = 0

        # Sayfa 1 - Özet
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2 * cm, y, "Flowmeter Quality Assurance Test Summary")
        y -= 1.5 * cm
        c.setFont("Helvetica", 12)

        summary_items = [
            self.device_id_label.text(),
            self.meter_type_display.text(),
            self.multiplier_label.text(),
            f"Water Meter Start Value: {start_val:.3f}",
            f"Water Meter End Value: {end_val:.3f}",
            f"Consumption: {delta_val:.3f} x Multiplier",
            f"Water Meter Count: {meter_litre:.2f} liter",
            self.total_label.text(),
            self.error_label.text(),
            self.test_approval_label.text(),
        ]
        for item in summary_items:
            c.drawString(2 * cm, y, item)
            y -= 1 * cm

        c.showPage()

        # Sayfa 2 - Seçilen Veriler
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2 * cm, height - 2 * cm, "Flowmeter Datas")
        c.setFont("Helvetica", 10)
        y = height - 3 * cm

        selected_rows = self.data_table.selectionModel().selectedRows()
        for i, index in enumerate(selected_rows):
            if y < 3 * cm:
                c.showPage()
                y = height - 2 * cm
                c.setFont("Helvetica", 10)
            flow = self.data_table.item(index.row(), 0).text()
            ts = self.data_table.item(index.row(), 1).text()
            c.drawString(2 * cm, y, f"{i + 1}. Water Count: {flow} lt   Timestamp: {ts}")
            y -= 0.7 * cm

        c.save()
        QMessageBox.information(self, "Export Completed", "PDF report exported.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowmeterApp()
    window.show()
    sys.exit(app.exec_())
