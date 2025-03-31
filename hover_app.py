import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt


class ExcelHoverApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Excel Hover Info App')
        self.setFixedSize(300, 150)

        layout = QVBoxLayout()

        self.label = QLabel('No Excel file selected')
        self.label.setWordWrap(True)
        self.label.setAlignment(Qt.AlignCenter)

        self.btn_choose = QPushButton('Choose Excel File')
        self.btn_choose.clicked.connect(self.choose_excel)

        self.btn_start = QPushButton('Start Hover Detection')
        self.btn_start.clicked.connect(self.start_hover)
        self.btn_start.setEnabled(False)

        layout.addWidget(self.label)
        layout.addWidget(self.btn_choose)
        layout.addWidget(self.btn_start)

        self.setLayout(layout)

    def choose_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', '', 'Excel Files (*.xlsx *.xls *.xlsm)')
        if path:
            self.excel_path = path
            self.label.setText(f'Selected: {path.split("/")[-1]}')
            self.btn_start.setEnabled(True)

    def start_hover(self):
        if not self.excel_path:
            QMessageBox.warning(self, 'No File', 'Please select an Excel file first.')
            return

        try:
            df = pd.read_excel(self.excel_path, engine='openpyxl')

            preview = df.head().to_string(index=False)
            QMessageBox.information(
                self,
                'Excel Loaded',
                f'Loaded {len(df)} rows.\n\nPreview:\n{preview}'
            )

            # We'll store this for use later in hover detection
            self.dataframe = df

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to load Excel file:\n{str(e)}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelHoverApp()
    window.show()
    sys.exit(app.exec_())
