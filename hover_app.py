import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel,
    QFileDialog, QMessageBox, QComboBox, QLineEdit, QHBoxLayout
)

from PyQt5.QtCore import Qt
from pathlib import Path


class ExcelHoverApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.selected_sheet = None
        self.dataframe = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Excel Hover Info App')
        self.setFixedSize(400, 180)

        layout = QVBoxLayout()

        self.label = QLabel('No Excel file selected')
        self.label.setWordWrap(True)
        self.label.setAlignment(Qt.AlignCenter)

        # File + Sheet row
        file_sheet_layout = QHBoxLayout()

        self.btn_choose = QPushButton('Choose Excel File')
        self.btn_choose.clicked.connect(self.choose_excel)

        self.sheet_dropdown = QComboBox()
        self.sheet_dropdown.setEnabled(False)
        self.sheet_dropdown.currentIndexChanged.connect(self.sheet_selected)

        file_sheet_layout.addWidget(self.btn_choose)
        file_sheet_layout.addWidget(self.sheet_dropdown)

        self.btn_start = QPushButton('Start Hover Detection')
        self.btn_start.clicked.connect(self.start_hover)
        self.btn_start.setEnabled(False)

        search_layout = QHBoxLayout()
        self.input_search = QLineEdit()
        self.input_search.setPlaceholderText("Enter value to search")
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.on_search_clicked)
        self.btn_search.setEnabled(False)  # only active after loading data

        search_layout.addWidget(self.input_search)
        search_layout.addWidget(self.btn_search)


        layout.addWidget(self.label)
        layout.addLayout(file_sheet_layout)
        layout.addWidget(self.btn_start)
        layout.addLayout(search_layout)

        self.setLayout(layout)

    def choose_excel(self):
        documents_path = str(Path.home() / "Documents")
        path, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', documents_path, 'Excel Files (*.xlsx *.xls *.xlsm)')
        if path:
            self.excel_path = path
            self.label.setText(f'Selected: {path.split("/")[-1]}')

            try:
                xls = pd.ExcelFile(path, engine='openpyxl')
                sheets = xls.sheet_names

                self.sheet_dropdown.clear()
                self.sheet_dropdown.addItems(sheets)
                self.sheet_dropdown.setEnabled(True)

                # Reset sheet selection
                self.selected_sheet = None
                self.btn_start.setEnabled(False)

            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to load Excel file:\n{str(e)}')
                self.sheet_dropdown.clear()
                self.sheet_dropdown.setEnabled(False)
                self.btn_start.setEnabled(False)

    def sheet_selected(self):
        if self.sheet_dropdown.count() > 0:
            self.selected_sheet = self.sheet_dropdown.currentText()

            # Enable Start button only if both file and sheet are valid
            if self.excel_path and self.selected_sheet:
                self.btn_start.setEnabled(True)
        else:
            self.btn_start.setEnabled(False)

    def start_hover(self):
        if not self.excel_path or not self.selected_sheet:
            QMessageBox.warning(self, 'Missing Info', 'Please select an Excel file and sheet first.')
            return

        try:
            df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header = 1, engine='openpyxl')

            if df.empty:
                QMessageBox.warning(self, 'Empty Sheet', 'The selected sheet contains no data.')
                return
            
            QMessageBox.information(
                self,
                'Excel Loaded',
                f'Sheet: {self.selected_sheet}'
            )

            self.dataframe = df
            self.btn_search.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to load sheet:\n{str(e)}')
    
    def on_search_clicked(self):
        if self.dataframe is None or self.dataframe.empty:
            QMessageBox.warning(self, "No Data", "Excel data is not loaded.")
            return

        search_value = self.input_search.text().strip()
        if not search_value:
            QMessageBox.information(self, "Input Needed", "Please enter a value to search.")
            return

        # Step 1: Get the first column name
        first_col = self.dataframe.columns[0]

        # Step 2: Find the row where the first column matches
        matches = self.dataframe[self.dataframe[first_col].astype(str) == search_value]

        if matches.empty:
            QMessageBox.information(self, "No Match", f"No match found for: {search_value}")
            return

        row = matches.iloc[0]

        # Step 3: Find the column with header "名稱"
        name_column = None
        for col in self.dataframe.columns:
            if str(col).strip() == "名稱":
                name_column = col
                break

        if not name_column:
            QMessageBox.warning(self, "Missing Column", "Column with header '名稱' not found.")
            return

        name_value = row[name_column]

        QMessageBox.information(self, "Match Found", f"名稱: {name_value}")



    def search_value(self, search_term: str):
        if self.dataframe is None or self.dataframe.empty:
            QMessageBox.warning(self, "No Data", "No Excel data loaded.")
            return

        # We'll search only in the first column
        first_col = self.dataframe.columns[0]

        # Search for the value (as string or numeric match)
        matches = self.dataframe[self.dataframe[first_col].astype(str) == str(search_term)]

        if not matches.empty:
            row_data = matches.iloc[0].to_dict()
            msg = "\n".join(f"{k}: {v}" for k, v in row_data.items())

            QMessageBox.information(self, "Match Found", msg)
        else:
            QMessageBox.information(self, "No Match", f"No match found for: {search_term}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelHoverApp()
    window.show()
    sys.exit(app.exec_())
