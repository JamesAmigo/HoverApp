from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel,
    QComboBox, QLineEdit, QPushButton
)
from PyQt5.QtGui import QIntValidator
from PyQt5.QtCore import pyqtSignal


class SheetSearchBar(QWidget):
    # creates signals
    searchRequested = pyqtSignal()
    fileChanged = pyqtSignal(str)
    sheetChanged = pyqtSignal(str)
    headerChanged = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Main Layout
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        # --- File / Sheet / Header Row ---
        file_row = QHBoxLayout()

        self.file_dropdown = QComboBox()
        self.file_dropdown.currentIndexChanged.connect(self._emit_file_changed)

        self.sheet_dropdown = QComboBox()
        self.sheet_dropdown.currentIndexChanged.connect(self._emit_sheet_changed)

        self.header_input = QLineEdit("2")
        self.header_input.setFixedWidth(50)
        self.header_input.setValidator(QIntValidator(1, 99999))
        self.header_input.textChanged.connect(self._emit_header_changed)

        for label, widget in [
            ("File Name:", self.file_dropdown),
            ("Sheet Name:", self.sheet_dropdown),
            ("Header Row:", self.header_input)
        ]:
            file_row.addWidget(QLabel(label))
            file_row.addWidget(widget)

        self.layout.addLayout(file_row)

        # --- Search Input ---
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter index to search")
        self.search_input.setValidator(QIntValidator(1, 9999999))
        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self.searchRequested.emit)

        search_row.addWidget(self.search_input)
        search_row.addWidget(self.search_btn)

        self.layout.addLayout(search_row)

    # --- Public Methods ---

    # Returns current state of the seach bar as a dict
    def get_selection(self):
        return {
            "file": self.file_dropdown.currentText(),
            "sheet": self.sheet_dropdown.currentText(),
            "header": int(self.header_input.text() or 1),
            "index": self.search_input.text().strip()
        }

    # Updates the file and sheet dropdowns
    def set_files_and_sheets(self, excel_files, file_sheets_dict, selected_file=None, selected_sheet=None):
        self.file_dropdown.blockSignals(True)
        self.sheet_dropdown.blockSignals(True)

        self.file_dropdown.clear()
        self.sheet_dropdown.clear()

        self.file_dropdown.addItems(excel_files.keys())

        file = selected_file or next(iter(excel_files), "")
        self.file_dropdown.setCurrentText(file)

        if file in file_sheets_dict:
            self.sheet_dropdown.addItems(file_sheets_dict[file])
            sheet = selected_sheet or file_sheets_dict[file][0]
            self.sheet_dropdown.setCurrentText(sheet)

        self.file_dropdown.blockSignals(False)
        self.sheet_dropdown.blockSignals(False)

        # After updating the dropdowns, emit the changed signals so that other widgets can call the functions they need
        self._emit_file_changed()
        self._emit_sheet_changed()
        self._emit_header_changed()


    # Sets the header row
    def set_header_row(self, row):
        self.header_input.setText(str(row))

    # Sets the index value
    def set_index_value(self, value):
        self.search_input.setText(str(value))

    # --- Private Emitters ---
    # Emits need to emit with arguments that match the receiving functions
    
    def _emit_file_changed(self):
        self.fileChanged.emit(self.file_dropdown.currentText())

    def _emit_sheet_changed(self):
        self.sheetChanged.emit(self.sheet_dropdown.currentText())

    def _emit_header_changed(self):
        text = self.header_input.text()
        if text.isdigit():
            self.headerChanged.emit(int(text))
