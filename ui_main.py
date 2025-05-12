import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QMessageBox, QTabWidget
)
from PyQt5.QtCore import Qt

from widgets.loading_spinner import LoadingSpinner
from widgets.flow_layout import FlowLayout
from widgets.add_column_dialog import AddColumnDialog
from widgets.column_chip import ColumnChip
from widgets.result_dialog import ResultDialog
from widgets.sheet_search_bar import SheetSearchBar

from utilities.theme_utils import load_theme_preference, apply_theme, get_next_theme, get_theme_icon
from utilities.sheet_header_utils import get_header_index, set_header_index
from utilities.sheet_load_utils import load_excel_sheet, clean_column_name, get_first_match, get_saved_header_row, get_row_dict


class ExcelFolderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_files = {}
        self.file_sheets_map = {}
        self.current_sheet = None
        self.shown_columns = []
        self.current_theme = load_theme_preference()
        self.open_result_dialogs = []

        self.init_ui()
        apply_theme(self.current_theme)
        self.update_theme_switch(self.current_theme)

    def init_ui(self):
        self.setWindowTitle('Excel Folder Search Tool')
        self.setMinimumWidth(700)
        self.setMinimumHeight(300)
        main_layout = QVBoxLayout()

        folder_row = QHBoxLayout()
        self.btn_choose_folder = QPushButton('Choose Folder')
        self.btn_choose_folder.setFixedHeight(40)
        self.btn_choose_folder.clicked.connect(self.choose_folder)
        self.label_info = QLabel('No folder selected')
        self.label_info.setFixedHeight(40)
        self.spinner = LoadingSpinner(self)
        self.theme_switch = QPushButton("🌙")
        self.theme_switch.setFixedSize(40, 40)
        self.theme_switch.clicked.connect(self.toggle_theme_icon)
        self.theme_switch.setProperty("role", "ThemeSwitch")

        folder_row.addWidget(self.btn_choose_folder)
        folder_row.addWidget(self.label_info)
        folder_row.addWidget(self.spinner)
        folder_row.addStretch()
        folder_row.addWidget(self.theme_switch)

        main_layout.addLayout(folder_row)
        main_layout.addSpacing(10)

        self.tabs = QTabWidget()
        search_tab = self.build_search_tab()
        self.tabs.addTab(search_tab, "Search")
        self.tabs.addTab(QLabel("Hover functionality coming soon..."), "Hover")
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)

    def build_search_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.sheet_search_bar = SheetSearchBar()
        self.sheet_search_bar.fileChanged.connect(self.on_file_selected)
        self.sheet_search_bar.sheetChanged.connect(self.on_sheet_selected)
        self.sheet_search_bar.headerChanged.connect(self.on_header_row_change)
        self.sheet_search_bar.searchRequested.connect(self.perform_search)

        layout.addWidget(self.sheet_search_bar)

        self.chip_layout = FlowLayout()
        chip_container = QWidget()
        chip_container.setLayout(self.chip_layout)
        layout.addWidget(chip_container)

        self.btn_add_column = QPushButton("Add Column")
        self.btn_add_column.clicked.connect(self.show_add_column_dialog)
        layout.addWidget(self.btn_add_column)

        tab.setLayout(layout)
        return tab

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if not folder:
            return

        self.spinner.start()
        self.label_info.setText(f'{folder}')
        QApplication.processEvents()

        self.excel_files.clear()
        self.file_sheets_map.clear()
        self.current_sheet = None

        for file in os.listdir(folder):
            if file.startswith("~$") or not file.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            full_path = os.path.join(folder, file)
            try:
                xls = pd.ExcelFile(full_path, engine='openpyxl')
                self.excel_files[file] = full_path
                self.file_sheets_map[file] = xls.sheet_names
            except Exception as e:
                print(f"Failed to load sheets for {file}: {e}")

        if not self.excel_files:
            self.label_info.setText('No Excel files found in folder.')
            return

        default_file = next(iter(self.excel_files))
        self.sheet_search_bar.set_files_and_sheets(self.excel_files, self.file_sheets_map, selected_file=default_file)

        saved = get_saved_header_row(default_file, self.file_sheets_map[default_file][0])
        if saved:
            self.sheet_search_bar.set_header_row(saved)
        self.load_sheet_with_header()
        self.spinner.stop()

    def on_file_selected(self, file):
        sheets = self.file_sheets_map.get(file, [])
        self.sheet_search_bar.sheet_dropdown.clear()
        self.sheet_search_bar.sheet_dropdown.addItems(sheets)
        if sheets:
            self.sheet_search_bar.sheet_dropdown.setCurrentIndex(0)

    def on_sheet_selected(self, sheet_name):
        if not sheet_name:
            return
        self.load_header_row()
        self.load_sheet_with_header()

    def load_sheet_with_header(self):
        selection = self.sheet_search_bar.get_selection()
        file = selection["file"]
        sheet = selection["sheet"]
        header_row = selection["header"]
        file_path = self.excel_files.get(file)

        try:
            excel_sheet = load_excel_sheet(file_path, sheet, header_row)
            excel_sheet.columns = [clean_column_name(col) for col in excel_sheet.columns]
            self.current_sheet = excel_sheet
            self.update_column_scope()
        except Exception as e:
            self.label_info.setText(f'Error reading sheet: {str(e)}')

    def update_column_scope(self):
        if self.current_sheet is None:
            return
        self.shown_columns = list(self.current_sheet.columns)
        self.clear_chips()
        for col in self.shown_columns:
            chip = ColumnChip(col, self.remove_column_from_scope)
            self.chip_layout.addWidget(chip)

    def clear_chips(self):
        while self.chip_layout.count():
            item = self.chip_layout.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()

    def remove_column_from_scope(self, column_name):
        if column_name in self.shown_columns:
            self.shown_columns.remove(column_name)

    def show_add_column_dialog(self):
        if self.current_sheet is None:
            return
        remaining = [col for col in self.current_sheet.columns if col not in self.shown_columns]
        dialog = AddColumnDialog(remaining, self.add_column_to_scope)
        dialog.exec_()

    def add_column_to_scope(self, column_name):
        if column_name not in self.shown_columns:
            self.shown_columns.append(column_name)
            chip = ColumnChip(column_name, self.remove_column_from_scope)
            self.chip_layout.addWidget(chip)

    def perform_search(self):
        if self.current_sheet is None or self.current_sheet.empty:
            return

        selection = self.sheet_search_bar.get_selection()
        search_term = selection["index"]
        if not search_term:
            return

        match = get_first_match(self.current_sheet, search_term)
        if match is not None:
            row_dict = match.to_dict()
            dialog = ResultDialog(
                search_term,
                row_dict,
                self.shown_columns,
                self.excel_files,
                self.file_sheets_map,
                selection["file"],
                selection["sheet"]
            )
            dialog.setModal(False)
            dialog.show()
            self.open_result_dialogs.append(dialog)
        else:
            QMessageBox.information(self, "No Match", f"No match found for: {search_term}")

    def toggle_theme_icon(self):
        next_theme = get_next_theme(self.current_theme)
        if apply_theme(next_theme):
            self.current_theme = next_theme
            self.update_theme_switch(self.current_theme)

    def update_theme_switch(self, theme_name):
        self.theme_switch.setText(get_theme_icon(theme_name))

    def on_header_row_change(self, row):
        selection = self.sheet_search_bar.get_selection()
        set_header_index(selection["file"], selection["sheet"], row)
        self.load_sheet_with_header()

    def load_header_row(self):
        selection = self.sheet_search_bar.get_selection()
        saved = get_saved_header_row(selection["file"], selection["sheet"])
        if saved:
            self.sheet_search_bar.set_header_row(saved)
