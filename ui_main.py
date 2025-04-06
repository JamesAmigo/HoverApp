import sys
import os
import json
import pandas as pd
from widgets.loading_spinner import LoadingSpinner
from widgets.flow_layout import FlowLayout
from widgets.add_column_dialog import AddColumnDialog
from widgets.column_chip import ColumnChip
from widgets.result_dialog import ResultDialog
from utilities.resource_utils import load_stylesheet, save_theme_preference, load_theme_preference
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QComboBox, QLineEdit, QMessageBox, QTabWidget
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

class ExcelFolderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_files = {}
        self.current_df = None
        self.shown_columns = []
        self.themes = {}
        resources_path = os.path.join(os.path.dirname(__file__), "Resources")
        for filename in os.listdir(resources_path):
            if filename.endswith(".qss"):
                theme_name = filename.split(".")[0]
                full_path = os.path.join(resources_path, filename)
                self.themes[theme_name] = load_stylesheet(full_path)


        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Excel Folder Search Tool')
        self.setMinimumWidth(700)
        self.setMinimumHeight(300)

        main_layout = QVBoxLayout()

        # --- Folder Row ---
        folder_row = QHBoxLayout()


        self.btn_choose_folder = QPushButton('Choose Folder')
        self.btn_choose_folder.setMinimumSize(100, 30)
        self.btn_choose_folder.setMaximumSize(150, 60)
        self.btn_choose_folder.clicked.connect(self.choose_folder)
        folder_row.addWidget(self.btn_choose_folder)


        self.label_info = QLabel('No folder selected')
        self.label_info.setMinimumWidth(100)
        self.label_info.setMaximumWidth(200)
        folder_row.addWidget(self.label_info)

        self.spinner = LoadingSpinner(self)
        folder_row.addWidget(self.spinner)

        folder_row.addStretch()

        self.theme_switch = QPushButton("🌙")
        self.theme_switch.setFixedSize(40, 40)
        self.theme_switch.setToolTip("Toggle Light/Dark Theme")
        self.theme_switch.clicked.connect(self.toggle_theme_icon)
        self.theme_switch.setProperty("role", "ThemeSwitch")
        folder_row.addWidget(self.theme_switch)

        main_layout.addLayout(folder_row)
        main_layout.addSpacing(10)

        # --- Tab Widget ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # === Tab 1: Search Tab ===
        search_tab = QWidget()
        search_tab_layout = QVBoxLayout()

        # File/Sheet Dropdowns
        file_sheet_layout = QVBoxLayout()

        top_row = QHBoxLayout()
        file_label = QLabel("File name:")
        file_label.setFixedSize(70, 30)
        self.file_dropdown = QComboBox()
        self.file_dropdown.setEnabled(False)
        self.file_dropdown.currentIndexChanged.connect(self.on_file_selected)
        top_row.addWidget(file_label)
        top_row.addWidget(self.file_dropdown)

        bottom_row = QHBoxLayout()
        sheet_label = QLabel("Sheet name:")
        sheet_label.setFixedSize(70, 30)
        self.sheet_dropdown = QComboBox()
        self.sheet_dropdown.setEnabled(False)
        self.sheet_dropdown.currentIndexChanged.connect(self.on_sheet_selected)
        bottom_row.addWidget(sheet_label)
        bottom_row.addWidget(self.sheet_dropdown)

        file_sheet_layout.addLayout(top_row)
        file_sheet_layout.addLayout(bottom_row)
        search_tab_layout.addLayout(file_sheet_layout)

        # Search Input
        search_input_layout = QHBoxLayout()
        self.input_search = QLineEdit()
        self.input_search.setPlaceholderText("Enter index to search")
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.perform_search)
        search_input_layout.addWidget(self.input_search)
        search_input_layout.addWidget(self.btn_search)
        search_tab_layout.addLayout(search_input_layout)

        # Column chips
        self.chip_container_widget = QWidget()
        self.chip_layout = FlowLayout()
        self.chip_container_widget.setLayout(self.chip_layout)

        self.btn_add_column = QPushButton("Add Column")
        self.btn_add_column.clicked.connect(self.show_add_column_dialog)

        search_tab_layout.addWidget(self.chip_container_widget)
        search_tab_layout.addWidget(self.btn_add_column)

        search_tab.setLayout(search_tab_layout)
        self.tabs.addTab(search_tab, "Search")

        # === Tab 2: Hover Placeholder ===
        hover_tab = QWidget()
        hover_tab_layout = QVBoxLayout()
        hover_tab_layout.addWidget(QLabel("Hover functionality coming soon..."))
        hover_tab.setLayout(hover_tab_layout)
        self.tabs.addTab(hover_tab, "Hover")

        # === Finalize Layout ===
        self.setLayout(main_layout)

        # Load and apply theme
        self.current_theme = load_theme_preference()
        self.apply_theme(self.current_theme)
        self.update_theme_icon()

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if not folder:
            return

        self.label_info.setText(f'{folder}')
        self.excel_files.clear()
        self.file_dropdown.clear()
        self.sheet_dropdown.clear()
        self.file_dropdown.setEnabled(False)
        self.sheet_dropdown.setEnabled(False)
        self.current_df = None

        for file in os.listdir(folder):
            if file.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                full_path = os.path.join(folder, file)
                self.excel_files[file] = full_path

        if not self.excel_files:
            self.label_info.setText('No Excel files found in folder.')
            return

        self.file_dropdown.addItems(self.excel_files.keys())
        self.file_dropdown.setEnabled(True)
        

    def on_file_selected(self, index):
        self.sheet_dropdown.clear()
        self.sheet_dropdown.setEnabled(False)
        self.current_df = None

        self.spinner.start()
        QApplication.processEvents()
        if index < 0:
            return

        filename = self.file_dropdown.currentText()
        file_path = self.excel_files[filename]

        try:
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            self.sheet_dropdown.addItems(xls.sheet_names)
            self.sheet_dropdown.setEnabled(True)
        except Exception as e:
            self.label_info.setText(f'Error loading file: {str(e)}')
        
        self.spinner.stop()

    def on_sheet_selected(self, index):
        
        self.spinner.start()
        QApplication.processEvents()
        if index < 0:
            return

        filename = self.file_dropdown.currentText()
        file_path = self.excel_files[filename]
        sheet = self.sheet_dropdown.currentText()
        if not sheet:
            self.label_info.setText("No sheet selected.")
            return

        try:
            df = pd.read_excel(file_path, sheet_name=sheet, header=1, engine='openpyxl')
            df.columns = [self.clean_column_name(col) for col in df.columns]

            # Drop columns that are unnamed or fully empty
            df = df.loc[:, ~df.columns.str.contains(r'^Unnamed', case=False)]
            df = df.dropna(axis=1, how='all')  # optional: remove columns where all rows are NaN

            self.current_df = df
            self.update_column_scope()
        except Exception as e:
            print(f'Error reading sheet: {str(e)}')
            self.label_info.setText(f'Error reading sheet: {str(e)}')
        
        self.spinner.stop()

    def clean_column_name(self, name):
        if not isinstance(name, str):
            return name
        name = name.replace('\n', ' ')             # Replace newline with space
        name = name.split('*')[0].strip()          # Remove anything after '*'
        return name

    def update_column_scope(self):
        if self.current_df is None:
            return

        self.shown_columns = list(self.current_df.columns)
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
        if self.current_df is None:
            return

        remaining = [col for col in self.current_df.columns if col not in self.shown_columns]
        dialog = AddColumnDialog(remaining, self.add_column_to_scope)
        dialog.exec_()

    def add_column_to_scope(self, column_name):
        if column_name not in self.shown_columns:
            self.shown_columns.append(column_name)
            chip = ColumnChip(column_name, self.remove_column_from_scope)
            self.chip_layout.addWidget(chip)

    def perform_search(self):
        if self.current_df is None or self.current_df.empty:
            return

        search_term = self.input_search.text().strip()
        if not search_term:
            return

        first_col = self.current_df.columns[0]
        matches = self.current_df[self.current_df[first_col].astype(str) == search_term]

        if not matches.empty:
            row = matches.iloc[0]
            row_dict = row.to_dict()

            display_dict = {
                k: v for k, v in row_dict.items()
                if k in self.shown_columns and pd.notna(v) and str(v).strip() != "" and not all(c.upper() == 'X' for c in str(v).strip())
            }


            if not hasattr(self, 'open_result_dialogs'):
                self.open_result_dialogs = []

            dialog = ResultDialog(f"Search Result: {search_term}", row_dict, self.shown_columns)
            dialog.setModal(False)
            dialog.show()
# Keep a reference
            self.open_result_dialogs.append(dialog)


        else:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("No Match")
            msg_box.setText(f"No match found for: {search_term}")
            msg_box.setStandardButtons(QMessageBox.Close)
            msg_box.setModal(False)
            msg_box.show()

    def toggle_theme_icon(self):
        match self.current_theme:
            case "light":
                self.current_theme = "pink"
                self.apply_theme("pink")
            case "pink":
                self.current_theme = "dark"
                self.apply_theme("dark")
            case "dark":
                self.current_theme = "light"
                self.apply_theme("light")
        
        
        save_theme_preference(self.current_theme)
    def update_theme_icon(self):
        match self.current_theme:
            case "light":
                self.theme_switch.setText("🌸")
            case "pink":
                self.theme_switch.setText("🌙")
            case "dark":
                self.theme_switch.setText("🔅")

    def apply_theme(self, theme_name):
        if theme_name in self.themes:
            QApplication.instance().setStyleSheet(self.themes[theme_name])
            self.current_theme = theme_name
            save_theme_preference(theme_name)
            self.update_theme_icon()
        else:
            print(f"Theme '{theme_name}' not found.")

    def handle_theme_toggle(self, value):
        theme = "dark" if value == 1 else "light"
        self.apply_theme(theme)
