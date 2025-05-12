import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QDialog, QScrollArea,
    QLabel, QMessageBox
)
from PyQt5.QtCore import Qt

from widgets.copyable_label import CopyableLabel
from widgets.sheet_search_bar import SheetSearchBar
from utilities.sheet_load_utils import load_excel_sheet, clean_column_name, get_first_match, get_row_dict
from utilities.sheet_header_utils import get_header_index, set_header_index
from functools import partial



class ResultDialog(QDialog):
    def __init__(self, search_index, row_dict, shown_columns, excel_files, file_sheets_map, selected_file, selected_sheet):
        super().__init__()
        self.setWindowTitle(f"{selected_sheet}: {search_index}")
        self.setMinimumSize(250, 500)
        self.setMaximumHeight(600)

        self.shown_columns = shown_columns
        self.row_dict = row_dict
        self.excel_files = excel_files
        self.file_sheets_map = file_sheets_map
        self.show_all = False


        self.main_layout = QVBoxLayout()

        # --- Sheet Search Bar (Reusable Component) ---
        self.search_bar = SheetSearchBar()
        self.search_bar.set_files_and_sheets(self.excel_files, self.file_sheets_map, selected_file, selected_sheet)
        self.search_bar.set_index_value(search_index)
        self.search_bar.searchRequested.connect(self.perform_search)
        self.main_layout.addWidget(self.search_bar)

        # --- Result Display Area ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_widget.setLayout(self.content_layout)
        self.scroll_area.setWidget(self.content_widget)
        self.main_layout.addWidget(self.scroll_area)

        # --- Toggle/Close Buttons ---
        button_row = QHBoxLayout()
        self.toggle_btn = QPushButton("Show All Fields")
        self.toggle_btn.clicked.connect(self.toggle_all_fields)
        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)
        button_row.addWidget(self.toggle_btn)
        button_row.addStretch()
        button_row.addWidget(self.close_btn)
        self.main_layout.addLayout(button_row)

        self.setLayout(self.main_layout)
        self.load_saved_header_row()
        self.render_content()

    def perform_search(self):
        selection = self.search_bar.get_selection()
        file = selection["file"]
        sheet = selection["sheet"]
        index_value = selection["index"]
        header_row = selection["header"]
        file_path = self.excel_files.get(file)
        self.row_dict = get_row_dict(file_path, sheet, header_row, index_value)
        self.render_content()

    def render_content(self):
        self.setWindowTitle(f"{self.search_bar.get_selection()['sheet']}: {self.search_bar.get_selection()['index']}")
        self.clear_layout(self.content_layout)
        self.content_layout.setAlignment(Qt.AlignTop)

        for key, value in self.row_dict.items():
            key_str = str(key)
            value_str = str(value)

            row = QHBoxLayout()

            btn_key = QPushButton(key_str)
            btn_key.clicked.connect(partial(self.toggle_column_status, key_str))
            btn_key.setProperty("role", "key")
            

            label_value = CopyableLabel(value_str)
            label_value.setWordWrap(True)
            label_value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            label_value.setProperty("role", "value")

            if self.show_all:
                if key_str not in self.shown_columns:
                    btn_key.setProperty("status", "hidden")
                    
                elif (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    btn_key.setProperty("status", "empty")
                    btn_key.setDisabled(True)
                    label_value.setText("-")
                else:
                    btn_key.setProperty("status", "normal")
            else:
                if key_str not in self.shown_columns:
                    continue
                if (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    continue
                btn_key.setProperty("status", "normal")

            btn_key.style().unpolish(btn_key)
            btn_key.style().polish(btn_key)
            btn_key.setFixedWidth(150)
            btn_key.setMinimumHeight(30)

            label_value.style().unpolish(label_value)
            label_value.style().polish(label_value)
            label_value.setFixedWidth(200)
            label_value.setMinimumHeight(30)

            row.addWidget(btn_key)
            row.addWidget(label_value)
            self.content_layout.addLayout(row)

    def toggle_column_status(self, col):
        if col in self.shown_columns:
            self.shown_columns.remove(col)
        else:
            self.shown_columns.append(col)
        self.render_content()

    def toggle_all_fields(self):
        self.show_all = not self.show_all
        self.toggle_btn.setText("Show Only Selected" if self.show_all else "Show All Fields")
        self.render_content()

    def clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
            else:
                sub_layout = item.layout()
                if sub_layout is not None:
                    self.clear_layout(sub_layout)
    def load_saved_header_row(self):
        selection = self.search_bar.get_selection()
        file = selection["file"]
        sheet = selection["sheet"]
        saved_row = get_header_index(file, sheet)

        if saved_row:
            self.search_bar.set_header_row(saved_row)

