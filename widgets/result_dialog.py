import sys
import os
import pandas as pd
from widgets.copyable_label import CopyableLabel
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QDialog, QScrollArea,
    QComboBox, QLabel, QLineEdit
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIntValidator
class ResultDialog(QDialog):
    def __init__(self, title, row_dict, shown_columns, excel_files, file_sheets_map):
        super().__init__()
        self.setWindowTitle(title)
        self.setMinimumSize(250, 500)
        self.setMaximumHeight(600)

        self.excel_files = excel_files
        self.file_sheets_map = file_sheets_map
        self.shown_columns = shown_columns
        self.row_dict = row_dict
        self.show_all = False
        self.current_df = None
        self.current_header_row = 2


        self.main_layout = QVBoxLayout()

        # --- File/Sheet/Header Row ---
        file_row = QHBoxLayout()
        self.file_dropdown = QComboBox()
        self.file_dropdown.addItems(list(excel_files.keys()))
        self.file_dropdown.currentIndexChanged.connect(self.on_file_selected)

        self.sheet_dropdown = QComboBox()
        first_file = list(excel_files.keys())[0]
        self.sheet_dropdown.addItems(file_sheets_map[first_file])
        self.sheet_dropdown.currentIndexChanged.connect(self.on_sheet_selected)

        self.header_input = QLineEdit("2")
        self.header_input.setFixedWidth(40)
        self.header_input.setValidator(QIntValidator(1, 999))
        
        file_row.addWidget(QLabel("File:"))
        file_row.addWidget(self.file_dropdown)
        file_row.addWidget(QLabel("Sheet:"))
        file_row.addWidget(self.sheet_dropdown)
        file_row.addWidget(QLabel("Header:"))
        file_row.addWidget(self.header_input)
        self.main_layout.addLayout(file_row)

        # --- Search Row ---
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter index")
        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self.perform_search)
        search_row.addWidget(self.search_input)
        search_row.addWidget(self.search_btn)
        self.main_layout.addLayout(search_row)

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
        self.render_content()
    def on_file_selected(self, index):
        file = self.file_dropdown.currentText()
        self.sheet_dropdown.clear()
        self.sheet_dropdown.addItems(self.file_sheets_map[file])

    def on_sheet_selected(self, index):
        pass  # Optional: trigger default search

    def perform_search(self):
        file = self.file_dropdown.currentText()
        sheet = self.sheet_dropdown.currentText()
        header = int(self.header_input.text() or 1)
        filepath = self.excel_files[file]

        try:
            df = pd.read_excel(filepath, sheet_name=sheet, header=header - 1, engine='openpyxl')
            df.columns = df.columns.astype(str)
            df = df.loc[:, ~df.columns.str.contains(r'^Unnamed', case=False)]
            df = df.dropna(axis=1, how='all')
            self.current_df = df

            search_term = self.search_input.text().strip()
            if search_term:
                match = df[df[df.columns[0]].astype(str) == search_term]
                if not match.empty:
                    self.row_dict = match.iloc[0].to_dict()
        except Exception as e:
            self.row_dict = {"Error": str(e)}

        self.render_content()
    def render_content(self):
        # Clear previous layout
        self.clear_layout(self.content_layout)
        self.content_layout.setAlignment(Qt.AlignTop)


        for key, value in self.row_dict.items():
            key_str = str(key)
            value_str = str(value)

            row = QHBoxLayout()
            label_key = CopyableLabel(key_str)

            label_key.setTextInteractionFlags(Qt.TextSelectableByMouse)
            label_value = CopyableLabel(value_str)
            label_value.setWordWrap(True)
            label_value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            label_key.setProperty("role", "key")
            label_value.setProperty("role", "value")

            # Customize status
            if self.show_all:
                if key_str not in self.shown_columns:
                    label_key.setProperty("status", "hidden")
                elif (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    label_key.setProperty("status", "empty")
                    label_value.setText("-")
                else:
                    label_key.setProperty("status", "normal")
            else:
                if key_str not in self.shown_columns:
                    continue
                if (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    continue
                label_key.setProperty("status", "normal")


            label_key.style().unpolish(label_key)
            label_key.style().polish(label_key)

            label_key.setFixedWidth(150)
            label_key.setMinimumHeight(30)
            label_value.setMinimumHeight(30)
            row.addWidget(label_key)
            row.addWidget(label_value)
            self.content_layout.addLayout(row)
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
                # it's likely a nested layout
                sub_layout = item.layout()
                if sub_layout is not None:
                    self.clear_layout(sub_layout)

    def on_file_changed(self, filename):
        self.selected_file = filename
        self.sheet_dropdown.clear()
        self.sheet_dropdown.addItems(self.file_sheets_map[filename])
        self.selected_sheet = self.file_sheets_map[filename][0]

    def on_sheet_changed(self, sheet_name):
        self.selected_sheet = sheet_name

