import sys
import os
import pandas as pd
from widgets.copyable_label import CopyableLabel
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QDialog, QScrollArea
)
from PyQt5.QtCore import Qt
class ResultDialog(QDialog):
    def __init__(self, title, row_dict, shown_columns):
        super().__init__()
        self.setWindowTitle(title)
        self.setMinimumHeight(250)
        self.setMinimumWidth(300)
        self.setMaximumHeight(600)

        self.row_dict = row_dict
        self.shown_columns = shown_columns
        self.show_all = False

        self.main_layout = QVBoxLayout()
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_widget.setLayout(self.content_layout)
        self.scroll_area.setWidget(self.content_widget)

        self.main_layout.addWidget(self.scroll_area)

        # Buttons
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
            label_key.setFixedHeight(30)
            label_value.setFixedHeight(30)
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
