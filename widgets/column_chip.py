import sys
import os
from PyQt5.QtWidgets import QWidget, QPushButton, QHBoxLayout, QSizePolicy


class ColumnChip(QWidget):
    def __init__(self, column_name, remove_callback):
        super().__init__()
        self.column_name = column_name
        self.remove_callback = remove_callback

        layout = QHBoxLayout()
        layout.setContentsMargins(5, 2, 5, 2)
        layout.setSpacing(5)

        self.remove_btn = QPushButton(str(column_name) +  " ✖")
        self.remove_btn.clicked.connect(self.remove_self)
        self.remove_btn.setProperty("role", "ColumnChip")

        layout.addWidget(self.remove_btn)
        self.setLayout(layout)

        self.setStyleSheet("""
            QPushButton {
                border-radius: 10px;
                padding: 2px 6px;
                min-width: 20px;
                min-height: 20px;
            }

        """)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

    def remove_self(self):
        self.remove_callback(self.column_name)
        self.setParent(None)
