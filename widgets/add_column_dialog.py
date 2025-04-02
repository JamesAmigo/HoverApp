import sys
import os
from PyQt5.QtWidgets import QPushButton, QVBoxLayout, QDialog, QListWidget, QListWidgetItem

class AddColumnDialog(QDialog):
    def __init__(self, available_columns, add_callback):
        super().__init__()
        self.setWindowTitle("Add Column")
        self.setFixedSize(300, 400)
        self.add_callback = add_callback

        layout = QVBoxLayout()
        self.list_widget = QListWidget()

        for col in available_columns:
            item = QListWidgetItem(col)
            self.list_widget.addItem(item)

        self.btn_add = QPushButton("Add Selected")
        self.btn_add.clicked.connect(self.add_selected)

        layout.addWidget(self.list_widget)
        layout.addWidget(self.btn_add)
        self.setLayout(layout)

    def add_selected(self):
        selected_items = self.list_widget.selectedItems()
        for item in selected_items:
            self.add_callback(item.text())
        self.accept()
