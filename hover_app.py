import sys
import os
import pandas as pd
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QComboBox, QLineEdit, QMessageBox,
    QDialog, QListWidget, QListWidgetItem, QLayout, QSizePolicy,QScrollArea
)
from PyQt5.QtCore import Qt, QSize, QRect, QPoint, QTimer
from PyQt5.QtGui import QFontMetrics, QFont, QMovie, QCursor, QClipboard


class LoadingSpinner(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)

        self.label = QLabel("Loading...")
        self.label.setStyleSheet("font-size: 10pt; color: gray;")
        self.label.setVisible(False)

        self.spinner = QLabel()
        self.spinner.setAlignment(Qt.AlignCenter)
        self.spinner.setStyleSheet("background: transparent;")
        self.spinner.setVisible(False)

        self.movie = QMovie(os.path.join(os.path.dirname(__file__), "spinner.gif"))
        self.movie.setScaledSize(QSize(24, 24))
        self.spinner.setMovie(self.movie)

        layout.addWidget(self.spinner)
        layout.addWidget(self.label)

        self.setVisible(False)


    def start(self):
        self.setVisible(True)
        self.label.setVisible(True)
        self.spinner.setVisible(True)
        self.movie.start()

    def stop(self):
        self.movie.stop()
        self.spinner.setVisible(False)
        self.label.setVisible(False)
        self.setVisible(False)
class CopyableLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.setCursor(Qt.IBeamCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            QApplication.clipboard().setText(self.text())
            self.show_copied_popup()
        super().mousePressEvent(event)

    def show_copied_popup(self):
        popup = QLabel("Copied!", self.window())
        popup.setStyleSheet(
            "QLabel {"
            "background-color: rgba(0, 0, 0, 180);"
            "color: white;"
            "padding: 4px 10px;"
            "border-radius: 6px;"
            "font-size: 9pt;"
            "}"
        )
        popup.setWindowFlags(Qt.ToolTip)
        popup.adjustSize()

        # Position near the cursor
        global_pos = QCursor.pos()
        local_pos = self.window().mapFromGlobal(global_pos)
        popup.move(self.window().mapToGlobal(local_pos + QPoint(10, 10)))
        popup.show()

        QTimer.singleShot(1000, popup.close)  # Auto-hide after 1 sec
      

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

        layout.addWidget(self.remove_btn)
        self.setLayout(layout)

        self.setStyleSheet("""
            QWidget {
                border: 1px solid gray;
                border-radius: 10px;
                padding: 2px 4px;
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #e0e0e0;
                border: 1px solid #999;
                border-radius: 10px;
                padding: 2px 6px;
                min-width: 16px;
                min-height: 16px;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
                border-color: #666;
            }
            QPushButton:pressed {
                background-color: #bbbbbb;
            }

        """)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

    def remove_self(self):
        self.remove_callback(self.column_name)
        self.setParent(None)

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

class FlowLayout(QLayout):
    def __init__(self, parent=None, margin=0, spacing=5):
        super(FlowLayout, self).__init__(parent)
        self.setContentsMargins(margin, margin, margin, margin)
        self.setSpacing(spacing)
        self.itemList = []

    def addItem(self, item):
        self.itemList.append(item)

    def count(self):
        return len(self.itemList)

    def itemAt(self, index):
        return self.itemList[index] if index < len(self.itemList) else None

    def takeAt(self, index):
        return self.itemList.pop(index) if index < len(self.itemList) else None

    def expandingDirections(self):
        return Qt.Orientations(Qt.Orientation(0))

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self.doLayout(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect):
        super(FlowLayout, self).setGeometry(rect)
        self.doLayout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self.itemList:
            size = size.expandedTo(item.minimumSize())
        size += QSize(2 * self.contentsMargins().top(), 2 * self.contentsMargins().top())
        return size

    def doLayout(self, rect, testOnly):
        x = rect.x()
        y = rect.y()
        lineHeight = 0

        for item in self.itemList:
            wid = item.widget()
            spaceX = self.spacing()
            spaceY = self.spacing()
            nextX = x + item.sizeHint().width() + spaceX
            if nextX - spaceX > rect.right() and lineHeight > 0:
                x = rect.x()
                y = y + lineHeight + spaceY
                nextX = x + item.sizeHint().width() + spaceX
                lineHeight = 0

            if not testOnly:
                item.setGeometry(QRect(QPoint(x, y), item.sizeHint()))

            x = nextX
            lineHeight = max(lineHeight, item.sizeHint().height())

        return y + lineHeight - rect.y()

class ExcelFolderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_files = {}
        self.current_df = None
        self.shown_columns = []

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Excel Folder Search Tool')
        self.setMinimumWidth(700)

        layout = QVBoxLayout()

        # Folder Row (Label + Button + File/Sheet Layout)
        folder_row = QHBoxLayout()

        self.label_info = QLabel('No folder selected')
        self.label_info.setMinimumWidth(100)
        self.label_info.setMaximumWidth(200)
        folder_row.addWidget(self.label_info)

        self.btn_choose_folder = QPushButton('Choose Folder')
        self.btn_choose_folder.setMinimumSize(100, 30)
        self.btn_choose_folder.setMaximumSize(150, 60)
        self.btn_choose_folder.clicked.connect(self.choose_folder)
        folder_row.addWidget(self.btn_choose_folder)

        # File/Sheet nested layout
        file_sheet_grid = QVBoxLayout()

        top_row = QHBoxLayout()
        bottom_row = QHBoxLayout()

        file_label = QLabel("File name:")
        file_label.setFixedSize(70, 30)
        self.file_dropdown = QComboBox()
        self.file_dropdown.setEnabled(False)
        self.file_dropdown.currentIndexChanged.connect(self.on_file_selected)

        sheet_label = QLabel("Sheet name:")
        sheet_label.setFixedSize(70, 30)
        self.sheet_dropdown = QComboBox()
        self.sheet_dropdown.setEnabled(False)
        self.sheet_dropdown.currentIndexChanged.connect(self.on_sheet_selected)

        top_row.addWidget(file_label)
        top_row.addWidget(self.file_dropdown)

        bottom_row.addWidget(sheet_label)
        bottom_row.addWidget(self.sheet_dropdown)

        file_sheet_grid.addLayout(top_row)
        file_sheet_grid.addLayout(bottom_row)

        folder_row.addLayout(file_sheet_grid)
        layout.addLayout(folder_row)

        # Spinner
        self.spinner = LoadingSpinner(self)
        layout.addWidget(self.spinner)

        # Search row
        search_layout = QHBoxLayout()
        self.input_search = QLineEdit()
        self.input_search.setPlaceholderText("Enter index to search")
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.perform_search)
        search_layout.addWidget(self.input_search)
        search_layout.addWidget(self.btn_search)

        # Column chips
        self.chip_container_widget = QWidget()
        self.chip_layout = FlowLayout()
        self.chip_container_widget.setLayout(self.chip_layout)

        self.btn_add_column = QPushButton("Add Column")
        self.btn_add_column.clicked.connect(self.show_add_column_dialog)

        layout.addLayout(search_layout)
        layout.addWidget(self.chip_container_widget)
        layout.addWidget(self.btn_add_column)

        self.setLayout(layout)

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if not folder:
            return

        self.label_info.setText(f'Selected folder: {folder}')
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
class ResultDialog(QDialog):
    def __init__(self, title, row_dict, shown_columns):
        super().__init__()
        self.setWindowTitle(title)
        self.setFixedHeight(600)
        self.setMinimumWidth(500)

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


        for key, value in self.row_dict.items():
            key_str = str(key)
            value_str = str(value)

            row = QHBoxLayout()
            label_key = CopyableLabel(key_str)

            label_key.setTextInteractionFlags(Qt.TextSelectableByMouse)
            label_value = CopyableLabel(value_str)
            label_value.setWordWrap(True)
            label_value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            
            key_styles = [
                "font-weight: bold",
                "font-size: 10pt",
                "border: 1px solid #ccc",
                "padding: 4px",
                "border-radius: 6px",
            ]

            value_styles = [
                "font-size: 8pt",
                "padding: 4px"
            ]

            if self.show_all:
                if key_str not in self.shown_columns:
                    key_styles.append("background-color: #e0e0e0")
                elif (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    key_styles.append("color: white; background-color: #868686")
                    label_value.setText("-")
                else:
                    key_styles.append("background-color: #f2f2f2")

            else:
                if key_str not in self.shown_columns:
                    continue
                if (not pd.notna(value)) or value_str.strip() == "" or all(c.upper() == 'X' for c in value_str.strip()):
                    continue

            label_key.setStyleSheet("; ".join(key_styles))
            label_value.setStyleSheet("; ".join(value_styles))

            label_key.setFixedWidth(150)
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
