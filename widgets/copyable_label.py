import sys
import os
from PyQt5.QtWidgets import QApplication, QLabel
from PyQt5.QtCore import Qt, QPoint, QTimer
from PyQt5.QtGui import QCursor

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
      
