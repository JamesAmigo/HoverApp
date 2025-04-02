import sys
import os
from PyQt5.QtWidgets import QHBoxLayout, QLabel
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QMovie


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