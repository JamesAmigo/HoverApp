from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
import sys
from hover_app import ExcelFolderApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setFont(QFont("微軟正黑體", 10))
    window = ExcelFolderApp()
    window.show()
    sys.exit(app.exec_())
