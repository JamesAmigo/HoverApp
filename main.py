from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
import sys
from utilities.resource_utils import get_resource_path
from ui_main import ExcelFolderApp
from PyQt5.QtGui import QIcon

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(get_resource_path(r"Resources/icon.ico")))
    app.setFont(QFont("微軟正黑體", 10))
    window = ExcelFolderApp()
    window.show()
    sys.exit(app.exec_())
