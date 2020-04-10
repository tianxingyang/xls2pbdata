import sys

from PyQt5.QtWidgets import QApplication

from package.pyqt5_gui import Xls2PBDataGui

if __name__ == "__main__":
    app = QApplication(sys.argv)

    w = Xls2PBDataGui()

    app.exec_()

    sys.exit()
