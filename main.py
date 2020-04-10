import sys
from PyQt5.QtWidgets import QApplication

from package.pyqt5_gui import Xls2PBDataGui

if __name__ == "__main__":
    xls2pbdata = QApplication(sys.argv)

    w = Xls2PBDataGui()

    xls2pbdata.exec_()

    sys.exit()
