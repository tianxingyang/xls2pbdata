import datetime
import time
import traceback

import xlrd
from PyQt5.QtCore import QCoreApplication, Qt
from PyQt5.QtWidgets import (QApplication, QButtonGroup, QCheckBox,
                             QDesktopWidget, QDialog, QDialogButtonBox,
                             QFileDialog, QHBoxLayout, QLabel, QLineEdit,
                             QMainWindow, QMessageBox, QPushButton,
                             QRadioButton, QSizePolicy, QSpacerItem,
                             QVBoxLayout, QWidget)

from package import consts, new_table_handler


class MyPushButton(QPushButton):
    def __init__(self, name, screen):
        super(MyPushButton, self).__init__(name)
        super().setFixedSize(screen.width() / 12, screen.height() / 16)


class MyLabel(QLabel):
    def __init__(self, name, screen):
        super(MyLabel, self).__init__(name)
        super().setFixedSize(screen.width() / 16, screen.height() / 43.2)
        super().setAlignment(Qt.AlignCenter)


class MyLineEdit(QLineEdit):
    def __init__(self, screen):
        super().__init__()
        super().setFixedHeight(screen.height() / 43.2)


class Xls2PBDataGui(QMainWindow):
    def __init__(self):
        super(Xls2PBDataGui, self).__init__()

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 30, 20, 30)
        select_range_box_layout = QHBoxLayout()
        select_xls_file_box_layout = QHBoxLayout()
        button_layout = QHBoxLayout()

        self._main_widget = QWidget()

        self._has_xls_file = False

        self.setWindowTitle('Xls To Bin - vito')

        screen = QDesktopWidget().screenGeometry()

        self._select_xls_file_button = MyPushButton('Select!', screen)
        self._convert_button = MyPushButton("Convert!", screen)

        self._xls_file_label = MyLabel('xls file', screen)

        self._xls_file_text = MyLineEdit(screen)

        self._public_rb = QRadioButton(self)
        self._server_rb = QRadioButton(self)
        self._client_rb = QRadioButton(self)
        self._all_rb = QRadioButton(self)
        self._all_rb.setChecked(True)

        self._public_rb.setText('public')
        self._server_rb.setText('server')
        self._client_rb.setText('client')
        self._all_rb.setText('all')

        self._range_button_box = QButtonGroup()
        self._range_button_box.addButton(self._public_rb, 1)
        self._range_button_box.addButton(self._server_rb, 2)
        self._range_button_box.addButton(self._client_rb, 3)
        self._range_button_box.addButton(self._all_rb, 4)

        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._public_rb)
        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._server_rb)
        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._client_rb)
        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._all_rb)
        select_range_box_layout.addStretch(1)

        main_layout.addLayout(select_range_box_layout)

        select_xls_file_box_layout.addWidget(self._xls_file_label)
        select_xls_file_box_layout.addWidget(self._xls_file_text)
        select_xls_file_box_layout.addSpacing(screen.width() / 30)

        main_layout.addLayout(select_xls_file_box_layout)

        button_layout.addWidget(self._select_xls_file_button)
        button_layout.addWidget(self._convert_button)

        main_layout.addLayout(button_layout)

        self._main_widget.setLayout(main_layout)

        self.setCentralWidget(self._main_widget)

        self._select_xls_file_button.clicked.connect(self.open_xls_file)
        self._convert_button.clicked.connect(self.convert)

        self.setFixedSize(screen.width() / 3, screen.height() / 3)
        self._main_widget.setFixedSize(screen.width() / 3, screen.height() / 3)
        size = self.geometry()
        self.move((screen.width() - size.width()) / 3,
                  (screen.height() - size.height()) / 3)

        self.show()

    def open_xls_file(self):
        file_name = QFileDialog.getOpenFileName(self, '请选择 excel 文件',
                                                consts.DEFAULT_EXCEL_PATH, '*.xlsx')
        if file_name[0]:
            self._has_xls_file = True
            self._xls_file_text.setText(file_name[0])

    def convert(self):
        if self._has_xls_file:
            self._convert_xls_2_bin()
        else:
            self.pop_err_box('请选择 excel 文件')

    @staticmethod
    def pop_err_box(err_msg):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle('ERROR')
        msg_box.setText(err_msg)
        msg_box.exec_()

    def _convert_xls_2_bin(self):
        file_type = ""
        check_id = self._range_button_box.checkedId()
        if check_id == 1:
            file_type = "public"
        elif check_id == 2:
            file_type = "server"
        elif check_id == 3:
            file_type = "client"
        elif check_id == 4:
            file_type = "all"

        try:
            handler = new_table_handler.NewTableHandler(
                self._xls_file_text.text(), file_type)
        except BaseException as err:
            self.pop_err_box(traceback.format_exc())
            return

        try:
            handler.generate_data_file()
        except BaseException as err:
            self.pop_err_box(traceback.format_exc())
            return

        try:
            handler.dump()
        except BaseException as err:
            self.pop_err_box(traceback.format_exc())
            return

        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle('OK')
        msg_box.setText('convert finish')
        msg_box.exec_()
