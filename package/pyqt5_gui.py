import datetime
import time

import xlrd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QButtonGroup, QCheckBox,
                             QDesktopWidget, QDialog, QDialogButtonBox,
                             QFileDialog, QHBoxLayout, QLabel, QLineEdit,
                             QMainWindow, QMessageBox, QPushButton,
                             QRadioButton, QSizePolicy, QSpacerItem,
                             QVBoxLayout, QWidget)

from .table_handler import TableHandler


class Xls2PBDataGui(QMainWindow):
    def __init__(self):
        super(Xls2PBDataGui, self).__init__()

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 30, 20, 30)
        select_range_box_layout = QHBoxLayout()
        select_range_box_layout.setContentsMargins(10, 0, 10, 0)
        select_xls_file_box_layout = QHBoxLayout()
        select_pb_file_box_layout = QHBoxLayout()
        select_sheet_box_layout = QHBoxLayout()
        conver_button_layout = QVBoxLayout()

        self._main_widget = QWidget()

        self._has_pb_file = False
        self._has_xls_file = False
        self._selected_sheets_name = []

        self.setWindowTitle('Xls To Bin - vito')

        self.statusBar().showMessage('Wait selecting files')

        self._select_pb_file_button = QPushButton('select pb file')
        self._select_pb_file_button.setFixedSize(120, 25)
        self._select_xls_file_button = QPushButton('select xls file')
        self._select_xls_file_button.setFixedSize(120, 25)
        self._select_sheet_button = QPushButton('select sheet')
        self._select_sheet_button.setFixedSize(120, 25)
        self._convert_button = QPushButton("Convert xls to bin")
        self._convert_button.setFixedSize(120, 25)

        self._pb_file_label = QLabel('protobuf file')
        self._pb_file_label.setFixedSize(80, 25)
        self._pb_file_label.setAlignment(Qt.AlignCenter)
        self._xls_file_label = QLabel('xls file')
        self._xls_file_label.setFixedSize(80, 25)
        self._xls_file_label.setAlignment(Qt.AlignCenter)
        self._xls_sheet_label = QLabel('xls sheet')
        self._xls_sheet_label.setFixedSize(80, 25)
        self._xls_sheet_label.setAlignment(Qt.AlignCenter)

        self._pb_file_text = QLineEdit(self)
        self._xls_file_text = QLineEdit(self)
        self._xls_sheet_text = QLineEdit(self)

        self._public_rb = QRadioButton(self)
        self._server_rb = QRadioButton(self)
        self._client_rb = QRadioButton(self)

        self._public_rb.setText('public')
        self._server_rb.setText('server')
        self._client_rb.setText('client')

        self._range_button_box = QButtonGroup()
        self._range_button_box.addButton(self._public_rb, 1)
        self._range_button_box.addButton(self._server_rb, 2)
        self._range_button_box.addButton(self._client_rb, 3)

        select_range_box_layout.addWidget(self._public_rb)
        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._server_rb)
        select_range_box_layout.addStretch(1)
        select_range_box_layout.addWidget(self._client_rb)
        select_range_box_layout.addStretch(8)

        main_layout.addLayout(select_range_box_layout)

        select_xls_file_box_layout.addWidget(self._xls_file_label)
        select_xls_file_box_layout.addWidget(self._xls_file_text)
        select_xls_file_box_layout.addWidget(self._select_xls_file_button)

        main_layout.addLayout(select_xls_file_box_layout)

        select_sheet_box_layout.addWidget(self._xls_sheet_label)
        select_sheet_box_layout.addWidget(self._xls_sheet_text)
        select_sheet_box_layout.addWidget(self._select_sheet_button)

        main_layout.addLayout(select_sheet_box_layout)

        select_pb_file_box_layout.addWidget(self._pb_file_label)
        select_pb_file_box_layout.addWidget(self._pb_file_text)
        select_pb_file_box_layout.addWidget(self._select_pb_file_button)

        main_layout.addLayout(select_pb_file_box_layout)

        conver_button_layout.addWidget(self._convert_button)

        main_layout.addLayout(conver_button_layout)

        self._main_widget.setLayout(main_layout)

        self.setCentralWidget(self._main_widget)

        self._select_pb_file_button.clicked.connect(self.open_proto_file)
        self._select_xls_file_button.clicked.connect(self.open_xls_file)
        self._convert_button.clicked.connect(self.convert)
        self._select_sheet_button.clicked.connect(self.select_xls_sheet)

        screen = QDesktopWidget().screenGeometry()
        self.setFixedSize(640, 360)
        self._main_widget.setFixedSize(640, 360)
        size = self.geometry()
        self.move((screen.width() - size.width()) / 3,
                  (screen.height() - size.height()) / 3)

        self.show()

    def open_proto_file(self):
        file_name = QFileDialog.getOpenFileName(self, 'open proto file',
                                                './proto/res', '*.proto')
        if file_name[0]:
            self._has_pb_file = True
            self._pb_file_text.setText(file_name[0])

    def open_xls_file(self):
        file_name = QFileDialog.getOpenFileName(self, 'open proto file',
                                                './config', '*.xlsx')
        if file_name[0]:
            self._has_xls_file = True
            self._xls_file_text.setText(file_name[0])

    def select_xls_sheet(self):
        if not self._has_xls_file:
            self.pop_err_box('select xls file first')
            return
        try:
            dialog = SelectSheetDialog(self._xls_file_text.text())
        except BaseException as err:
            self.pop_err_box(err.__str__())
            raise

        if dialog.exec_():
            self._selected_sheets_name = dialog.get_selected_checkboxes()
            self._xls_sheet_text.setText(self._selected_sheets_name.__str__())

    def convert(self):
        if self._has_pb_file and self._has_xls_file:
            self._convert_xls_2_bin()
        elif self._has_xls_file is False or self._has_pb_file is False:
            self.pop_err_box('both xls and protobuf file should be selected')

    @staticmethod
    def pop_err_box(err_msg):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle('ERROR')
        msg_box.setText(err_msg)
        msg_box.exec_()

    def selected_sheets_name(self):
        return self._selected_sheets_name

    # TODO convert bytes to the path same as xls path in ResData
    def _convert_xls_2_bin(self):
        file_type = ""
        check_id = self._range_button_box.checkedId()
        if check_id == 1:
            file_type = "public"
        elif check_id == 2:
            file_type = "server"
        elif check_id == 3:
            file_type = "client"

        try:
            handler = TableHandler(self._xls_file_text.text(),
                                      self._selected_sheets_name,
                                      self._pb_file_text.text(), file_type)
        except BaseException as err:
            print(err)
            self.pop_err_box(err.__str__())
            return

        try:
            handler.generate_data_file()
        except BaseException as err:
            self.pop_err_box(err.__str__())
            return

        dest = '.\\gamedata\\'
        dest_log = dest
        if check_id == 1:
            dest_log = dest + 'public\\'
            dest = dest + 'public\\'
        elif check_id == 2:
            dest_log = dest + 'server\\'
            dest = dest + 'server\\'
        elif check_id == 3:
            dest_log = dest + 'client_log\\'
            dest = dest + 'client\\'

        try:
            handler.dump(dest, dest_log)
        except BaseException as err:
            self.pop_err_box(err.__str__())
            return

        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle('OK')
        msg_box.setText('convert finish')
        msg_box.exec_()


class SelectSheetDialog(QDialog):
    def __init__(self, xls_file_path):
        super(SelectSheetDialog, self).__init__()

        self._workbook = xlrd.open_workbook(xls_file_path)
        self._sheets = []
        self._checkboxes = []
        self._checkbox_check_all = QCheckBox('check all', self)
        self._checkbox_check_all.stateChanged.connect(self.check_all)

        for sheet_index in range(0, self._workbook.nsheets):
            sheet = self._workbook.sheet_by_index(sheet_index).name
            self._sheets.append(sheet)
            self._checkboxes.append(QCheckBox(sheet, self))

        self.setWindowTitle('select sheet')
        self.resize(240, (self._workbook.nsheets + 1) * 50)

        checkboxes_layout = QVBoxLayout()
        checkboxes_layout.addWidget(self._checkbox_check_all)
        for checkbox in self._checkboxes:
            checkbox.stateChanged.connect(self.single_check)
            checkboxes_layout.addWidget(checkbox)

        dialog_layout = QVBoxLayout()
        dialog_layout.addLayout(checkboxes_layout)

        # 创建ButtonBox，用户确定和取消
        self._button_box = QDialogButtonBox()
        self._button_box.setOrientation(Qt.Horizontal)  # 设置为水平方向
        self._button_box.setStandardButtons(
            QDialogButtonBox.Cancel | QDialogButtonBox.Ok)  # 确定和取消两个按钮
        # 连接信号和槽
        self._button_box.accepted.connect(self.accept)  # 确定
        self._button_box.rejected.connect(self.reject)  # 取消
        spacer_item = QSpacerItem(20, 48, QSizePolicy.Minimum,
                                  QSizePolicy.Expanding)
        dialog_layout.addItem(spacer_item)
        dialog_layout.addWidget(self._button_box)
        self.setLayout(dialog_layout)

    def check_all(self):
        if self._checkbox_check_all.checkState() == Qt.Checked:
            for checkbox in self._checkboxes:
                checkbox.setChecked(True)
        elif self._checkbox_check_all.checkState() == Qt.Unchecked:
            for checkbox in self._checkboxes:
                checkbox.setChecked(False)

    def single_check(self):
        all_checked = True
        one_checked = False
        for checkbox in self._checkboxes:
            if not checkbox.isChecked():
                all_checked = False
                break

        for checkbox in self._checkboxes:
            if checkbox.isChecked():
                one_checked = True

        if all_checked:
            self._checkbox_check_all.setCheckState(Qt.Checked)
        elif one_checked:
            self._checkbox_check_all.setTristate()
            self._checkbox_check_all.setCheckState(Qt.PartiallyChecked)
        else:
            self._checkbox_check_all.setTristate(False)
            self._checkbox_check_all.setCheckState(Qt.Unchecked)

    def get_selected_checkboxes(self):
        selected_checkboxes = []
        for checkbox in self._checkboxes:
            if checkbox.isChecked():
                selected_checkboxes.append(checkbox.text())

        return selected_checkboxes


def get_cell_value(cell):
    cell_type = cell.ctype
    if cell_type == 0:
        return None
    elif cell_type == 2:
        return cell.value
    elif cell_type == 3:
        date = datetime(*xlrd.xldate_as_tuple(cell.value, 0))
        cell_value = int(time.mktime(date.timetuple()))
        return cell_value
