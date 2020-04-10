# -*- coding:utf-8 -*-

import importlib
import json
import os
import sys
import time
from datetime import datetime
from json.decoder import JSONDecodeError

import xlrd
from google.protobuf.internal import api_implementation
from google.protobuf.internal.containers import (
    RepeatedCompositeFieldContainer, RepeatedScalarFieldContainer)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QButtonGroup, QCheckBox,
                             QDesktopWidget, QDialog, QDialogButtonBox,
                             QFileDialog, QHBoxLayout, QLabel, QLineEdit,
                             QMainWindow, QMessageBox, QPushButton,
                             QRadioButton, QSizePolicy, QSpacerItem,
                             QVBoxLayout, QWidget)


class GlobalVariables:
    '''global variables'''
    _res_proto_path = "proto\\res"
    _gen_pb_file_cmd = ".\\proto.bat"

    @staticmethod
    def res_proto_path():
        '''return pb files path'''
        return GlobalVariables._res_proto_path

    @staticmethod
    def gen_pb_file_cmd():
        '''return cmd to generate python file from all pb file in given path'''
        return GlobalVariables._gen_pb_file_cmd


class Xls2DataHandler:
    '''change given excel data to binary by given pb file'''

    def __init__(self, xls_path, sheets_name, pb_file_path, file_range):

        pb_file_name = os.path.splitext(os.path.basename(pb_file_path))[0]
        self._file_range = file_range

        current_dir = os.getcwd()
        os.chdir(GlobalVariables.res_proto_path())
        os.system(GlobalVariables.gen_pb_file_cmd())
        os.chdir(current_dir)
        print("convert proto finish")

        # read xls file
        try:
            self._workbook = xlrd.open_workbook(xls_path)
        except FileNotFoundError as err:
            print("file %s not found", xls_path)
            print("error: ", err)
            raise
        except xlrd.biffh.XLRDError as err:
            print("open workbook error: ", err)
            raise

        # read sheets by given name
        self._sheets = []
        for sheet_name in sheets_name:
            try:
                self._sheets.append(self._workbook.sheet_by_name(sheet_name))
            except xlrd.biffh.XLRDError as err:
                print("read sheet failed, Reason: ", err)
                raise

        # import module
        try:
            self._pb_module = importlib.import_module("proto.res." +
                                                      pb_file_name + '_pb2')
        except ImportError as err:
            print("import module failed, Reason: ", err)
            raise

        self._pb_obj_dic = {}
        self._sheet_name_2_message_name_dict = {}
        # get obj_list
        # convert sheet name to message name by json
        with open("xls2pb.json", encoding="utf-8") as json_file:
            try:
                tmp_dict = json.load(json_file)
            except JSONDecodeError as err:
                print("json file not well formatted ", err)
                raise
            try:
                self._sheet_name_2_message_name_dict = \
                    tmp_dict["SheetNameToMessageName"]
            except KeyError as err:
                print(err)
                raise

        for sheet in self._sheets:
            message_name = ""
            try:
                message_name = self._sheet_name_2_message_name_dict[
                    sheet.name][file_range]
            except KeyError as err:
                continue
                # print(err)
                # raise
            try:
                self._pb_obj_dic[message_name] = getattr(
                    self._pb_module, message_name + "_list")()
            except AttributeError as err:
                print("%s_list not exist in protobuf file" % message_name)
                print("error: ", err)
                raise
        self._log = ''

    def get_message_name(self, sheet):
        try:
            a = self._sheet_name_2_message_name_dict[sheet.name][self._file_range]
        except KeyError as err:
            return None
        return a

    def generate_data_file(self):
        self._parse_sheets()

    def dump(self, dest, dest_log):
        for sheet in self._sheets:
            if self.get_message_name(sheet) == None:
                continue
            data = self._pb_obj_dic[self.get_message_name(
                sheet)].SerializeToString()
            file_name = dest + self.get_message_name(sheet) + ".bytes"
            dest_file = open(file_name, 'wb+')
            dest_file.write(data)
            dest_file.close()
            data = self._pb_obj_dic[self.get_message_name(sheet)]
            log_file = open(
                dest_log + self.get_message_name(sheet) + '.log', 'w+')
            # print(dest + self.get_message_name(sheet))
            # print("log write")
            log_file.write(str(data))
            ignored_col_log = open('ignored_col.log', 'wb')
            ignored_col_log.write(self._log.encode('utf-8'))
            ignored_col_log.close()
            log_file.close()

    def _parse_sheets(self):
        # traverse each row
        for sheet in self._sheets:
            if self.get_message_name(sheet) == None:
                continue

            print('    converting begin : ' + str(sheet.name))
            for current_row_num in range(1, sheet.nrows):
                # add new obj to list
                pb_obj = self._pb_obj_dic[self.get_message_name(
                    sheet)].items_list.add()
                if self._parse_row(current_row_num, pb_obj, sheet):
                    del self._pb_obj_dic[self.get_message_name(sheet)].items_list[
                        len(self._pb_obj_dic[self.get_message_name(sheet)].items_list) - 1]

            print('    converting end : ' + str(sheet.name))
        # print(self._pb_obj_dic)

    def _parse_row(self, row_index, pb_obj, sheet):
        is_null = True
        with open('xls2pb.json', encoding='utf-8') as json_file:
            cell_name_to_pb_name_dict = json.load(
                json_file)["ColNameToStructureName"]
            for current_col_num in range(0, sheet.ncols):
                current_cell = sheet.cell(row_index, current_col_num)
                cell_value = self._get_cell_value(current_cell)
                # print ("converting row[%d], col[%d] : %s %s" % (row_index + 1, current_col_num + 1, str(current_cell), str(cell_value)))
                if cell_value is not None:
                    member_name = sheet.cell(0, current_col_num).value
                    member_name = member_name.replace(" ", "")
                    assert (member_name is
                            not None), "member name should not be NULL"
                    try:
                        member_name = cell_name_to_pb_name_dict[
                            self.get_message_name(sheet)][member_name]
                    except KeyError:
                        self._log = self._log + "ignored col %s\n" % \
                            member_name
                        continue

                    if hasattr(pb_obj, member_name):
                        member = pb_obj.__getattribute__(member_name)
                        assert (isinstance(member,
                                           RepeatedCompositeFieldContainer)is not True), \
                            "member type shuld not be repeated class"
                        if isinstance(member, RepeatedScalarFieldContainer):
                            try:
                                member.append(cell_value)
                                is_null = False
                            except TypeError:
                                # 尝试转字符串
                                try:
                                    member.append(str(cell_value))
                                    is_null = False
                                except TypeError:
                                    # 尝试转数值
                                    if cell_value == u"":
                                        # 数值 - 空表格
                                        pass
                                    else:
                                        try:
                                            member.append(int(cell_value))
                                            is_null = False
                                        except TypeError:
                                            print("[Repeated] UnKnown Exception row[%d], col[%d] : %s %s" % (
                                                row_index + 1, current_col_num + 1, str(current_cell), str(cell_value)))
                                            raise
                        elif isinstance(member, str):
                            pb_obj.__setattr__(member_name, str(cell_value))
                            is_null = False
                        elif isinstance(member, int):
                            if cell_value == u"":
                                # 数值 - 空表格
                                # print ("empty_cell row[%d], col[%d] : %s " % (row_index + 1, current_col_num + 1, str(current_cell)))
                                pass
                            else:
                                try:
                                    pb_obj.__setattr__(
                                        member_name, int(cell_value))
                                    is_null = False
                                except BaseException:
                                    print("Int Exception row[%d], col[%d] : %s  %s" % (
                                        row_index + 1, current_col_num + 1, str(current_cell), str(cell_value)))
                                    raise
                        else:
                            try:
                                pb_obj.__setattr__(member_name, cell_value)
                                is_null = False
                            except TypeError:
                                print("Other Exception row[%d], col[%d] : %s  %s" % (
                                    row_index + 1, current_col_num + 1, str(current_cell), str(cell_value)))
                                raise
        return is_null

    @staticmethod
    def _get_cell_value(cell):
        cell_value = cell.value
        cell_type = cell.ctype
        if cell_type == 0:
            return None

        if cell_type == 1:
            return cell_value

        if cell_type == 2:
            return int(cell_value)

        if cell_type == 3:
            date = datetime(*xlrd.xldate_as_tuple(cell.value, 0))
            cell_value = int(time.mktime(date.timetuple()))
            return cell_value

        if cell_type == 4:
            return bool(cell_value)


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
            handler = Xls2DataHandler(self._xls_file_text.text(),
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


if __name__ == '__main__':
    xls2pbdata = QApplication(sys.argv)

    w = Xls2PBDataGui()

    xls2pbdata.exec_()

    sys.exit()
