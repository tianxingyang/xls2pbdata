import datetime
import importlib
import json
import os
import sys
import time

import xlrd
from google.protobuf.internal import api_implementation
from google.protobuf.internal.containers import (
    RepeatedCompositeFieldContainer, RepeatedScalarFieldContainer)

from package.consts import *
from package.logging_wrapper import *


class Error(Exception):
    pass


class FieldNameEmptyError(Error):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


class SheetNameNotFoundInProtoError(Error):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


class NewTableHandler:
    def __init__(self, xls_path, file_range):
        super().__init__()

        # 转化 proto 文件为 py 文件
        current_dir = os.getcwd()
        os.chdir(RES_PROTO_PATH)
        os.system(GEN_PB_CMD)
        os.chdir(current_dir)
        log_debug("convert proto finish")

        # read xls file
        try:
            self._workbook = xlrd.open_workbook(xls_path)
        except FileNotFoundError as err:
            raise
        except xlrd.biffh.XLRDError as err:
            raise

        self._client_pb_obj = None

        # 导入 client public server
        try:
            sys.path.append(RES_PROTO_PATH)
            self._client_pb_module = importlib.import_module("client_pb2")
            self._server_pb_module = importlib.import_module("server_pb2")
            self._public_pb_module = importlib.import_module("public_pb2")
        except ImportError as err:
            raise

        # 其实只有一个 sheet 是需要转的，如果有多的 sheet，一定是注释
        self._sheets = []
        for sheet in self._workbook.sheets():
            log_debug("start converting %s" % sheet.name)
            # 判断 sheet 命名是否出错
            self._client_class_type = self._get_attr(
                self._client_pb_module, sheet.name)
            self._public_class_type = self._get_attr(
                self._public_pb_module, sheet.name)
            self._server_class_type = self._get_attr(
                self._server_pb_module, sheet.name)
            if not self._client_class_type and not self._public_class_type and not self._server_class_type:
                log_info("跳过 sheet[%s]" % sheet.name)
                continue

            self._client_pb_obj = None
            self._public_pb_obj = None
            self._server_pb_obj = None
            # 如果这个 sheet 存在于 client proto 中，则生成这个对象
            if self._client_class_type:
                self._client_pb_obj = self._client_class_type()
            # 如果这个 sheet 存在于 public proto 中，则生成这个对象
            if self._public_class_type:
                self._public_pb_obj = self._public_class_type()
            # 如果这个 sheet 存在于 server proto 中，则生成这个对象
            if self._server_class_type:
                self._server_pb_obj = self._server_class_type()
            self._sheets.append(sheet)

    def generate_data_file(self):
        for sheet in self._sheets:
            log_debug("start converting %s" % sheet.name)

            # 读取每一行，从第三行开始，第一行是汉字名字，方便策划配置，第二行是 protobuf 中 field 的名字
            for row in range(2, sheet.nrows):
                client_row_obj = None
                public_row_obj = None
                server_row_obj = None
                cell_list = sheet.row(row)
                if self._client_class_type:
                    client_row_obj = self._client_class_type.M()
                if self._public_class_type:
                    public_row_obj = self._public_class_type.M()
                if self._server_class_type:
                    server_row_obj = self._server_class_type.M()

                for idx, val in enumerate(cell_list):
                    # 获取该 cell 对应的 field 名字
                    field_name = sheet.cell(1, idx).value
                    field_name = field_name.replace(" ", "")
                    if field_name is None:
                        raise FieldNameEmptyError
                    if client_row_obj:
                        if hasattr(self._client_class_type.M, field_name):
                            # 如果该 sheet  和 cell 都在 client 的配置中，则赋值
                            self._assign_by_cell(
                                val, field_name, client_row_obj)
                    if public_row_obj:
                        if hasattr(self._public_class_type.M, field_name):
                            # 如果该 sheet  和 cell 都在 client 的配置中，则赋值
                            self._assign_by_cell(
                                val, field_name, public_row_obj)
                    if server_row_obj:
                        if hasattr(self._server_class_type.M, field_name):
                            # 如果该 sheet  和 cell 都在 client 的配置中，则赋值
                            self._assign_by_cell(
                                val, field_name, server_row_obj)
                if client_row_obj is not None:
                    self._client_pb_obj.items_list.append(client_row_obj)
                if public_row_obj is not None:
                    self._public_pb_obj.items_list.append(public_row_obj)
                if server_row_obj is not None:
                    self._server_pb_obj.items_list.append(server_row_obj)

    def _get_attr(self, module, name):
        if hasattr(module, name):
            return getattr(module, name)
        else:
            return None

    def _get_field(self, obj, name):
        try:
            field = getattr(obj(), name)
        except AttributeError as err:
            return None
        return field

    def _assign_by_cell(self, cell, field_name, row_obj):
        cell_value = self._get_cell_value(cell)
        if cell_value is None:
            return

        # 如果 field 是 repeated 类型
        field = getattr(row_obj, field_name)
        if isinstance(field, RepeatedScalarFieldContainer):
            field.append(cell_value)
        # 如果是其他类型，直接赋值
        else:
            row_obj.__setattr__(field_name, cell_value)

    def dump(self):
        for sheet in self._sheets:
            if self._client_pb_obj:
                client_data = self._client_pb_obj.SerializeToString()
                with open(CLIENT_DATA_PATH + sheet.name + ".bytes", "wb+") as f:
                    f.write(client_data)
            if self._server_pb_obj:
                server_data = self._server_pb_obj.SerializeToString()
                with open(SERVER_DATA_PATH + sheet.name + ".bytes", "wb+") as f:
                    f.write(server_data)
            if self._public_pb_obj:
                public_data = self._public_pb_obj.SerializeToString()
                with open(PUBLIC_DATA_PATH + sheet.name + ".bytes", "wb+") as f:
                    f.write(public_data)

    @staticmethod
    def _get_cell_value(cell):
        cell_value = cell.value
        cell_type = cell.ctype
        if cell_type == xlrd.XL_CELL_EMPTY:
            return None

        if cell_type == xlrd.XL_CELL_TEXT:
            if cell_value == "":
                return None
            return str(cell_value)

        if cell_type == xlrd.XL_CELL_NUMBER:
            if cell_value == u"":
                return None
            return int(cell_value)

        if cell_type == 3:
            date = datetime(*xlrd.xldate_as_tuple(cell.value, 0))
            cell_value = int(time.mktime(date.timetuple()))
            return cell_value

        if cell_type == 4:
            return bool(cell_value)
