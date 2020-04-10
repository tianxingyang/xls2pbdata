import datetime
import importlib
import json
import os
import time

import xlrd
from google.protobuf.internal import api_implementation
from google.protobuf.internal.containers import (
    RepeatedCompositeFieldContainer, RepeatedScalarFieldContainer)

from . import consts, logging_wrapper


class TableHandler:
    def __init__(self, xls_path, sheets_name, pb_file_path, file_range):

        pb_file_name = os.path.splitext(os.path.basename(pb_file_path))[0]
        self._file_range = file_range

        current_dir = os.getcwd()
        os.chdir(consts.RES_PROTO_PATH)
        os.system(consts.GEN_PB_CMD)
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
            except json.JSONDecodeError as err:
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
