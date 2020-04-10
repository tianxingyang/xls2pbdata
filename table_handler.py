import importlib
import json
import os

import xlrd

import consts
import logging_wrapper


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
