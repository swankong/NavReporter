# -*- coding:utf-8 -*-

#
# NavParser.py
#
# Parse the Excel format data files and
#
# extract useful information.
#
# All rights reserved, 2019
#

import os
import re
import threading
import xlrd

import datetime as dt

from Calendar import Calendar

from xlrd.book import Book
from xlrd.sheet import Sheet
from xlrd.sheet import Cell


class NavNode(object):

    def __init__(self):
        self.data_file = None
        self.unit_nav = None
        self.accum_ret = None
        self.ttl_nav = None


class NavParser(object):

    """ Make a singleton class """
    _instance_lock = threading.Lock()
    @classmethod
    def instance(cls):
        """ Singleton pattern """
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:
                if not hasattr(cls, '_instance'):
                    cls._instance = NavParser()
        return cls._instance

    def __init__(self):
        self.dt_date = None
        self.data_files = None
        self.date_patterns = None
        self.input_path = None
        self.nav_map = None
        self.unit_nav_items = None
        self.accum_ret_items = None
        self.ttl_nav_items = None
        self.prod_names = None
        self.res_dict = dict()

    def set_conf_params(self, mod_config):
        self.dt_date = Calendar.instance().get_prev_trading_day()
        self.data_files = mod_config.data_files
        self.prod_names = mod_config.prod_names
        self.date_patterns = mod_config.date_patterns
        self.input_path = mod_config.input_path
        self.unit_nav_items = mod_config.unit_nav_items
        self.accum_ret_items = mod_config.accum_ret_items
        self.ttl_nav_items = mod_config.ttl_nav_items

        for i, (f_name, date_ptn) in enumerate(zip(self.data_files, self.date_patterns)):
            # date_pattern 和 data_file 两个 列表必须一一对应，否则会出错
            if date_ptn == 'yyyymmdd':
                self.data_files[i] = ''.join([f_name, self.dt_date.strftime('%Y%m%d'), '.xls'])
            elif date_ptn == 'yyyy-mm-dd':
                self.data_files[i] = ''.join([f_name, self.dt_date.strftime('%Y-%m-%d'), '.xls'])
        self.nav_map = {f_:0 for f_ in self.data_files} # 文件映射
        self.res_dict = {prod_:NavNode() for prod_ in self.prod_names}


    def parse_nav_files(self):
        """ 解析文件，提取需要信息 """
        for prod_idx, data_file in enumerate(self.data_files):
            f_path = os.path.join(self.input_path, data_file)
            try:
                with xlrd.open_workbook(f_path) as nav_wb:
                    sheet_names = nav_wb.sheet_names()
                    nav_sheet = nav_wb.sheet_by_name(sheet_names[0])
            except IOError as e:
                print('Cannot open file. 无法打开数据文件.', e)
                raise  # 继续传递给上一层做异常消息处理
            self.__get_unit_nav(nav_sheet, prod_idx)    # 提取unit nav
            self.__get_accum_ret(nav_sheet, prod_idx)   # 提取return
            self.__get_ttl_nav(nav_sheet, prod_idx)     # 提取total net assets
        return self.res_dict

    def __get_unit_nav(self, nav_sheet, prod_idx):
        for row in nav_sheet.get_rows():
            for col in row:
                if self.unit_nav_items[prod_idx] in str(col.value):
                    prod_name = self.prod_names[prod_idx]
                    r_ptn = r'-?\d+\.?\d*'
                    unit_nav = re.search(r_ptn, col.value).group()
                    self.res_dict[prod_name].data_file = self.data_files[prod_idx]
                    self.res_dict[prod_name].unit_nav = unit_nav
                    return

    def __get_accum_ret(self, nav_sheet, prod_idx):
        for row in nav_sheet.get_rows():
            for col in row:
                if self.accum_ret_items[prod_idx] in str(col.value):
                    prod_name = self.prod_names[prod_idx]
                    r_ptn = r'-?\d+\.?\d*'
                    ret_col = str(row[1].value)  # 紧挨标题栏
                    accum_ret = re.search(r_ptn, ret_col).group()
                    self.res_dict[prod_name].accum_ret = accum_ret + '%'
                    return

    def __get_ttl_nav(self, nav_sheet, prod_idx):
        mkv_idx = 0
        for row in nav_sheet.get_rows():
            for i, col in enumerate(row):
                if '市值' == col.value:
                    mkv_idx = i
                elif self.ttl_nav_items[prod_idx] in str(col.value):
                    prod_name = self.prod_names[prod_idx]
                    r_ptn = r'-?\d+\.?\d*'
                    ttl_nav_col = str(row[mkv_idx].value)
                    ttl_nav = re.search(r_ptn, ttl_nav_col).group()
                    self.res_dict[prod_name].ttl_nav = ttl_nav
                    return

