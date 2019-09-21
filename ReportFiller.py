# -*- coding: utf-8 -*-

#
# ReportFiller.py
#
# Fill product information into the Excel report.
#
# All rights reserved, 2019
#

import xlrd
import xlwt

import threading
import datetime as dt

import xlutils.copy
import xlutils.styles

from Calendar import Calendar


class ReportFiller(object):

    """ Make a singleton class """
    _instance_lock = threading.Lock()
    @classmethod
    def instance(cls):
        """ Singleton pattern """
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:
                if not hasattr(cls, '_instance'):
                    cls._instance = ReportFiller()
        return cls._instance

    def __init__(self):
        self.dt_date = None
        self.report_file = None
        self.prod_names = None

    def set_conf_params(self, conf_info):
        self.dt_date = Calendar.instance().get_prev_trading_day()
        self.report_file = conf_info.report_file
        self.prod_names = conf_info.prod_names

    def fill_report_files(self, res_dict):
        date_tag = '{m}月{d}日'.format(m=self.dt_date.month, d=self.dt_date.day)
        # xlrd 读取 Excel，注意，这里只读
        with xlrd.open_workbook(self.report_file, formatting_info=True) as rpt_wb:
            sheet_names = rpt_wb.sheet_names()
            rpt_sheet = rpt_wb.sheet_by_name(sheet_names[0])
            wrt_wb = xlutils.copy.copy(rpt_wb)
            wrt_sheet = wrt_wb.get_sheet(sheet_names[0])
            for k, v in res_dict.items():
                row_1, col_1 = self.__get_cell_pos(rpt_sheet, b'\xe5\x8d\x95\xe4\xbd\x8d\xe5\x87\x80\xe5\x80\xbc'.decode('utf-8'), k, date_tag)
                row_2, col_2 = self.__get_cell_pos(rpt_sheet, b'\xe8\xa7\x84\xe6\xa8\xa1'.decode('utf-8'), k, date_tag)
                row_3, col_3 = self.__get_cell_pos(rpt_sheet, b'\xe6\x94\xb6\xe7\x9b\x8a\xe7\x8e\x87'.decode('utf-8'), k, None)
                self.__write_cell_pos(wrt_sheet, row_1, col_1, v.unit_nav)
                self.__write_cell_pos(wrt_sheet, row_2, col_2, v.ttl_nav)
                self.__write_cell_pos(wrt_sheet, row_3, col_3, v.accum_ret)
        wrt_wb.save(self.report_file)

    def __get_cell_pos(self, work_sheet, item_name, prod_name, date_tag):
        """ 查找目标单元格坐标 """
        row_iter = iter(work_sheet.get_rows())  # 使用迭代器，不中断
        x_cord, y_cord = 0, 0
        for row_ in row_iter:
            for col_ in iter(row_):
                if item_name in str(col_.value):  # 找到 项目栏目开始位置
                    break
            else:  # 如果内循环中断，则外循环也中断
                x_cord += 1
                continue
            break
        else:
            raise ValueError('Cannot find corresponding item in report file. 报告文件中找不到相应栏目.')
        row_ = next(row_iter)  # 迭代一行
        x_cord += 1

        if date_tag is not None:
            for col_ in row_:   # 找到相应的列
                if date_tag in str(col_.value):
                    break
                y_cord += 1
            else:
                raise ValueError('Cannot find correspoding date tag in report file. 报告文件中找不到相应的日期标记.')

        x_offset = 0
        for row_ in row_iter:
            for col_idx, col_ in enumerate(row_):
                if str(col_.value) == '':
                    continue
                elif prod_name in str(col_.value):
                    if date_tag is None:
                        y_cord = col_idx + 2
                    break
            else:  # 如果内循环中断，外循环也中断
                x_cord += 1
                x_offset += 1
                continue
            break
        else:
            raise ValueError('Cannot find corresponding prod name in report file. 报告文件中找不到相应产品名.')
        if x_offset > 10:
            # 偏移过大，疑似找到了其他栏目中的相同产品
            raise ValueError('Cannot find corresponding prod name in report file. 报告文件中找不到相应产品名.')
        return x_cord + 1, y_cord

    def __write_cell_pos(self, work_sheet, row_cord, col_cord, wrt_content):
        """
        :param work_sheet: 被写入工作簿
        :param row_cord:   单元格行坐标
        :param col_cord:   单元格列坐标
        :param wrt_content: 写入内容
        :return: 无
        """
        # 写入到报告工作簿
        out_row = work_sheet._Worksheet__rows.get(row_cord)  # 获得被写入行
        old_cell = out_row._Row__cells.get(col_cord)         # 根据列坐标定位写入单元格并拷贝属性
        work_sheet.write(row_cord, col_cord, wrt_content)    # 写入单元格
        out_row = work_sheet._Worksheet__rows.get(row_cord)  # 重新获得写入行
        new_cell = out_row._Row__cells.get(col_cord)         # 获得被写入单元格属性
        new_cell.xf_idx = old_cell.xf_idx                    # 将原属性赋值给新单元格，相当于恢复格式



