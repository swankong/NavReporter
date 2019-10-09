# -*- coding:utf-8 -*-

#
# EmailScanner.py
#
# Scan email account and download data files.
#
# All rights reserved, 2019
#

import os
import re
import poplib
import email
import smtplib
import base64
import threading
import shutil
import quopri

import datetime as dt

from email.parser import Parser
from Calendar import Calendar


class EmailScanner(object):

    """ Make a singleton class """
    _instance_lock = threading.Lock()
    @classmethod
    def instance(cls):
        """ Singleton pattern """
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:
                if not hasattr(cls, '_instance'):
                    cls._instance = EmailScanner()
        return cls._instance

    def __init__(self):
        self.dt_date = None
        self.pop3_host = None
        self.pop3_user = None
        self.pop3_pass = None
        self.data_files = None
        self.file_names = None
        self.data_providers = None
        self.date_patterns = None
        self.nav_save_path = None
        self.nav_map = None
        self.__mail_scanned = 0  # private email counter
        self.__MAX_NUM_SCAN = 500

    def set_conf_params(self, mod_config):
        self.dt_date = Calendar.instance().get_prev_trading_day()
        self.pop3_host = mod_config.pop3_host      
        self.pop3_user = mod_config.pop3_user
        self.pop3_pass = mod_config.pop3_password
        self.data_files = mod_config.data_files.copy()
        self.file_names = mod_config.data_files.copy()
        self.data_providers = mod_config.data_providers
        self.date_patterns = mod_config.date_patterns
        self.nav_save_path = mod_config.output_path  # Data file path
        for i, (f_name, date_ptn) in enumerate(zip(self.data_files, self.date_patterns)):
            # date_pattern 和 data_file 两个列表必须一一对应，否则会出错
            # 上述实现效率并不高，但方便，如果讲求效率需要优化
            if date_ptn == 'yyyymmdd':
                self.data_files[i] = ''.join([f_name, self.dt_date.strftime('%Y%m%d'), '.xls'])
            elif date_ptn == 'yyyy-mm-dd':
                self.data_files[i] = ''.join([f_name, self.dt_date.strftime('%Y-%m-%d'), '.xls'])
        self.nav_map = {f_:0 for f_ in self.data_files} # 文件映射

    def scan_nav_files(self):
        mail_set, server_conn = self.__get_mail_set()
        try :
            self.__search_target_mail(mail_set, server_conn)
        except ValueError as e:  # Cannot get all the needed data files.
            self.__replicate_old_nav_files()
        self.__close_connection(server_conn)

    def __get_mail_set(self):
        server_conn = poplib.POP3(self.pop3_host)
        server_conn.user(self.pop3_user)
        server_conn.pass_(self.pop3_pass)
        msg_cnt, msg_size = server_conn.stat()  # stat()返回邮件数量和占用空间
        resp_, mails_, octets_ = server_conn.list()  # list()返回所有邮件的编号
        return mails_, server_conn

    def __valid_sender(self, from_str):
        """ 判断邮件发送者是否是数据服务商 """
        for mail_sender in self.data_providers:
            if from_str.find(mail_sender) != -1:
                return True
        else:
            return False

    def __get_mail_content(self, mail_msg, nav_files):
        attachment_file = None
        nav_save_path = self.nav_save_path
        for part in mail_msg.walk():
            if part.is_multipart():
                continue
            file_name = part.get_filename()
            # 保存附件
            if file_name:  # Attachment
                # Decode filename
                hdr_ = email.header.Header(file_name)
                dh = email.header.decode_header(hdr_)
                dh_fname = dh[0][0].decode('ASCII')
                decoded_fname = self.__decode_mail_text(dh_fname)
                if not decoded_fname in self.data_files:  # 不是所需要的附件，返回
                    return attachment_file
                data_ = part.get_payload(decode=True)
                with open(os.path.join(nav_save_path, decoded_fname), 'wb') as att_file:
                    att_file.write(data_)
                attachment_file = decoded_fname
                break
            else:
                continue # 不是附件，跳过
        return attachment_file

    def __finish_scan(self):
        scan_list = [x for x in self.nav_map.values()]
        if all(scan_list):
            return True
        else:
            return False

    def __search_target_mail(self, mails_, mail_server):
        # 获取最新一封邮件, 注意索引号从1开始:
        n_mails = len(mails_)
        mail_parser = Parser()
        for i in range(n_mails, 0, -1):
            try :
                msg_head = mail_server.top(i, 0)
                head_dict = mail_parser.parsestr('\n'.join([h_.decode('GB2312') for h_ in msg_head[1]]))
                if not self.__valid_sender(head_dict['From']):  # 不是目标发送人，跳过
                    self.__increase_mail_counter()
                    continue
                resp, lines, octets = mail_server.retr(i)
                # lines存储了邮件的原始文本的每一行,  可以获得整个邮件的原始文本:
                msg_str = '\r\n'.join([x.decode('UTF-8') for x in lines])
                msg_content = mail_parser.parsestr(msg_str)
            except Exception as e:
                self.__increase_mail_counter()
                continue
            attachment_file = self.__get_mail_content(msg_content, self.data_files)
            if not attachment_file:
                self.__increase_mail_counter()
                continue
            self.nav_map[attachment_file] = 1
            if self.__finish_scan():
                return
            self.__increase_mail_counter()

    def __close_connection(self, server_conn):
        # 关闭连接:
        server_conn.quit()

    def __decode_mail_text(self, mail_text):
        aa, bb = '', ''
        res = []
        if ('UTF-8' in mail_text.upper()) or ('GBK' in mail_text.upper()):  # 字符集编码
            text_list = mail_text.split('\r\n')  # 之前处理使用换行分割，现在去掉换行符
            for text_line in text_list:    # 逐行处理
                if ('UTF-8' in text_line.upper()) or ('GBK' in text_line.upper()):
                    reg_ptn = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='   # 正则表达式模式匹配报文头信息
                    text_line = text_line.strip()   # 去掉空格
                    charset, encoding, encoded_text = re.match(reg_ptn, text_line).groups()  # 正则表达式匹配
                    if encoding is 'B':  # 获得信息，这段邮件正文是Base64 编码
                        byte_string = base64.b64decode(encoded_text)    # base64 解码
                    elif encoding is 'Q':  # 获得信息，这段邮件正文是Quoprint 编码
                        byte_string = quopri.decodestring(encoded_text)  # QP解码
                    aa = byte_string.decode(charset)    # 字符编码的解码，注意：邮件编码和字符编码不是一回事
                    res.append(aa)         # 处理完放到行列表中
                else:
                    bb = text_line         # 无序解码处理
                    res.append(bb)
            return ''.join(res)            # 处理完拼接成一整段内容
        else:
            return mail_text

    def __increase_mail_counter(self):
        self.__mail_scanned += 1
        print(self.__mail_scanned)
        if self.__mail_scanned > self.__MAX_NUM_SCAN:
            raise ValueError('Scanned more than 500 mails and not find all the data files. 扫描超过500邮件没有发现所有目标数据文件.')

    def __replicate_old_nav_files(self):
        """
            对于当日数据文件缺失情况，如果需要的文件在邮件中没有找到，
            则复制前一日的数据文件作为代替，这样实现最简单，后续的解析
            文件程序直接把前一日数据当作当日数据解析，但必须保证前一日
            数据文件存在。
        :return:
        """
        def __make_file_name(ref_date, date_ptn, file_name):
            if date_ptn == 'yyyymmdd':
                ret_name = ''.join([file_name, ref_date.strftime('%Y%m%d'), '.xls'])
            elif date_ptn == 'yyyy-mm-dd':
                ret_name = ''.join([file_name, ref_date.strftime('%Y-%m-%d'), '.xls'])
            else:
                ret_name = ''
            return ret_name
        
        rep_date = Calendar.instance().get_prev_trading_day(ref_date=self.dt_date)  # 再向前一个交易日 T-2
        # 遍历文件映射
        for i, (expt_file, v) in enumerate(self.nav_map.items()):
            if not v:  # 未找到的文件
                old_file = __make_file_name(rep_date, self.date_patterns[i], self.file_names[i])
                old_path = os.path.join(self.nav_save_path, old_file)
                new_path = os.path.join(self.nav_save_path, expt_file)
                shutils.copy(old_path, new_path)



