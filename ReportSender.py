# -*- coding: utf-8 -*-

#
# ReportSender.py
#
# Send the report by email to stakeholders.
#
# All rights reserved, 2019
#


import os
import threading
import smtplib
import datetime as dt


from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from Calendar import Calendar


class ReportSender(object):
    """ Send the report file by email to stakeholders.
        Make a singleton class """
    _instance_lock = threading.Lock()

    @classmethod
    def instance(cls):
        """ Singleton pattern """
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:
                if not hasattr(cls, '_instance'):  # double lock check
                    cls._instance = ReportSender()
        return cls._instance

    def __init__(self):
        self.dt_date = None
        self.src_smtp = None
        self.smtp_user = None
        self.smtp_pass = None
        self.des_accounts = None
        self.msg_accounts = None
        self.report_file = None

    def set_conf_params(self, mod_config):
        self.dt_date = Calendar.instance().get_prev_trading_day()
        self.src_smtp = mod_config.src_smtp
        self.smtp_user = mod_config.src_user
        self.smtp_pass = mod_config.src_password
        self.des_accounts = mod_config.des_accounts
        self.msg_accounts = mod_config.msg_accounts
        self.report_file = mod_config.report_file

    def send_report_mail(self):
        mail_from = self.smtp_user    # 发件人
        mail_to = self.des_accounts   # 收件人列表
        smtp_svr = self.src_smtp      # smtp服务器
        rpt_msg = '数据每日汇总'
        message_ = MIMEMultipart()
        mail_text = MIMEText(rpt_msg, 'plain', 'UTF-8')    # 正文
        message_.attach(mail_text)     # 邮件正文写入
        message_['Subject'] = '数据每日汇总-{d}'.format(d=self.dt_date.strftime('%Y%m%d'))
        message_['From'] = mail_from
        with open(self.report_file, 'rb') as f:
            attach_report = MIMEApplication(f.read())     # 创建附件文件并编码
        attach_report['Content-type'] = 'application/octet-stream'
        attach_report.add_header('Content-Disposition', 'attachment', filename=os.path.basename(self.report_file))
        message_.attach(attach_report)    # 添加附件
        # connect to server and send the mail
        smtp_obj = smtplib.SMTP()
        smtp_obj.connect(smtp_svr, 25)
        smtp_obj.login(self.smtp_user, self.smtp_pass)    # 登录
        smtp_obj.sendmail(mail_from, mail_to, message_.as_string())    # 发送
        smtp_obj.close()    # 关闭连接，结束

    def send_error_message(self, msg):
        mail_from = self.smtp_user
        mail_to = self.msg_accounts
        smtp_svr = self.src_smtp
        message_ = MIMEMultipart()
        mail_text = MIMEText(msg, 'plain', 'UTF-8')
        message_.attach(mail_text)
        message_['Subject'] = '[错误报告]：每日汇总-{d}-出现错误，请检查'.format(d=self.dt_date.strftime('%Y%m%d'))
        message_['From'] = mail_from
        # connect to server and send the mail
        smtp_obj = smtplib.SMTP()
        smtp_obj.connect(smtp_svr, 25)
        smtp_obj.login(self.smtp_user, self.smtp_pass)
        smtp_obj.sendmail(mail_from, mail_to, message_.as_string())
        smtp_obj.close()


