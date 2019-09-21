# -*- Utf-8 -*-

#
# NavReporter.py
#
# Automatically scan email, download excels,
#
# parse the excel daily report, collect the data
#
# and send them to stakeholders.
#
# All rights reserved, 2019
#

from EmailScanner import EmailScanner
from NavParser import NavParser
from Config import Config
from ReportFiller import ReportFiller
from ReportSender import ReportSender
from Calendar import Calendar


class NavReporter(object):

    def __init__(self):
        self.scanner = EmailScanner.instance()
        self.parser = NavParser.instance()
        self.filler = ReportFiller.instance()
        self.sender = ReportSender.instance()

    def run(self, conf_file):
        try:
            conf_mod = Config(conf_file)
            conf_mod.load_config()
            conf_info = conf_mod.parse_config()
            Calendar.instance().load_calendar_file(conf_info.calendar.calendar_file)
            if Calendar.instance().is_trading_day():
                self.scanner.set_conf_params(conf_info.scanner)
                self.parser.set_conf_params(conf_info.parser)
                self.filler.set_conf_params(conf_info.filler)
                self.sender.set_conf_params(conf_info.sender)
                #
                self.scanner.scan_nav_files()
                nav_dict = self.parser.parse_nav_files()
                self.filler.fill_report_files(nav_dict)
                self.sender.send_report_mail()
            else:
                pass
        except Exception as e:
            err_msg = '程序运行出现异常，相关信息:{e_info}。请检查报告程序'.format(e_info=e)
            self.sender.send_error_message(err_msg)


if __name__ == '__main__':
    n_rpt = NavReporter()
    n_rpt.run('./NavReporter/config.cfg')