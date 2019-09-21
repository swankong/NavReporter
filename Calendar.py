# -*- Utf-8 -*-

#
# Calendar.py
#
# Market calendar class.
#
# All rights reserved, 2019
#

import pandas as pd
import datetime as dt
import threading


class Calendar(object):
    """Make a singleton class """
    _instance_lock = threading.Lock()

    @classmethod
    def instance(cls):
        """ Singleton pattern """
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:
                if not hasattr(cls, '_instance'):  # double lock check
                    cls._instance = Calendar()
        return cls._instance

    def __init__(self):
        self.dt_today = dt.date(2019, 9, 19) # dt.datetime.now().date()
        self.dt_trade_dates = None

    def load_calendar_file(self, f_path):
        calen_df = pd.read_csv(f_path, delimiter=',', header=0, index_col=None)
        self.dt_trade_dates = [dt.datetime.strptime(str(x), '%Y%m%d').date() for x in calen_df.tradedate]

    def is_trading_day(self):
        return self.dt_today in self.dt_trade_dates

    def get_prev_trading_day(self, ref_date=None):
        if ref_date is None:
            if self.dt_today in self.dt_trade_dates:
                idx = self.dt_trade_dates.index(self.dt_today)
                return self.dt_trade_dates[idx - 1]
        else:
            if ref_date in self.dt_trade_dates:
                idx = self.dt_trade_dates.index(ref_date)
                if idx > 0:
                    return self.dt_trade_dates[idx - 1]
                else:
                    return None
            else:
                return None
