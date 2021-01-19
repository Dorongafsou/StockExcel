from datetime import datetime
from pandas_datareader import data

import time


class Stock(object):
    def __init__(self, stock_ticker):
        self.stock_ticker = stock_ticker.upper()
        # need to take it out from the builder
        self.__dict__.update(self.update_stock(stock_ticker))
        self.time_stock = time.time()
        self.date_time = datetime.fromtimestamp(self.time_stock)
        self.date_time_str = self.date_time.strftime("%H:%M:%S%f")

    def update_stock(self, stock_ticker: str) -> dict:
        return self.convert_stock_dict(data.get_quote_yahoo([stock_ticker]).to_dict())

    def get_data_history(self, start_d='2019-1-1', end_d='2019-12-31'):
        return data.DataReader(self.stock_ticker,
                               start=start_d,
                               end=end_d,
                               data_source='yahoo')['Adj Close']

    @staticmethod
    def convert_stock_dict(dict_stock: dict) -> dict:
        return {k: v[list(v.keys())[0]] for k, v in dict_stock.items()}


def timing(f):
    def wrap(*args, **kwargs):
        time1 = time.time()
        ret = f(*args, **kwargs)
        time2 = time.time()
        print('{:s} function took {:.3f} ms'.format(f.__name__, (time2 - time1) * 1000.0))

        return ret

    return wrap


