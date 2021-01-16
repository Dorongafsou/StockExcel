from datetime import datetime
from pprint import pprint

from pandas_datareader import data

import time


class LiveStock(object):
    def __init__(self, stock_ticker):
        self.stock_ticker = stock_ticker

    def update_stock(self, stock_ticker: str) -> dict:
        self.__dict__.update(self.get_stock_info(stock_ticker))

    def get_stock_info(self, stock_ticker):
        return self.convert_stock_dict(data.get_quote_yahoo([stock_ticker]).to_dict())

    def get_data_history(self, start_d='2019-1-1', end_d='2019-12-31'):
        return data.DataReader(self.stock_ticker,
                               start=start_d,
                               end=end_d,
                               data_source='yahoo')['Adj Close']

    @staticmethod
    def convert_stock_dict(dict_stock: dict) -> dict:
        return {k: v[list(v.keys())[0]] for k, v in dict_stock.items()}


if __name__ == '__main__':
    pprint(LiveStock("NVDA").price)
