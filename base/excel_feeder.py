import os
import threading
import time
import traceback
from datetime import datetime
from pprint import pprint
from multiprocessing import Process, Pool
import xlwings as xw
from xlwings import Book

from base.stock import Stock
from xlwings.constants import DeleteShiftDirection

DEFULT_START_D = '2019-11-01'
DEFULT_END_D = '2019-12-31'
INDEX_TO_START_STOCK_VAL = 3
stocks_translate_dict = {}


class ExcelFeeder(object):
    def __init__(self, full_name="tmp_excel.xlsx"):
        self.date_start_before = DEFULT_START_D
        self.date_end_before = DEFULT_END_D
        self.full_name = full_name
        self.work_book = self.create_excel(full_name)
        self.create_sheets()

        self.preprocess_excel()
        self.start_loop()
        self._stocks_translate_dict = stocks_translate_dict

    def create_sheets(self):
        self.sheet_algo_stock = self.work_book.sheets[0]
        self.name_stock_algo = "Algo Stock"
        self.sheet_algo_stock.name = self.name_stock_algo
        self.sheet_graph_stock = self.work_book.sheets.add()
        self.name_stock_graph = "Graph Stock"
        self.sheet_graph_stock.name = self.name_stock_graph
        self.sheet_work_stock = self.work_book.sheets.add()
        self.name_stock_work = "Live Stock"
        self.sheet_work_stock.name = self.name_stock_work
        self.sheet_work_stock.activate()

    def preprocess_excel(self):
        self.sheet_work_stock.range("B1").value = "Stock Treader Excel"
        self.sheet_work_stock.range("I1").api.Font.Bold = True
        self.sheet_work_stock.range("A2:j2").color = (96, 211, 249)  # blue header
        self.sheet_work_stock.range("A2:A2").color = (255, 165, 0)  # orange header
        self.sheet_work_stock.range("H2:J2").color = (255, 165, 0)  # orange header
        self.sheet_work_stock.range('A3:j32').color = (230, 230, 230)  # gray body
        self.sheet_work_stock.range('A2:j32').api.Borders.Weight = 3
        self.sheet_work_stock.range('A2:j32').api.Font.Bold = True

        self.sheet_work_stock.range("A3:E32").options(transpose=True).value = list(stocks_translate_dict.values())
        self.sheet_work_stock.range("A2").value = ['stock', "value", "bid", 'ask', "min", "max", "open",
                                                   "higher to send", "lower to send"]
        xw.Range('H2:I2').autofit()

        # graph chose
        self.sheet_work_stock.range("P2:Q2").color = (255, 165, 0)  # orange header
        self.sheet_work_stock.range("P2").value = ['start date', "end date"]
        self.sheet_work_stock.range("P3").value = ["2019-11-1", "2019-12-31"]
        xw.Range('P3:Q3').autofit()
        self.sheet_work_stock.range("P3:Q3").color = (230, 230, 230)  # gray body
        self.sheet_work_stock.range("P2:Q3").api.Borders.Weight = 3
        self.sheet_work_stock.range("P2:Q3").api.Font.Bold = True

    @staticmethod
    def thread_run(arg_dict: dict):
        cntr = 1
        mul = 1
        timer = time.time()
        while timer + 3 > time.time():
            stock_name = arg_dict.get("stock_display_name")
            ticker = arg_dict.get("ticker")
            ttt = time.time()
            print(str(stock_name) + f" ** START ** #{cntr} " + str(ttt))
            index = arg_dict.get("index")
            full_name = arg_dict.get("full_name")
            name_stock_work = arg_dict.get("name_stock_work")
            stock = Stock(stock_ticker=ticker)
            stock_vals = [
                "price", "bid", "ask", "regularMarketDayLow", "regularMarketDayHigh", "regularMarketOpen"
            ]
            tmp_char = "B"
            for i, type_val in enumerate(stock_vals):
                xw.Book(full_name).sheets[name_stock_work].range(
                    f"{chr(ord(tmp_char) + i)}{index + INDEX_TO_START_STOCK_VAL}").value = [
                    getattr(stock, type_val)
                ]

    def start_loop(self):
        global stocks_translate_dict
        print("start loop")
        while True:
            self.live_stock_sheet()
            print(f"live_stock_sheet " * 10)
            self.graph_stock_sheet_run()
            print(stocks_translate_dict)

            print("amazing")
            time.sleep(1)

    def graph_stock_sheet_run(self):
        global stocks_translate_dict
        print("self.get_date_start() ", self.get_date_start())
        print("self.date_start_before ", self.date_start_before)
        print("  self.get_date_end() ", self.get_date_end())
        print("  self.date_end_before ", self.date_end_before)
        print("self.get_date_start() is not self.date_start_before")
        print(self.get_date_start() != self.date_start_before)
        print(self.get_date_end() != self.date_end_before)
        if self.get_date_start() != self.date_start_before or self.get_date_end() != self.date_end_before:
            print("in if " * 10)
            self.delete_all_charts()
            self.date_start_before = self.get_date_start()
            self.date_end_before = self.get_date_end()

        try:
            for i, chart in enumerate(self.sheet_graph_stock.charts):
                if not chart.api[1].ChartTitle.Text in list(stocks_translate_dict.keys()):
                    print(" chart.delete() " * 100)
                    chart.delete()
                    self.sheet_graph_stock.range((1000, ((i + 1) * 2)), (3000, ((i + 2) * 2))).clear_contents()
        except Exception as e:
            print(str(e))
            traceback.print_exc()
        index_t = 1
        for ticker, heb_name in stocks_translate_dict.items():
            self.add_chart_graph_sheet(stock_name=ticker, index=index_t)
            index_t += 1

    def get_date_start(self):
        date_start_excel = xw.Book(self.full_name).sheets[self.name_stock_work].range(f"P3").value
        if not date_start_excel:
            return DEFULT_START_D
        else:
            date_start_excel = date_start_excel.strftime("%Y-%m-%d")
        return date_start_excel

    def get_date_end(self):
        date_end_excel = xw.Book(self.full_name).sheets[self.name_stock_work].range(f"Q3").value
        if not date_end_excel:
            return DEFULT_START_D
        else:
            date_end_excel = date_end_excel.strftime("%Y-%m-%d")
        return date_end_excel

    def delete_all_charts(self):
        self.sheet_graph_stock.clear_contents()
        for chart in self.sheet_graph_stock.charts:
            chart.delete()

    def add_chart_graph_sheet(self, stock_name, index):
        for chart in self.sheet_graph_stock.charts:
            if chart.api[1].ChartTitle.Text in stock_name:
                return
        stock_t = Stock(stock_name)
        sht = self.sheet_graph_stock
        row_num = 1000
        col_num = index * 2
        print(f"start {self.get_date_start()}")
        print(f"end {self.get_date_end()}")
        sht.range((row_num, col_num)).value = stock_t.get_data_history(start_d=self.get_date_start(),
                                                                       end_d=self.get_date_end())
        # sht.range(data_loc).value = stock_t.get_data_history(start_d='2019-11-1', end_d='2019-12-31')
        chart = sht.charts.add()
        chart.set_source_data(sht.range((row_num, col_num)).expand())
        # chart.set_source_data(sht.range(data_loc).expand())
        chart.chart_type = 'line_markers'
        chart.api[1].SetElement(2)  # Place chart title at the top
        chart.api[1].ChartTitle.Text = stock_name  # Change text of the chart title
        chart.width = 1000
        chart.top = 211.0 * (index - 1)

    def live_stock_sheet(self):
        import copy
        global stocks_translate_dict
        list_dict = []
        trds = []
        index = 0
        temp_dict = {}
        flag_stop = True
        i = 0
        while name := xw.Book(self.full_name).sheets[self.name_stock_work].range(
                f"A{i + INDEX_TO_START_STOCK_VAL}").value:
            print(f"name {name}")
            if name:
                temp_dict[name] = name
            i += 1
        stocks_translate_dict = copy.deepcopy(temp_dict)
        for ticker, heb_name in stocks_translate_dict.items():
            arg_dict = {
                "stock_display_name": heb_name,
                "ticker": ticker,
                "index": index,
                "full_name": self.full_name,
                "name_stock_work": self.name_stock_work,
            }
            list_dict.append(arg_dict)
            # self.thread_run(arg_dict)
            x = threading.Thread(target=self.thread_run, args=(arg_dict,))
            trds.append(x)
            x.daemon = True
            x.start()
            index += 1
        # with Pool(len(list_dict)) as p:
        #     p.map(self.thread_run,
        #           list_dict)
        print("waiting to finish...... ")
        for thread in trds:
            thread.join()

    @staticmethod
    def create_excel(full_name: str) -> Book:
        try:
            [i.kill() for i in xw.apps]
            os.remove(full_name)
            wb = xw.Book(full_name)
        except Exception as e:
            wb = xw.Book()
            wb.save(full_name)
        print("finish create_excel")
        return wb
