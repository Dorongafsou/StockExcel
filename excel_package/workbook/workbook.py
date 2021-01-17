import os
import xlwings as xw
from xlwings import Book
from threading import Thread

from excel_package.sheet.live_sheet import LiveSheet
from excel_package.sheet.stock_close_price_g import StockCloseGraph
from excel_package.sheet.stock_cumulative_daly_return import StockCumulativeDalyReturn
from excel_package.sheet.stock_return_per_day import StockReturnPerDay
from excel_package.sheet.stock_total_money import StockTotalGraph
from excel_package.utills.utill_setting import LIVE_STOCK, GRAPH_STOCK, GRAPH_TOTAL, GRAPH_RETURN_PER_DAY, \
    CUMULATIVE_DALY_RETURN

SHEET_NAMES_DEFAULT = {
    LIVE_STOCK: LiveSheet,
    GRAPH_STOCK: StockCloseGraph,
    GRAPH_TOTAL: StockTotalGraph,
    GRAPH_RETURN_PER_DAY: StockReturnPerDay,
    CUMULATIVE_DALY_RETURN: StockCumulativeDalyReturn,
}



DEFAULT_START_D = '2019-11-01'
DEFAULT_END_D = '2019-12-31'
START_DATE_CELL = "Q2"
END_DATE_CELL = "Q3"
# sheets names
LIVE_STOCK = "Live Stock"
GRAPH_STOCK = "Graph Stock"
GRAPH_TOTAL = "Graph total money"
GRAPH_RETURN_PER_DAY = "Graph retrun per day "
INDEX_TO_START_STOCK_VAL = 3


class WorkBook(object):
    def __init__(self, name):
        self._name = name
        self.wb = self.create_excel(name)
        self.sheet_names = SHEET_NAMES_DEFAULT
        self.sheet_list = self.create_sheets()

    def pre_run(self):
        [sheet.pre_run() for sheet in self.sheet_list]

    def real_time(self):
        threads = [Thread(target=self.sheet_list[0].run_sheet)]
        threads += [Thread(target=self.run_graph, args=[self.sheet_list[1:]])]
        [t.start() for t in threads]


    @staticmethod
    def run_graph(graph_sheets):
        while True:
            [graph_s.run_sheet() for graph_s in graph_sheets]

    def create_sheets(self):
        sheet_object = [sheet_t[1](sheet_t[0], index, self._name) for index, sheet_t in enumerate(SHEET_NAMES_DEFAULT.items())]
        sheet_object.reverse()
        if len(sheet_object):
            sheet = sheet_object[0]
            sheet.xlwing_sheet = self.wb.sheets[0]
            for sheet in sheet_object[1:]:
                sheet.xlwing_sheet = self.wb.sheets.add()
            sheet_object[-1].xlwing_sheet.activate()
        sheet_object.sort()
        return sheet_object

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
