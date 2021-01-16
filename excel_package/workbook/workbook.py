import os
import xlwings as xw
from xlwings import Book

from excel_package.sheet.live_sheet import LiveSheet
from excel_package.sheet.stock_close_price_g import StockCloseGraph
LIVE_STOCK = "Live Stock"
GRAPH_STOCK = "Graph Stock"
SHEET_NAMES_DEFAULT = {
    LIVE_STOCK: LiveSheet,
    GRAPH_STOCK: StockCloseGraph,
}


class WorkBook(object):
    def __init__(self, name):
        self._name = name
        self.wb = self.create_excel(name)
        self.sheet_names = SHEET_NAMES_DEFAULT
        self.sheet_list = self.create_sheets()

    def pre_run(self):
        [sheet.pre_run() for sheet in self.sheet_list]

    def real_time(self):

        [sheet.run_sheet() for sheet in self.sheet_list]

    def add_sheet(self, name, index):
        pass

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
