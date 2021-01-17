import time

from excel_package.sheet.sheet import Sheet

import xlwings as xw
from threading import Thread

from excel_package.stock_data.live_stock import LiveStock
from excel_package.utills.utill_setting import INDEX_TO_START_STOCK_VAL


class LiveSheet(Sheet):

    def pre_run(self):
        self._xlwing_sheet.range("B1").value = "Stock trader excel_package"
        self._xlwing_sheet.range("I1").api.Font.Bold = True
        self._xlwing_sheet.range("A2:j2").color = (96, 211, 249)  # blue header
        self._xlwing_sheet.range("A2:A2").color = (255, 165, 0)  # orange header
        self._xlwing_sheet.range("H2:J2").color = (255, 165, 0)  # orange header
        self._xlwing_sheet.range('A3:j32').color = (230, 230, 230)  # gray body
        self._xlwing_sheet.range('A2:j32').api.Borders.Weight = 3
        self._xlwing_sheet.range('A2:j32').api.Font.Bold = True

        # self._xlwing_sheet.range("A3:E32").options(transpose=True).value = list(stocks_translate_dict.values())
        self._xlwing_sheet.range("A2").value = ['stock', "value", "bid", 'ask', "min", "max", "open",
                                                "higher to send", "lower to send"]
        xw.Range('H2:I2').autofit()

        # graph chose
        self._xlwing_sheet.range("P1").value = "User option"
        self._xlwing_sheet.range("P2:P3").color = (255, 165, 0)  # orange header
        self._xlwing_sheet.range("P2").value = ['start date', "2019-11-1"]
        self._xlwing_sheet.range("P3").value = ["end date", "2019-12-31"]
        xw.Range('P3:Q3').autofit()
        self._xlwing_sheet.range("Q2:Q3").color = (230, 230, 230)  # gray body
        self._xlwing_sheet.range("P2:Q3").api.Borders.Weight = 3
        self._xlwing_sheet.range("P2:Q3").api.Font.Bold = True

    @staticmethod
    def thread_run(arg_dict: dict):
        ticker = arg_dict.get("ticker")
        index = arg_dict.get("index")
        stock = LiveStock(stock_ticker=ticker)
        stock.update_stock(ticker)
        stock_vals = [
            "price", "bid", "ask", "regularMarketDayLow", "regularMarketDayHigh", "regularMarketOpen"
        ]
        tmp_char = "B"
        try:
            for i, type_val in enumerate(stock_vals):
                xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{chr(ord(tmp_char) + i)}{index}").value = [
                    getattr(stock, type_val)
                ]
            xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f'H{index}').value = [int(time.time() - 1610825060.4012973)]
        except Exception as ex:
            return None

    def run_sheet(self):
        index = 0
        # Add friendly_Name
        threads = {num: None for num in range(30)}

        while True:
            for index in range(30):
                name_tk = xw.Book(self.file_name).sheets[self.name].range(f"A{index + INDEX_TO_START_STOCK_VAL}").value
                if not name_tk:
                    continue
                if threads[index] is not None and threads[index].is_alive():
                    continue
                arg_dict = {
                    "ticker": name_tk,
                    "index": index + INDEX_TO_START_STOCK_VAL,
                    "name": self.name,
                    "file_name": self.file_name
                }
                threads[index] = Thread(target=self.thread_run, args=(arg_dict,))
                threads[index].start()
