import time
from xlwings.utils import rgb_to_int

from excel_package.sheet.sheet import Sheet

from excel_package.Email.mail import send_mail
import xlwings as xw
from threading import Thread

from excel_package.stock_data.live_stock import LiveStock
from excel_package.utills.utill_setting import INDEX_TO_START_STOCK_VAL, HYPER_LINK


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
                                                "Buy/Sell", "count", "condition"]
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
        #     Add hyper link
        self._xlwing_sheet.range("L1").value = "Graph Links"
        links_val = [HYPER_LINK.format(sheet.name) for sheet in xw.Book(self.file_name).sheets]
        for i, link in zip(range(len(links_val)), links_val):
            self._xlwing_sheet.range(f"L{i + 2}").value = link

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
                xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"B{index}").api.Font.Color = [rgb_to_int((107, 142, 35)), rgb_to_int((139, 0, 0))][getattr(stock, "color") == "red"]

                xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{chr(ord(tmp_char) + i)}{index}").value = [
                    getattr(stock, type_val)
                ]
                if xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"J{index}").value is True:
                    mail_content = [str(obj) for obj in xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"H{index}:I{index}").options(numbers=int).value]
                    send_mail(getattr(stock, 'displayName'), ' '.join(mail_content))
                    xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"H{index}:J{index}").clear_contents()


        except Exception as ex:
            return None

    def run_sheet(self):
        index = 0
        # Add friendly_Name
        threads = {num: None for num in range(30)}

        while True:
            for index in range(30):
                try:
                    # Need to take one time all the stock names
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
                except Exception as e:
                    print(f" fail run_sheet live sheet {e}")
                    continue
