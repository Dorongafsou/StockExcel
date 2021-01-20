import time
from xlwings.utils import rgb_to_int
from excel_package.sheet.sheet import Sheet
import yfinance as yf
from excel_package.Email.mail import send_mail
import xlwings as xw
from threading import Thread

from excel_package.stock_data.live_stock import LiveStock
from excel_package.stock_data.stock_interval import StockInterval
from excel_package.utills.utill_setting import *


class LiveSheet(Sheet):

    def pre_run(self):
        self._xlwing_sheet.range("B1").value = "Stock trader excel_package"
        self._xlwing_sheet.range("I1").api.Font.Bold = True
        self._xlwing_sheet.range("A2:j2").color = BLUE_HEADER  # blue header
        self._xlwing_sheet.range("A2:A2").color = ORANGE_HEADER  # orange header
        self._xlwing_sheet.range("I2:M2").color = ORANGE_HEADER  # orange header
        self._xlwing_sheet.range('A3:M32').color = (230, 230, 230)  # gray body
        self._xlwing_sheet.range('A2:M32').api.Borders.Weight = 3
        self._xlwing_sheet.range('A2:M32').api.Font.Bold = True

        # self._xlwing_sheet.range("A3:E32").options(transpose=True).value = list(stocks_translate_dict.values())
        self._xlwing_sheet.range("A2").value = ['stock', "value", "bid", 'ask', "min", "max", "open", "volume", "interval",
                                                "days ago", "Buy/Sell", "count", "condition"]

        # graph chose
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}1").value = "Graph dates"
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}2:{GRAPH_DATES_COL[0]}3").color = ORANGE_HEADER  # orange header
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}2").value = ['start date', "2019-11-1"]
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}3").value = ["end date", "2019-12-31"]
        xw.Range(f'{GRAPH_DATES_COL[1]}2:{GRAPH_DATES_COL[1]}3').autofit()
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[1]}2:{GRAPH_DATES_COL[1]}3").color = (230, 230, 230)  # gray body
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}2:{GRAPH_DATES_COL[0]}3").api.Borders.Weight = 3
        self._xlwing_sheet.range(f"{GRAPH_DATES_COL[0]}2:{GRAPH_DATES_COL[1]}3").api.Font.Bold = True
        #     Add hyper link
        self._xlwing_sheet.range(f"{LINK_COL}1").value = "Graph Links"
        links_val = [HYPER_LINK.format(sheet.name) for sheet in xw.Book(self.file_name).sheets]
        for i, link in zip(range(len(links_val)), links_val):
            self._xlwing_sheet.range(f"{LINK_COL}{i + 2}").value = link

        # Mail
        self._xlwing_sheet.range(f"{MAIL_COL[0]}1").value = "Mail for updates"
        self._xlwing_sheet.range(f"{MAIL_COL}2").api.Borders.Weight = 3
        self._xlwing_sheet.range(f"{MAIL_COL}2").api.Font.Bold = True

    @staticmethod
    def thread_run(arg_dict: dict):
        ticker = arg_dict.get("ticker")
        index = arg_dict.get("index")
        stock = LiveStock(stock_ticker=ticker)
        stock.update_stock(ticker)
        stock_vals = [
            "price", "bid", "ask", "regularMarketDayLow", "regularMarketDayHigh", "regularMarketOpen", "regularMarketVolume"
        ]
        tmp_char = "B"
        try:
            for i, type_val in enumerate(stock_vals):
                xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"B{index}").api.Font.Color = \
                    [rgb_to_int((107, 142, 35)), rgb_to_int((139, 0, 0))][getattr(stock, "color") == "red"]

                xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{chr(ord(tmp_char) + i)}{index}").value = [
                    getattr(stock, type_val)
                ]
                if xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"M{index}").value is True:
                    mail_content = [str(obj) for obj in
                                    xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"K{index}:L{index}").options(
                                        numbers=int).value]

                    dst_address = xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{MAIL_COL}2").value
                    if dst_address is not None:
                        send_mail(getattr(stock, 'displayName'), ' '.join(mail_content), dst_address)
                        xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"K{index}:M{index}").clear_contents()
                        xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{MAIL_COL}2").color = (255, 255, 255)

                    elif dst_address is None:
                        xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"{MAIL_COL}2").color = (139, 0, 0)
                interval_date = xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"I{index}:J{index}").value
                # move it to object oriented
                if interval_date[0]:
                    df_stock_interval = StockInterval().get_interval(tk=ticker, interval=interval_date[0], period=interval_date[1])
                    shape_len = df_stock_interval.shape[1] + 1
                    col = ((index-INDEX_TO_START_STOCK_VAL) * (shape_len + 1))

                    TK_NAME_LINE = 1
                    STOCK_INTERVAL_LINE = 2

                    xw.Book(arg_dict.get("file_name")).sheets[INTERVAL_SHEET].range(
                        (TK_NAME_LINE, col + 2)).value = getattr(stock, 'displayName') if hasattr(stock, 'displayName') else ticker

                    xw.Book(arg_dict.get("file_name")).sheets[INTERVAL_SHEET].range(
                        (TK_NAME_LINE, col + 2)).color = ORANGE_HEADER

                    xw.Book(arg_dict.get("file_name")).sheets[INTERVAL_SHEET].range(
                        (STOCK_INTERVAL_LINE, col + 2)).value = df_stock_interval

                    xw.Book(arg_dict.get("file_name")).sheets[INTERVAL_SHEET].range(
                        (STOCK_INTERVAL_LINE, col + 2),  (STOCK_INTERVAL_LINE, col + 1 + shape_len)).color = BLUE_HEADER

                    xw.Book(arg_dict.get("file_name")).sheets[arg_dict.get("name")].range(f"I{index}:J{index}").clear_contents()

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
