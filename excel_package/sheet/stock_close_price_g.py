from excel_package.sheet.graph_sheet import GraphSheet
import time

from excel_package.stock_data.live_stock import LiveStock


class StockCloseGraph(GraphSheet):

    def pre_run(self):
        pass

    def run_sheet(self):
        print("run_sheet graph stock")
        self.global_graph()

    def get_df(self, stock_name):
        stock_t = LiveStock(stock_name)
        return stock_t.get_data_history(start_d=self.get_date_start(),
                                        end_d=self.get_date_end())["Close"]
