from excel_package.sheet.graph_sheet import GraphSheet
import time

from excel_package.stock_data.live_stock import LiveStock


class StockReturnPerDay(GraphSheet):

    def pre_run(self):
        pass

    def run_sheet(self):
        self.graph_type = "bar_clustered"
        while True:
            print("StockReturnPerDay")
            self.global_graph()
            time.sleep(5)

    def get_df(self, stock_name):
        stock_t = LiveStock(stock_name)
        df = stock_t.get_data_history(start_d=self.get_date_start(),
                                      end_d=self.get_date_end())
        df["return"] = df["Close"].pct_change(1)

        return df["return"]
