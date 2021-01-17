from excel_package.sheet.graph_sheet import GraphSheet
import time

from excel_package.stock_data.live_stock import LiveStock


class StockTotalGraph(GraphSheet):

    def pre_run(self):
        super(StockTotalGraph, self).pre_run()
        pass

    def run_sheet(self):
        print("StockTotalGraph")
        self.global_graph()

    def get_df(self, stock_name):
        stock_t = LiveStock(stock_name)
        df = stock_t.get_data_history(start_d=self.get_date_start(),
                                      end_d=self.get_date_end())
        df["total money"] = df["Close"] * df["Volume"]

        return df["total money"]
