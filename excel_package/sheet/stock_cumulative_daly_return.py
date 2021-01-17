from excel_package.sheet.graph_sheet import GraphSheet

from excel_package.stock_data.live_stock import LiveStock


class StockCumulativeDalyReturn(GraphSheet):

    def pre_run(self):
        super(StockCumulativeDalyReturn, self).pre_run()
        pass

    def run_sheet(self):
        print("stock_cumulative_daly_return")
        self.global_graph()

    def get_df(self, stock_name):
        stock_t = LiveStock(stock_name)
        df = stock_t.get_data_history(start_d=self.get_date_start(),
                                      end_d=self.get_date_end())
        df["return"] = df["Close"].pct_change(1)
        df["Cumulative Return"] = (1+df["return"]).cumprod()
        return df["Cumulative Return"]
