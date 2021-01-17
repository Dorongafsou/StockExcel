import traceback
from abc import ABC

from base.excel_feeder import INDEX_TO_START_STOCK_VAL
from excel_package.sheet.sheet import Sheet
import xlwings as xw

from excel_package.stock_data.live_stock import LiveStock
from excel_package.utills.utill_setting import LIVE_STOCK, DEFAULT_START_D, START_DATE_CELL, END_DATE_CELL, \
    DEFAULT_END_D


class GraphSheet(Sheet, ABC):
    def __init__(self, name, index, file_name):
        Sheet.__init__(self, name, index, file_name)
        self.date_start_before = DEFAULT_START_D
        self.date_end_before = DEFAULT_END_D
        self.graph_type = 'line_markers'

    def pre_run(self):
        pass

    def run_sheet(self):
        pass

    def get_all_tk(self):
        tk_list = []
        for index in range(30):
            tk_list += [xw.Book(self.file_name).sheets[LIVE_STOCK].range(f"A{index + INDEX_TO_START_STOCK_VAL}").value]
        return tk_list

    def get_date_start(self):
        date_start_excel = xw.Book(self.file_name).sheets[LIVE_STOCK].range(START_DATE_CELL).value
        if not date_start_excel:
            return DEFAULT_START_D
        else:
            date_start_excel = date_start_excel.strftime("%Y-%m-%d")
        return date_start_excel

    def get_date_end(self):
        date_end_excel = xw.Book(self.file_name).sheets[LIVE_STOCK].range(END_DATE_CELL).value
        if not date_end_excel:
            return DEFAULT_END_D
        else:
            date_end_excel = date_end_excel.strftime("%Y-%m-%d")
        return date_end_excel

    def delete_all_charts(self):
        xw.Book(self.file_name).sheets[self.name].clear_contents()
        for chart in xw.Book(self.file_name).sheets[self.name].charts:
            chart.delete()

    def global_graph(self):
        tk_list = self.get_all_tk()
        if self.get_date_start() != self.date_start_before or self.get_date_end() != self.date_end_before:
            self.delete_all_charts()
            self.date_start_before = self.get_date_start()
            self.date_end_before = self.get_date_end()

        try:
            for i, chart in enumerate(xw.Book(self.file_name).sheets[self.name].charts):
                if not chart.api[1].ChartTitle.Text in tk_list:
                    chart.delete()
                    xw.Book(self.file_name).sheets[self.name].range((1000, ((i + 1) * 2)),
                                                                    (3000, ((i + 2) * 2))).clear_contents()
        except Exception as e:
            print(str(e))
            traceback.print_exc()
        index_t = 1
        for ticker in tk_list:
            if ticker:
                self.add_chart_graph_sheet(stock_name=ticker, index=index_t)
            index_t += 1

    def add_chart_graph_sheet(self, stock_name, index):
        try:
            for chart in xw.Book(self.file_name).sheets[self.name].charts:
                if chart.api[1].ChartTitle.Text in stock_name:
                    return
            sht = xw.Book(self.file_name).sheets[self.name]
            row_num = 1000
            col_num = index * 2
            sht.range((row_num, col_num)).value = self.get_df(stock_name)
            chart = sht.charts.add()
            chart.set_source_data(sht.range((row_num, col_num)).expand())
            chart.chart_type = self.graph_type
            chart.api[1].SetElement(2)  # Place chart title at the top
            chart.api[1].ChartTitle.Text = stock_name  # Change text of the chart title
            chart.width = 1000
            chart.top = 211.0 * (index - 1)
        except Exception as e:
            print(f"fail add_chart_graph_sheet {e}")
            return

    def get_df(self, stock_name):
        stock_t = LiveStock(stock_name)
        return stock_t.get_data_history(start_d=self.get_date_start(),
                                        end_d=self.get_date_end())["Adj Close"]
