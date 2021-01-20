import yfinance as yf
import xlwings as xw
from datetime import datetime
from pandas_datareader import data
from pandas_datareader import data as pdr
# a = data.get_quote_yahoo("NVDA").to_dict()
# print(a)
sht = xw.Book("dorong.xlsx").sheets["Stock Interval"]
cells = sht.cells
cells.autofit()
cells.api.Borders.Weight = 3
cells.api.Font.Bold = True

# rng = sht.range("A4")
# rng.value = ""
# rng.value = "linear_benefityy"
# rng_val = rng.api.Validation
# rng_val.add(Type=3, AlertStyle=1, Operator=1, Formula1="linear_benefit,linear_cost,sigmoid_benefit,sigmoid_cost")
# sht.api.Shapes("Drop Down 1").ControlFormat.Value = 1
# print()
# rng.api.validation.Formula1[1:]

# sht.range("A4").api.Validation.add(3,1,3,"linear_benefit,linear_cost,sigmoid_benefit,sigmoid_cost")

# data = yf.download(tickers="NVDA", period="1d", interval="15m")
# print(data.shape[1])
# sht.range((1,5)).value = data
# sht.range('A5').value = data.tail(60)
