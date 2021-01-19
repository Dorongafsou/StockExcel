import yfinance as yf
import xlwings as xw
from datetime import datetime
from pandas_datareader import data
from pandas_datareader import data as pdr

sht = xw.Book("dorong.xlsx").sheets["Live Stock"]
data = yf.download(tickers="NVDA", period="4d", interval="5m")
print(type(data.tail()))
print(data.tail(3))
sht.range('A5').value = data.tail(60)
