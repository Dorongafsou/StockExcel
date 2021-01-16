import pythoncom

from base.excel_feeder import ExcelFeeder

EXCEL_NAME = "dorong.xlsx"


def main():
    ExcelFeeder(EXCEL_NAME)


if __name__ == '__main__':
    main()
