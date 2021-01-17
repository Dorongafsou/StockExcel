from excel_package import Excel
from excel_package.Excel import Excel

EXCEL_NAME = "dorong.xlsx"


def main():
    excel = Excel(EXCEL_NAME)
    excel.run()


if __name__ == '__main__':
    main()
