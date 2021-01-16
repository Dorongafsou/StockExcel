from excel_package.workbook.workbook import WorkBook


class Excel(object):
    def __init__(self, name_excel):
        self._name_excel = name_excel

    def run(self):
        workbook = WorkBook(self._name_excel)
        workbook.pre_run()
        workbook.real_time()
