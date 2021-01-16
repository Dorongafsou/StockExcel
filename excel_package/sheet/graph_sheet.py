from abc import ABC

from excel_package.sheet.sheet import Sheet


class GraphSheet(Sheet, ABC):
    def pre_run(self):
        pass

    def run_sheet(self):
        pass
