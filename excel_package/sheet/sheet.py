from abc import ABC, abstractmethod
from functools import total_ordering


@total_ordering
class Sheet(ABC):
    def __init__(self, name, index, file_name):
        self.name = name
        self.index = index
        self._xlwing_sheet = None
        self.file_name = file_name

    def __str__(self):
        return self.name

    def __lt__(self, other):
        return self.index < other.index

    def __eq__(self, other):
        return self.index == other.index

    @property
    def xlwing_sheet(self):
        return self._xlwing_sheet

    @xlwing_sheet.setter
    def xlwing_sheet(self, sheet):
        sheet.name = self.name
        self._xlwing_sheet = sheet

    @abstractmethod
    def pre_run(self):
        pass

    @abstractmethod
    def run_sheet(self):
        pass
