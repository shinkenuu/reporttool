import abc
from openpyxl import Workbook

REPORT_ROOT_PATH = '/mnt/jatobrfiles/Weaver/reports/'


class Report(metaclass=abc.ABCMeta):
    def __init__(self):
        self.matrix = [[]]
        self.wb = Workbook()
        self.ws = self.wb.active

    class Header(metaclass=abc.ABCMeta):
        def __init__(self, header_name: str, number_format: str, offset: int):
            self.header_name = header_name
            self.number_format = number_format
            self.offset = offset

    @abc.abstractmethod
    def generate_report(self):
        pass

    def write_to_disc(self):
        self.wb.save('{}{}.xlsx'.format(REPORT_ROOT_PATH, str(self).replace('.', '/')))

    # TODO code
    def write_matrix_to_xl(self):
        pass
