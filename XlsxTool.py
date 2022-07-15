from openpyxl import Workbook
import os
from typing import Union

first_line = ['id', 'title', 'method', 'url', 'pre_sql', 'req_data', 'extract', 'assert_list']


class XlsxTool:
    def __init__(self, filename: str):
        """
        生成首行
        :param filename:只需要文件名即可
        """
        if filename in os.listdir(os.getcwd() + '/XlsxCase'):
            raise Exception("文件已存在")
        self.book = Workbook()
        self.sheet = self.book.active
        self.filename = filename
        self.write(1, first_line)
        ...

    def write(self, id_: Union[str, int], values: Union[list, None]):
        A = 65
        for x in range(len(values)):
            self.sheet[chr(A + x) + str(id_)] = values[x]
        self.book.save(os.getcwd() + '/XlsxCase' + '/' + self.filename)


if __name__ == "__main__":
    XlsxTool('1.xlsx').write(2, [1, 2, 3])
