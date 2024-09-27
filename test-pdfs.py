import pprint
from utils.read_excel import Excel_file

from typing import List


def test():
    e = Excel_file("./utils/10008361744.xlsm")
    index = e.create_index(create_json=True)
    pprint.pprint(index)


test()
