# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import build_stock
from build_stock import BuildStock
import sys


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.





# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # print_hi('PyCharm')
    bs = BuildStock()
    bs.load_workbook()
    name_option = bs.get_stock_list_from_transaction()
    bs.write_items_in_stock_master_file1(name_option)

    print(name_option)

