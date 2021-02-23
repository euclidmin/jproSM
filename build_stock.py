import openpyxl
import sys


def funcname():
    return sys._getframe(1).f_code.co_name + "()"


class BuildStock:
    def __init__(self):
        self.stock_master_file = None
        self.transaction_file = None
        self.stock_master_workbook = None
        self.transaction_workbook = None
        self.stock_list_from_transaction = None

        self.name_option_cells = None
        self.name_option_dict = {}
        print(funcname())

    def load_workbook(self):
        print(funcname())

        def _open_transaction_file():
            print(funcname())
            return openpyxl.load_workbook('스마트스토어_주문조회_20210102_1027.xlsx')

        def _open_stock_master_file():
            print(funcname())
            return openpyxl.load_workbook('stock.xlsx')

        self.stock_master_workbook = _open_stock_master_file()
        self.transaction_workbook = _open_transaction_file()


    def write_items_in_stock_master_file(self, items):
        sheet = self.stock_master_workbook.active
        sheet.title = "재고리스트"

        for i, item_name in enumerate(items.keys()):
            sheet.cell(row=i+1, column=1, value=item_name)
            sheet.cell(row=i+1, column=2, value=items[item_name])

        self.stock_master_workbook.save("stock.xlsx")


    def write_items_in_stock_master_file1(self, items):
        stock_master_item_value = self.get_stock_list_from_stock_master_file()

        sheet = self.stock_master_workbook.active
        sheet.title = "재고리스트"
        sheet.cell(row=1, column=1, value="판매품명")
        sheet.cell(row=1, column=2, value="재고개수")
        for i, item_name in enumerate(items.keys()):
            if item_name not in stock_master_item_value:
                print(item_name)
                sheet.cell(row=i+2, column=1, value=item_name)
                sheet.cell(row=i+2, column=2, value=items[item_name])

        self.stock_master_workbook.save("stock.xlsx")



    def get_stock_list_from_transaction(self):
        def get_name_and_option_cells_from_transaction():
            transaction_worksheet = self.transaction_workbook['주문조회']
            max_cnt = transaction_worksheet.max_row
            ret = transaction_worksheet['G2':'H' + str(max_cnt)]
            # print(ret)
            return ret

        def make_name_and_option(cells):
            name_option_dict = {}
            for index in range(0, len(cells)):
                name = cells[index][0].value
                option = cells[index][1].value
                name_option = name + ' | ' + option
                # 재고 아이템별 중복 없이 dict에 초기값 0과 함께 저장
                name_option_dict[name_option] = 0
                # print(str(index) + ' ' + name_option)
            return name_option_dict

        cells = get_name_and_option_cells_from_transaction()
        ret = make_name_and_option(cells)

        return ret

    def get_stock_list_from_stock_master_file(self):
        stock_master_file_worksheet = self.stock_master_workbook['재고리스트']
        max_cnt = stock_master_file_worksheet.max_row
        cells = stock_master_file_worksheet['A2':'B' + str(max_cnt)]
        # print(ret)

        item_name_value_dict = {}
        for index in range(0, len(cells)):
            item_name = cells[index][0].value
            value = cells[index][1].value
            # 재고 아이템별 중복 없이 dict에 초기값 0과 함께 저장
            item_name_value_dict[item_name] = value
            # print(str(index) + ' ' + name_option)
        return item_name_value_dict







