from datetime import datetime

import xlsxwriter


class ExcelSaver(object):
    def __init__(self):
        self.f = xlsxwriter.Workbook(f'collating documents {str(datetime.now())}.xlsx')
        self.worksheet1 = self.f.add_worksheet('sheet1')  # 只用一个sheet
        self.worksheet1.write_row("A1", ['案件文书名称', '审结年份', '所属省份', '关键词', '关键词所在段落'])
        self.now_row = 2

    def add_row(self, document_name, completion_year, completion, key, para):
        self.worksheet1.write_row(f"A{self.now_row}", [document_name, completion_year, completion, key, para])
        self.now_row += 1

    def save_xlsx(self):  # __del__处理的话self.f 已经被关掉了会报错
        print('文件处理完毕，正在导出.xlsx')
        self.f.close()