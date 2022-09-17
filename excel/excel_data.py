# -*- coding: UTF-8 -*-

import xlwt


class EXCEL_DATA(object):

    def __init__(self, xmind_result, execl_path):
        self.execl_path = execl_path
        self.execl_data = xmind_result

    def gen_execl(self):
        workbook = xlwt.Workbook(encoding="utf-8")  # 实例化book对象
        sheet = workbook.add_sheet("xlwt方法")  # 生成sheet
        # 写入标题
        for col, column in enumerate(["系统名称", "用例名称", "操作步骤", "检查点"]):
            sheet.write(0, col, column)
        # 写入每一行
        for row, data in enumerate(self.execl_data):
            for col, col_data in enumerate(data):
                sheet.write(row + 1, col, col_data)
        gen_path = self.execl_path.replace('xmind', 'xls')
        workbook.save(gen_path)
