# -*- coding: UTF-8 -*-
import tkinter as tk
from tkinter import filedialog
from xmindparser import xmind_to_dict

from excel.excel_data import EXCEL_DATA

# 宽度
width = 760
# 高度
height = 600


class XMIND_GUI(object):

    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        self.init_window_name.title("Xmind解析")
        # self.init_window_name.configure(bg="red")
        screenwidth = self.init_window_name.winfo_screenwidth()  # 获取显示屏宽度
        screenheight = self.init_window_name.winfo_screenheight()  # 获取显示屏高度
        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)  # 设置窗口居中参数
        self.init_window_name.geometry(size)  # 让窗口居中显示

        self.but_upload = tk.Button(self.init_window_name, text=u'上传xmind文件', command=self.upload_files,
                                    bg='#dfdfdf')
        self.but_upload.grid(row=0, column=0)
        # 显示文件路径
        self.text = tk.Text(self.init_window_name, width=55, height=10, bg='#f0f0f0')
        self.text.grid(row=1, column=0)

        def gen_execl_file():
            lines = self.text.get(1.0, tk.END)

            # 分隔成每行
            for line in lines.splitlines():
                if line == '':
                    break
                self.import_xmind(line)
                self.execl_path.insert(tk.END, line.replace('xmind', 'xls') + '\n')
                self.execl_path.update()

        self.button = tk.Button(self.init_window_name, text=u"生成Excel", width=10, command=gen_execl_file,
                                bg='#dfdfdf')
        self.button.grid(row=3, column=0)

        # 显示文件路径
        self.execl_path = tk.Text(self.init_window_name, width=55, height=10, bg='#f0f0f0')
        self.execl_path.grid(row=4, column=0)

    def upload_files(self):
        """上传多个文件，并插入text中"""
        select_files = filedialog.askopenfilenames(title=u"可选择1个或多个文件")
        for file in select_files:
            self.text.insert(tk.END, file + '\n')
            self.text.update()

    def import_xmind(self, file_path):
        sheets = xmind_to_dict(file_path)
        for sheet in sheets:
            xmind_result = self.xmind_result(sheet['topic'])
            excel = EXCEL_DATA(xmind_result, file_path)
            excel.gen_execl()

    @staticmethod
    def xmind_result(sheet):
        system_name = sheet['title']
        xmind_result = []
        for a in sheet['topics']:
            case_name = a['title']
            for b in a['topics']:
                operating_steps = b['title']
                for c in b['topics']:
                    check_points = c['title']
                    row_data = []
                    row_data.append(system_name)
                    row_data.append(case_name)
                    row_data.append(operating_steps)
                    row_data.append(check_points)
                    xmind_result.append(row_data)
        return xmind_result
