import tkinter as tk
import csv
from pathlib import Path
import os
import sys

from ttkbootstrap.dialogs import Messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd

import initTool
import Judge

# 初始化源文件所在位置
BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
# 图片文件位置
IMG_PATH = Path(BASE_PATH) / 'data' / 'image'
# 应用数据保存位置
SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'


class ShujukuFrame(ttk.Frame):
    '''数据库'''

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.root = master
        self.root.showframe.destroy()
        self.root.showframe = self
        self.root.showframe.pack(side=TOP, fill=BOTH, expand=True)

        # 初始化主题
        setting_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))
        theme_name = setting_df['current_theme'][0]
        style = ttk.Style()
        style.theme_use(theme_name)
        self.create_paget()
        self.create_pageb()

    def create_paget(self):
        '''1 paget'''
        self.framet = ttk.LabelFrame(self, text='数据库', bootstyle=SUCCESS)
        self.framet.pack(side=TOP, fill=BOTH, expand=False, padx=10, pady=10)
        try:
            self.co_list = list(pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv')).columns)
        except:
            self.co_list = []
            Messagebox.show_info_success("请先添加单位！", "温馨提示")
        # 1.1 单位
        row1 = ttk.Frame(self.framet)
        row1.pack(side=TOP, fill=X, expand=TRUE)
        row1.columnconfigure(index=0, weight=0)
        row1.columnconfigure(index=1, weight=1)
        ttk.Label(
            master=row1, text='单位', width=5
        ).grid(
            row=0, column=0, sticky=E, padx=(10, 0), pady=5
        )
        self.combobox = ttk.Combobox(row1, values=self.co_list, width=200, style=SUCCESS)
        self.combobox.current = ''
        self.combobox.grid(
            row=0, column=1, sticky=W, padx=(0, 10), pady=5
        )

        # 1.2姓名
        row2 = ttk.Frame(self.framet)
        row2.pack(side=TOP, fill=X, expand=TRUE)
        for index in [0, 2, 4]:
            row2.columnconfigure(index=index, weight=0)
        for index in [1, 3, 5]:
            row2.columnconfigure(index=index, weight=1)

        ttk.Label(master=row2, text='姓名', width=5).grid(
            row=0,
            column=0,
            sticky="e",
            columnspan=1,
            padx=(10, 0), pady=5
        )
        self.add_name = tk.StringVar()
        self.add_name.set('')
        self.entry2 = ttk.Entry(
            master=row2,
            textvariable=self.add_name,
            width=15,
            style=SUCCESS
        )
        self.entry2.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 0),
            pady=5,
            columnspan=1
        )
        # 1.3 工种类别
        ttk.Label(
            master=row2,
            text='工种\n类别',
            width=5,
        ).grid(row=0, column=2, sticky="e", padx=(10, 0), pady=5)
        self.option_menu_list1 = ["", "特殊工种", "普通工种", ""]
        self.GZLB_Var = tk.StringVar(value=self.option_menu_list1[0])
        self.OpMenu1 = ttk.OptionMenu(
            row2,
            self.GZLB_Var,
            *self.option_menu_list1,
            bootstyle=('outline', SUCCESS),
        )
        self.OpMenu1.config(width=15)
        self.OpMenu1.grid(
            row=0,
            column=3,
            sticky="ew",
            padx=(0, 0), pady=5,
            columnspan=1
        )
        # 1.4 年龄类别
        ttk.Label(
            master=row2,
            text='年龄\n类别',
        ).grid(
            row=0,
            column=4,
            padx=(10, 5), pady=5,
            sticky="e"
        )
        self.option_menu_list2 = ["", "超龄", "未超龄", ""]
        self.NLLB_Var = tk.StringVar(value=self.option_menu_list2[0])
        self.OpMenu2 = ttk.OptionMenu(
            row2,
            self.NLLB_Var,
            *self.option_menu_list2,
            bootstyle=('outline', SUCCESS),
        )
        self.OpMenu2.config(width=15)
        self.OpMenu2.grid(
            row=0,
            column=5,
            sticky="ew",
            padx=(0, 10), pady=5,
            columnspan=1
        )

        # 1.5 日期表
        row3 = ttk.Frame(self.framet)
        row3.pack(side=TOP, fill=X, expand=TRUE)
        for index in [0, 2]:
            row3.columnconfigure(index=index, weight=0)
        for index in [1, 3, 4, 5, 6]:
            row3.columnconfigure(index=index, weight=1)

        ttk.Label(
            master=row3,
            text='培训\n日期',
            width=5
        ).grid(row=0, column=0, sticky="e", padx=(10, 0), pady=5)
        self.data_Table = ttk.DateEntry(
            row3,
            bootstyle=SUCCESS,
            width=9
        )

        self.data_Table.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 10), pady=5,
            columnspan=1
        )
        #  1.6 身份
        ttk.Label(
            master=row3,
            text='身份',
            width=5,
        ).grid(
            row=0,
            column=2,
            padx=(5, 0), pady=5,
            sticky="e"
        )
        self.option_menu_list5 = ["", "作业人员", "管理人员", ""]
        self.SF_Var = tk.StringVar(value=self.option_menu_list5[0])
        self.OpMenu3 = ttk.OptionMenu(
            row3,
            self.SF_Var,
            *self.option_menu_list5,
            bootstyle=('outline', SUCCESS),
        )
        self.OpMenu3.grid(
            row=0,
            column=3,
            sticky="ew",
            padx=(0, 0), pady=5,
            columnspan=1
        )
        # 1.7 重置
        ttk.Button(
            row3,
            text='重置',
            bootstyle=SUCCESS,
            command=self.chongzhi
        ).grid(
            row=0,
            column=4,
            sticky="ew",
            padx=(10, 10),
            pady=5
        )
        # 1.8 查询
        ttk.Button(
            row3,
            text='查询',
            bootstyle=SUCCESS,
            command=lambda: self.chaxun(
                self.data_Table.entry.get(),  # 日期
                self.combobox.get(),  # 单位
                self.entry2.get(),  # 姓名
                self.GZLB_Var.get(),  # 工种类别
                self.NLLB_Var.get(),  # 年龄类别
                self.SF_Var.get(),  # 身份
            )
        ).grid(
            row=0,
            column=5,
            sticky="ew",
            padx=(0, 10),
            pady=5
        )
        # 1.9 删除
        ttk.Button(
            row3,
            text='删除',
            bootstyle=DANGER,
            command=self.dele
        ).grid(
            row=0,
            column=6,
            sticky="ew",
            padx=(0, 10), pady=5
        )

    def create_pageb(self):
        '''2 pageb'''
        # 放在frameb上的笔记本Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=BOTTOM, fill=BOTH, expand=True, padx=10, pady=(0, 10))
        # 名单总表
        self.frame1 = ttk.Frame(self.notebook)
        self.notebook.add(self.frame1, text="名单总表")
        self.tableValues1 = initTool.Files().csv_to_treeviewlist(SAVE_PATH, 'prjData.csv')[1:]
        self.tableColumns1 = ['序号', '单位', '姓名', '工种', '性别', '年龄', '身份证号', '家庭住址', '手机', '培训日期', '备注']
        A = initTool.Tab1_2(self.frame1, self.tableValues1, self.tableColumns1)
        self.treeview1 = A.treeview

    def chaxun(self, date, unit, name, GZLB, NLLB, SF):
        '''
        根据
        日期,单位,姓名,工种类别,年龄类别,身份
        查找对应培训数据
        '''
        df = pd.read_csv(Path.joinpath(SAVE_PATH, 'prjData.csv'), encoding='utf-8')
        with open(Path.joinpath(SAVE_PATH, 'chaxunData.csv'), 'w', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(
                ['序号', '单位', '姓名', '工种', '性别', '年龄', '身份证号', '家庭住址', '手机', '培训日期', '备注'])
            for row in range(df.shape[0]):
                row_val = list(df.loc[row])
                if Judge.judge(row_val, date, unit, name, GZLB, NLLB, SF):
                    writer.writerow(df.iloc[row, [i for i in range(df.shape[1])]])
            f.close()
        initTool.Files().fresh_csv(SAVE_PATH, 'chaxunData.csv')
        self.tableValues1 = initTool.Files().csv_to_treeviewlist(SAVE_PATH, 'chaxunData.csv')
        self.fresh(self.treeview1, self.tableValues1)

    def chongzhi(self):
        '''重置'''
        self.destroy()
        self.root.showframe = ShujukuFrame(self.root)
        self.root.showframe.pack(side=TOP, fill=BOTH, expand=True)

    def dele(self):
        '''删除某行'''
        try:
            a = int(self.treeview1.selection()[0])
            for selection in self.treeview1.selection():
                self.rowval = int(selection)
                pd.read_csv(Path.joinpath(SAVE_PATH, 'prjData.csv')).drop([a - 1]).to_csv(
                    Path.joinpath(SAVE_PATH, 'prjData.csv'), index=False)
                initTool.Files().fresh_csv(SAVE_PATH, 'prjData.csv')
                self.treeview1.delete(selection)
                self.fresh(self.treeview1, initTool.Files().csv_to_treeviewlist(SAVE_PATH, 'prjData.csv'))
        except:
            Messagebox.show_info_success('请选择需要删除的人员！', '温馨提示')

    def fresh(self, treeview, tableValues):
        '''刷新页面'''
        for _ in map(treeview.delete, treeview.get_children("")):
            pass
        try:
            row = 1
            for data in tableValues[1:]:
                treeview.insert('', 'end', text=row, value=data, iid=str(row))
                row += 1
        except:
            pass
