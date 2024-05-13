import tkinter as tk
from pathlib import Path
import os
import sys
import csv

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import pandas as pd

BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'


class DataEntryForm_1(ttk.Window):
    '''添加一级单位输入框'''

    def __init__(self, master):
        super().__init__(title='添加')
        self.root = master
        self.unit = ttk.StringVar(value="")

        self.create_form_entry("单位名称", self.unit)
        self.create_buttonbox()

    def create_form_entry(self, label, variable):
        '''创建输入框'''
        container = tk.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)

        lbl = tk.Label(master=container, text=label, width=10)
        lbl.pack(side=LEFT, padx=5)

        self.ent = tk.Entry(master=container, textvariable=variable)
        self.ent.pack(side=LEFT, padx=5, fill=X, expand=YES)

    def create_buttonbox(self):
        '''创建按钮'''
        self.container = tk.Frame(self)
        self.container.pack(fill=X, expand=YES, pady=(15, 10))

        self.sub_btn = tk.Button(
            master=self.container,
            text="添加",
            command=self.on_submit,
            width=6,
        )
        self.sub_btn.pack(side=RIGHT, padx=5)
        self.sub_btn.focus_set()

        cnl_btn = tk.Button(
            master=self.container,
            text="取消",
            command=self.on_cancel,
            width=6,
        )
        cnl_btn.pack(side=RIGHT, padx=5)

    def on_submit(self):
        '''提交函数'''
        self.co_name = self.ent.get()
        try:
            co_dict_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'))
        except:
            co_dict_df = pd.DataFrame()
        if (not self.co_name in co_dict_df.columns) and self.co_name != '':
            init_list = [None, self.co_name]
            for i in range(co_dict_df.shape[0] - 2):
                init_list.append('')
            co_dict_df[self.co_name] = init_list
            co_dict_df.to_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'), index=False)
            self.on_cancel()
        else:
            Messagebox.show_info_success('输入有误!', '温馨提示')

    def on_cancel(self):
        '''取消按钮'''
        self.destroy()


class DataEntryForm_2(ttk.Window):
    '''添加二级单位'''

    def __init__(self, master):
        super().__init__(title='二级单位添加')
        self.root = master
        self.second_unit = ttk.StringVar(value="")
        # 一级输入框
        container1 = tk.Frame(self)
        container1.pack(fill=X, expand=YES, pady=5)
        lbl1 = tk.Label(master=container1, text="一级单位", width=10)
        lbl1.pack(side=LEFT, padx=5)
        try:
            self.units = list(pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv')).columns)
        except:
            self.units = []

        self.ent1 = ttk.Combobox(master=container1, values=self.units)
        self.ent1.pack(side=LEFT, padx=5, fill=X, expand=YES)
        self.ent1.current = ''
        # 二级输入框
        container2 = tk.Frame(self)
        container2.pack(fill=X, expand=YES, pady=5)
        lbl2 = tk.Label(master=container2, text="二级单位", width=10)
        lbl2.pack(side=LEFT, padx=5)
        self.ent2 = ttk.Entry(master=container2, textvariable=self.second_unit)
        self.ent2.pack(side=LEFT, padx=5, fill=X, expand=YES)

        self.create_buttonbox()

    def create_buttonbox(self):
        '''创建按钮'''
        self.container = tk.Frame(self)
        self.container.pack(fill=X, expand=YES, pady=(15, 10))

        self.sub_btn = tk.Button(
            master=self.container,
            text="添加",
            command=self.on_submit,
            width=6,
        )
        self.sub_btn.pack(side=RIGHT, padx=5)
        self.sub_btn.focus_set()

        cnl_btn = tk.Button(
            master=self.container,
            text="取消",
            command=self.on_cancel,
            width=6,
        )
        cnl_btn.pack(side=RIGHT, padx=5)

    def on_submit(self):
        '''提交函数'''
        self.co_name = self.ent1.get()
        self.second_unit_name = self.ent2.get()
        try:
            co_dict_df = pd.DataFrame(pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv')))
            if str(co_dict_df[self.co_name][len(co_dict_df) - 1]) != 'nan':
                addlist = []
                for i in range(int(co_dict_df.shape[1])): addlist.append('')
                for col in co_dict_df.columns:
                    if str(col) == self.co_name:
                        addlist[list(co_dict_df.columns).index(col)] = self.second_unit_name
                        break
                if (self.co_name in co_dict_df.columns) and self.second_unit_name != '':
                    with open(Path.joinpath(SAVE_PATH, 'co_dict.csv'), 'a', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        writer.writerow(addlist)
                        f.close()

                    self.root.fresh(self.root.treeview1)
                    self.on_cancel()
                else:
                    Messagebox.show_info_success('输入有误!', '温馨提示')
            else:
                temlist = list(co_dict_df[self.co_name])
                for i in range(1, len(temlist)):
                    if str(temlist[i]) == 'nan':
                        temlist[i] = self.second_unit_name
                        break
                co_dict_df[self.co_name] = temlist
                co_dict_df.to_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'), index=False)
                self.root.fresh(self.root.treeview1)
                self.on_cancel()
        except:
            Messagebox.show_info_success('请先添加单位!', '温馨提示')
            self.on_cancel()

    def on_cancel(self):
        '''取消函数'''
        self.destroy()


if __name__ == "__main__":
    app = ttk.Window("Data Entry", "flatly", resizable=(False, False))
    # DataEntryForm(app)
    app.mainloop()
