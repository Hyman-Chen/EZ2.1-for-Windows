import tkinter as tk
from pathlib import Path
import os
import sys

from ttkbootstrap.dialogs import Messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd

import data
import data_entry

BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
IMG_PATH = Path(BASE_PATH) / 'data' / 'image'

SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'
tempDATA_PATH = Path(BASE_PATH) / 'data' / 'temp'


class AppendcoFrame(ttk.Frame):
    '''添加单位'''

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.root = master
        self.root.showframe.destroy()
        self.root.showframe = self
        self.pack(side=TOP, fill=BOTH, expand=True)
        # 图片
        self.images = [
            ttk.PhotoImage(file=IMG_PATH / 'icons8_add_folder_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'add_24px.png')
        ]
        # 初始化主题
        setting_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))
        theme_name = setting_df['current_theme'][0]
        style = ttk.Style()
        style.theme_use(theme_name)

        self.create_paget()
        self.create_pageb()

    def create_paget(self):
        '''1 paget'''
        self.framet = ttk.LabelFrame(self, text='单位管理', bootstyle=SUCCESS)
        self.framet.pack(side=TOP, fill=BOTH, expand=False, padx=10, pady=10)
        # 1.1 添加一级单位
        ttk.Button(
            self.framet,
            text='添加一级单位',
            bootstyle=('outline', INFO),
            width=13,
            image=self.images[1],
            command=self.add_co,
            compound=LEFT
        ).pack(side=TOP, fill=X, padx=20, pady=10)
        mid_frame = ttk.Frame(self.framet)
        mid_frame.pack(side=TOP, fill=X)
        # 1.2 添加二级单位
        ttk.Button(
            mid_frame,
            text='添加二级单位',
            bootstyle=SUCCESS,
            width=11,
            padding=(0, 9),
            command=self.add_second_unit
        ).pack(side=LEFT, fill=X, expand=True, padx=20, pady=0)
        # 1.3 删除一级（或二级）单位
        ttk.Button(
            mid_frame,
            text='   删除一级（或二级）单位',
            bootstyle=DANGER,
            width=13,
            padding=(0, 9),
            command=self.dele
        ).pack(side=LEFT, fill=X, expand=True, padx=20, pady=0)
        # 1.#
        ttk.Label(
            self.framet,
            text="功能说明：二级单位必须全部包含考勤表中的所有二级单位，\n未列出的二级单位不参与人数统计。"
        ).pack(side=TOP, fill=X, padx=20, pady=10)

    def create_pageb(self):
        '''2 pageb'''
        # 放在frameb上的笔记本Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        # 单位管理
        self.frame1 = ttk.Frame(self.notebook)
        self.notebook.add(self.frame1, text="单位管理")
        # 插入单位树
        self.treeScroly = tk.Scrollbar(self.frame1)
        self.treeScroly.pack(side="right", fill="y")
        self.tableColumns1 = ['单位']
        self.treeview1 = ttk.Treeview(self.frame1, selectmode="extended",
                                      yscrollcommand=self.treeScroly.set,
                                      show='tree headings',
                                      height=40, bootstyle=SUCCESS)

        self.treeview1.pack(expand=TRUE, fill=BOTH)
        self.treeScroly.config(command=self.treeview1.yview)
        try:
            df = pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'))
            for head in list(df.columns):
                myidx = self.treeview1.insert('', 'end', text=head, value=data, iid=head)
                for value in df[head]:
                    if str(value).upper() != 'NAN' and str(value).upper() != '@NAN':
                        self.treeview1.insert(myidx, 'end', text=value, value=data, iid='@' + head + "," + str(value))
        except:
            pass

    def dele(self):
        '''删除单位'''
        try:
            self.selected_item = self.treeview1.selection()[0]
            if not '@' in self.selected_item:
                pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv')).drop(str(self.selected_item), axis=1).to_csv(
                    Path.joinpath(SAVE_PATH, 'co_dict.csv'), index=False)
                self.fresh(self.treeview1)
            else:
                df = pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'))
                head = str(self.selected_item.split(",")[0])[1:]
                varlist = list(df[head])
                varlist.remove(self.selected_item.split(",")[1])
                varlist.append('')
                df[head] = varlist
                last_row = df.loc[len(df) - 1]
                for i in last_row:
                    if i != '':
                        break
                    elif i == last_row[-1]:
                        df.drop(len(df) - 1)

                df.to_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'), index=False)

                self.fresh(self.treeview1)
                Messagebox.show_info_success(('已删除%s!' % self.selected_item.split(",")[1]), '温馨提示')
        except:
            Messagebox.show_info_success('请选择需要删除的一级（或二级）单位！', '温馨提示')

    def fresh(self, treeview):
        '''刷新页面'''
        for _ in map(treeview.delete, treeview.get_children("")):
            pass
        try:
            df = pd.read_csv(Path.joinpath(SAVE_PATH, 'co_dict.csv'))
            for head in list(df.columns):
                myidx = self.treeview1.insert('', 'end', text=head, value=data, iid=head)
                for value in df[head]:
                    if str(value).upper() != 'NAN' and str(value).upper() != '@NAN':
                        self.treeview1.insert(myidx, 'end', text=value, value=data, iid='@' + head + "," + str(value))
        except:
            pass

    def add_co(self):
        '''添加单位'''
        data_entry.DataEntryForm_1(self)
        self.fresh(self.treeview1)

    def add_second_unit(self):
        '''添加二级单位'''
        data_entry.DataEntryForm_2(self)
        self.fresh(self.treeview1)
