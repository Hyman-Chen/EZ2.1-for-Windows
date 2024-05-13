import csv
from pathlib import Path
import os
import sys
from tkinter import filedialog

from ttkbootstrap.dialogs import Messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd

import main

# 初始化源文件地址
BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
# 图片地址
IMG_PATH = Path(BASE_PATH) / 'data' / 'image'
# 应用数据文件地址
SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'
# 初始化四个保存位置
DATA_PATH = Path(BASE_PATH)
Mer_PATH = Path(BASE_PATH)
KQ_PATH = Path(BASE_PATH)
RS_PATH = Path(BASE_PATH)


class SettingFrame(ttk.Frame):
    '''2 设置栏'''

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.root = master
        self.root.showframe.destroy()
        self.root.showframe = self
        self.pack(side=LEFT, fill=BOTH, expand=TRUE)
        # 图片
        self.images = [
            ttk.PhotoImage(file=IMG_PATH / 'fanhui_32px.png')
        ]
        # 主题元组
        self.theme_names = (
            'cosmo',
            'flatly',
            'litera',
            'minty',
            'lumen',
            'sandstone',
            'yeti',
            'pulse',
            'united',
            'morph',
            'journal',
            'darkly',
            'superhero',
            'solar',
            'cyborg',
            'vapor',
            'simplex',
            'cerculean'
        )
        # 设置主题
        setting_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))
        theme_name = setting_df['current_theme'][0]
        style = ttk.Style()
        style.theme_use(theme_name)

        # 2.1 返回主页
        self.top_frame = ttk.Frame(self)
        self.top_frame.pack(side=TOP, fill=X, expand=False)
        self.back = ttk.Button(
            self.top_frame,
            text='返回主页',
            bootstyle=(LINK),
            padding=(0, 0, 5, 0),
            command=self.fanhui,
            image=self.images[0],
            compound=LEFT)
        self.back.pack(side=LEFT)
        # 2.2 设置标签
        self.lab_frame = ttk.Frame(self)
        self.lab_frame.pack(side=TOP, fill=X, expand=False)
        ttk.Label(self.lab_frame, text="设置", font=("黑体", 12), anchor=CENTER).pack(fill=X)

        # 2.3 左边的标签栏
        self.left_frame = ttk.Frame(self)
        self.left_frame.pack(side=LEFT, fill=BOTH, expand=False, padx=(20, 10))
        # 2.3.# 标签1
        ttk.Label(self.left_frame, text="用户与项目", font=("黑体", 10)).pack(fill=X)
        # 2.3.1 用户名
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        ttk.Button(self.left_frame, text='用户名', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.3.2 主题
        ttk.Button(self.left_frame, text='主   题', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.3.3 项目名称
        ttk.Button(self.left_frame, text='项目名称', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.3.4 项目简称
        ttk.Button(self.left_frame, text='项目简称', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.3.# 标签2
        ttk.Label(self.left_frame).pack(fill=X)
        ttk.Label(self.left_frame, text="文件管理", font=("黑体", 10)).pack(fill=X)
        # 2.3.5 名单汇总路径
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        ttk.Button(self.left_frame, text='名单汇总保存位置', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.3.6 总表保存路径
        ttk.Button(self.left_frame, text='导出总表保存位置', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.2.7 现场人数文件路径
        ttk.Button(self.left_frame, text='现场人数保存位置', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.2.8 实时人数文件路径
        ttk.Button(self.left_frame, text='实时人数保存位置', bootstyle=(DISABLED, PRIMARY)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)

        # 2.4 用户输入
        self.right_frame = ttk.Frame(self)
        self.right_frame.pack(side=RIGHT, fill=BOTH, expand=TRUE)
        # 2.4.#
        ttk.Label(self.right_frame, font=("黑体", 10)).pack(fill=X)
        # 2.4.1 admin
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X, padx=(0, 20))
        self.admin = ttk.StringVar()
        self.current_admin = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['admin'][0]
        self.entry1 = ttk.Entry(self.right_frame, textvariable=self.admin, justify=LEFT, bootstyle=INFO)
        self.entry1.pack(fill=X, padx=(0, 20))
        self.admin.set(self.current_admin)
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X, padx=(0, 20))

        # 2.4.2 current_theme
        self.theme = ttk.Variable()
        self.current_theme = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['current_theme'][0]
        self.om1 = ttk.OptionMenu(self.right_frame, self.theme, *style.theme_names())
        self.theme.set(self.current_theme)
        self.om1.pack(side=TOP, fill=X, padx=(0, 20))
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X, padx=(0, 20))

        # 2.4.3 prj_name
        self.prj_name = ttk.Variable()
        self.current_prj = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['prj_name'][0]
        self.entry2 = ttk.Entry(self.right_frame, textvariable=self.prj_name, justify=LEFT, bootstyle=INFO)
        self.prj_name.set(self.current_prj)
        self.entry2.pack(side=TOP, fill=X, padx=(0, 20))
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X, padx=(0, 20))

        # 2.4.4 simplified_name
        self.save_name = ttk.Variable()
        self.current_save = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['simplified_name'][0]
        self.entry3 = ttk.Entry(self.right_frame, textvariable=self.save_name, justify=LEFT, bootstyle=INFO)
        self.save_name.set(self.current_save)
        self.entry3.pack(side=TOP, fill=X, padx=(0, 20))
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X, padx=(0, 20))
        # 2.4.#
        ttk.Label(self.right_frame).pack(fill=X)
        ttk.Label(self.right_frame, font=("黑体", 10)).pack(fill=X)
        ttk.Separator(self.left_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(fill=X)
        # 2.4.5 path1
        self.path1_frame = ttk.Frame(self.right_frame, height=0, bootstyle=SECONDARY)
        self.path1_frame.pack(side=TOP, fill=X, padx=(0, 20))
        self.path1 = ttk.StringVar()
        self.entryp1 = ttk.Entry(self.path1_frame,
                                 textvariable=self.path1,
                                 justify=LEFT,
                                 bootstyle=INFO,
                                 width=40
                                 )
        self.entryp1.pack(side=LEFT, fill=X, expand=TRUE)
        self.current_path1 = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['merge_file_path'][0]
        self.path1.set(self.current_path1)
        self.tk_btp1 = ttk.Button(self.path1_frame,
                                  text='> ',
                                  bootstyle=(LINK, LIGHT),
                                  command=self.change_path1,
                                  padding=(10, 0, 1.5, 0)
                                  )
        self.tk_btp1.pack(side=RIGHT, fill=BOTH, padx=(0, 1), pady=0)
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(side=TOP, fill=X, padx=(0, 20))

        # 2.4.6 path2
        self.path2_frame = ttk.Frame(self.right_frame, height=0, bootstyle=SECONDARY)
        self.path2_frame.pack(side=TOP, fill=X, padx=(0, 20))
        self.path2 = ttk.StringVar()
        self.entryp2 = ttk.Entry(self.path2_frame,
                                 textvariable=self.path2,
                                 justify=LEFT,
                                 bootstyle=INFO,
                                 width=40
                                 )
        self.entryp2.pack(side=LEFT, fill=X, expand=TRUE)
        self.current_path2 = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['data_base_path'][0]
        self.path2.set(self.current_path2)
        self.tk_btp2 = ttk.Button(self.path2_frame,
                                  text='> ',
                                  bootstyle=(LINK, LIGHT),
                                  command=self.change_path2,
                                  padding=(10, 0, 1.5, 0))
        self.tk_btp2.pack(side=RIGHT, fill=BOTH, padx=(0, 1), pady=0)
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(side=TOP, fill=X, padx=(0, 20))

        # 2.4.7 path3
        self.path3_frame = ttk.Frame(self.right_frame, height=0, bootstyle=SECONDARY)
        self.path3_frame.pack(side=TOP, fill=X, padx=(0, 20))
        self.path3 = ttk.StringVar()
        self.entryp3 = ttk.Entry(self.path3_frame,
                                 textvariable=self.path3,
                                 justify=LEFT,
                                 bootstyle=INFO,
                                 width=40
                                 )
        self.entryp3.pack(side=LEFT, fill=X, expand=TRUE)
        self.current_path3 = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['kaoqin_path'][0]
        self.path3.set(self.current_path3)
        self.tk_btp3 = ttk.Button(self.path3_frame,
                                  text='> ',
                                  bootstyle=(LINK, LIGHT),
                                  command=self.change_path3,
                                  padding=(10, 0, 1.5, 0))
        self.tk_btp3.pack(side=RIGHT, fill=BOTH, padx=(0, 1), pady=0)
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(side=TOP, fill=X, padx=(0, 20))
        # 2.4.8 path4
        self.path4_frame = ttk.Frame(self.right_frame, height=0, bootstyle=SECONDARY)
        self.path4_frame.pack(side=TOP, fill=X, padx=(0, 20))
        self.path4 = ttk.StringVar()
        self.entryp4 = ttk.Entry(self.path4_frame,
                                 textvariable=self.path4,
                                 justify=LEFT,
                                 bootstyle=INFO,
                                 width=40
                                 )
        self.entryp4.pack(side=LEFT, fill=X, expand=TRUE)
        self.current_path4 = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['renshu_path'][0]
        self.path4.set(self.current_path4)
        self.tk_btp4 = ttk.Button(self.path4_frame,
                                  text='> ',
                                  bootstyle=(LINK, LIGHT),
                                  command=self.change_path4,
                                  padding=(10, 0, 1.5, 0))
        self.tk_btp4.pack(side=RIGHT, fill=BOTH, padx=(0, 1), pady=0)
        ttk.Separator(self.right_frame, orient=HORIZONTAL, bootstyle=SECONDARY).pack(side=TOP, fill=X, padx=(0, 20))

        # 2.4.9 保存,恢复默认按钮
        ttk.Label(self.right_frame).pack(side=TOP, fill=X)
        bottom_frame = ttk.Frame(self.right_frame)
        bottom_frame.pack(side=TOP, fill=X, expand=False)
        ttk.Button(
            bottom_frame,
            text='保存',
            bootstyle=SUCCESS,
            width=10,
            command=lambda: self.save()
        ).pack(side=LEFT, fill=X, expand=False, padx=(20, 40), pady=(0, 0))

        ttk.Button(
            bottom_frame,
            text='恢复默认',
            bootstyle=INFO,
            width=10,
            command=lambda: self.reset()
        ).pack(side=LEFT, fill=X, expand=False, padx=(20, 40), pady=(0, 0))

    def fanhui(self):
        '''返回'''
        self.back.destroy()
        main.WelcomeFrame(self.root)

    def change_path1(self):
        '''选择文件保存位置'''
        path1 = filedialog.askdirectory()
        if path1 != '':
            self.path1.set(path1)

    def change_path2(self):
        '''选择文件保存位置'''
        path2 = filedialog.askdirectory()
        if path2 != '':
            self.path2.set(path2)

    def change_path3(self):
        '''选择文件保存位置'''
        path3 = filedialog.askdirectory()
        if path3 != '':
            self.path3.set(path3)

    def change_path4(self):
        '''选择文件保存位置'''
        path4 = filedialog.askdirectory()
        if path4 != '':
            self.path4.set(path4)

    def save(self):
        '''保存'''
        setting_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))
        setting_df.loc[0, "admin"] = self.entry1.get()
        setting_df.loc[0, "current_theme"] = self.theme.get()
        setting_df.loc[0, "prj_name"] = self.entry2.get()
        setting_df.loc[0, "simplified_name"] = self.entry3.get()
        setting_df.loc[0, "merge_file_path"] = self.entryp1.get()
        setting_df.loc[0, "data_base_path"] = self.entryp2.get()
        setting_df.loc[0, "kaoqin_path"] = self.entryp3.get()
        setting_df.loc[0, "renshu_path"] = self.entryp4.get()
        for i in range(8):
            if setting_df.iloc[0, i] == '':
                setting_df.iloc[0, i] = ' '
        setting_df.to_csv(Path.joinpath(SAVE_PATH, 'setting.csv'), index=False)
        style = ttk.Style()
        style.theme_use(self.theme.get())
        Messagebox.show_info_info('保存成功!', '温馨提示!')

    def reset(self):
        '''恢复默认'''
        configure_path1 = str(Mer_PATH).replace('\\', '/')
        configure_path2 = str(DATA_PATH).replace('\\', '/')
        configure_path3 = str(KQ_PATH).replace('\\', '/')
        configure_path4 = str(RS_PATH).replace('\\', '/')

        setting_list = [
            'Admin', 'litera', ' ', ' ',
            configure_path1, configure_path2, configure_path3, configure_path4
        ]
        head_list = ['admin', 'current_theme', 'prj_name', 'simplified_name', 'merge_file_path',
                     'data_base_path', 'kaoqin_path', 'renshu_path']
        with open(Path.joinpath(SAVE_PATH, 'setting.csv'), 'w', encoding='utf-8', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(head_list)
            writer.writerow(setting_list)
            csvfile.close()

        style = ttk.Style()
        style.theme_use('litera')
        self.admin.set('Admin')
        self.theme.set('litera'),
        self.prj_name.set(' ')
        self.save_name.set(' ')
        self.path1.set(configure_path1),
        self.path2.set(configure_path2),
        self.path3.set(configure_path3),
        self.path4.set(configure_path4),
        Messagebox.show_info_info('已恢复默认设置!', '温馨提示!')
