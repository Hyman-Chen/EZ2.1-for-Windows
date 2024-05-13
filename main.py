import os
import sys
from pathlib import Path
from itertools import cycle
from PIL import Image, ImageTk, ImageSequence

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Bootstyle
import pandas as pd

import setting_page
import page_1
import page_2
import page_3
import page_4
import page_5
import page_6
import clock

# 初始化源程序所在文件夹
BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
# 初始化图片文件夹
IMG_PATH = Path(BASE_PATH) / 'data' / 'image'
# 初始化应用数据文件夹
SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'


# 定义App
class App(ttk.Window):
    '''主页'''

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # 初始化标题栏
        self.title('EZ')
        self.iconbitmap(Path.joinpath(IMG_PATH, 'ico_64px.ico'))

        # 初始化主题
        setting_df = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))
        theme_name = setting_df['current_theme'][0]
        style = ttk.Style()
        style.theme_use(theme_name)

        # 初始化窗口
        self.geometry('700x470')
        self.minsize(700, 470)
        # 1 左边栏
        self.cf = CollapsingFrame(self)
        self.cf.pack(fill=BOTH, side=LEFT)
        # 2 设置栏
        self.sf = SettingFrame(self)
        self.sf.pack(side=TOP, fill=BOTH)
        # 3 显示栏
        self.showframe = ttk.Frame(self)
        self.showframe.pack(side=TOP, fill=BOTH)
        WelcomeFrame(self)
        self.bind('<Escape>', lambda x: self.fullscreen(x))
        self.mainloop()

    def fullscreen(self, event=None):
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))


class CollapsingFrame(ttk.Frame):
    """1 左边栏"""

    def __init__(self, master):
        super().__init__(master)

        # 图片
        self.images = [
            ttk.PhotoImage(file=IMG_PATH / 'info_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'icons8_double_up_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'icons8_double_right_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'info_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'info_24px_1.png')
        ]

        # 初始化颜色
        self.configure(bootstyle='dark')
        self.root = master
        ttk.Button(self, text="主页", bootstyle=DARK, command=self.show).grid(row=0, column=0, stick="nsew")
        style_color = DARK
        Bootstyle.ttkstyle_widget_color(DARK)
        self.cumulative_rows = 0

        # 1.1 数据中心
        self.group1 = ttk.Frame(self, bootstyle=style_color, padding=10)
        # 1.1.1 名单汇总
        self.bt1 = ttk.Button(
            self.group1,
            image=self.images[0],
            compound=LEFT,
            text='  名单汇总',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_1.MingdanhuizongFrame(self.root)
        )
        self.bt1.pack(fill=X)
        # 1.1.2 现场人数
        self.bt2 = ttk.Button(
            self.group1,
            image=self.images[0],
            compound=LEFT,
            text='  现场人数',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_2.XianchangqingkuangFrame(self.root)

        )
        self.bt2.pack(fill=X)
        # 1.1.3 实时人数
        self.bt3 = ttk.Button(
            self.group1,
            image=self.images[0],
            compound=LEFT,
            text='  实时人数',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_3.Shishirenshu(self.root)

        )
        self.bt3.pack(fill=X)
        # 1.1.4 数据库
        self.bt4 = ttk.Button(
            self.group1,
            image=self.images[0],
            compound=LEFT,
            text='  数据库',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_4.ShujukuFrame(self.root)
        )
        self.bt4.pack(fill=X)
        # 1.1.5 手动导入
        self.bt5 = ttk.Button(
            self.group1,
            image=self.images[0],
            compound=LEFT,
            text='  手动导入',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_5.ZongbiaozhizuoFrame(self.root)

        )
        self.bt5.pack(fill=X)

        self.add(child=self.group1, title='  数据管理', bootstyle=DARK)
        # 1.2 参建单位
        self.group2 = ttk.Frame(self, bootstyle=style_color, padding=10)
        # 1.2.1 单位管理
        self.bt6 = ttk.Button(
            self.group2,
            image=self.images[0],
            compound=LEFT,
            text='  单位管理',
            bootstyle=(DEFAULT, DARK),
            command=lambda: page_6.AppendcoFrame(self.root)

        )
        self.bt6.pack(fill=X)
        self.add(self.group2, title='  参建单位', bootstyle=DARK)

    def show(self):
        WelcomeFrame(self.root)

    def add(self, child, title="", bootstyle=DARK, **kwargs):
        if child.winfo_class() != 'TFrame':
            return

        style_color = bootstyle
        Bootstyle.ttkstyle_widget_color(bootstyle)
        frm = ttk.Frame(self, bootstyle=style_color)
        frm.grid(row=self.cumulative_rows + 1, column=0, sticky=EW)

        # header title
        header = ttk.Label(
            master=frm,
            image=self.images[4],
            text=title,
            bootstyle=(style_color, INVERSE),
            compound=LEFT,
        )
        if kwargs.get('textvariable'):
            header.configure(textvariable=kwargs.get('textvariable'))
        header.pack(side=LEFT, fill=BOTH, padx=10)

        # header toggle button
        def _func(c=child):
            return self._toggle_open_close(c)

        btn = ttk.Button(
            master=frm,
            image=self.images[1],
            bootstyle=style_color,
            command=_func
        )
        btn.pack(side=RIGHT)

        # assign toggle button to child so that it can be toggled
        child.btn = btn
        child.grid(row=self.cumulative_rows + 2, column=0, sticky=NSEW)

        child.grid_remove()
        child.btn.configure(image=self.images[2])
        # increment the row assignment
        self.cumulative_rows += 2

    def _toggle_open_close(self, child):
        """Open or close the section and change the toggle button
        image accordingly.

        Parameters:

            child (Frame):
                The child element to add or remove from grid manager.
        """
        if child.winfo_viewable():
            child.grid_remove()
            child.btn.configure(image=self.images[2])
        else:
            child.grid()
            child.btn.configure(image=self.images[1])


class SettingFrame(ttk.Frame):
    '''2 设置栏'''

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.root = master
        # 图片
        self.images = [
            ttk.PhotoImage(file=IMG_PATH / 'icons8_settings_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'icons8_folder_24px.png'),
            ttk.PhotoImage(file=IMG_PATH / 'icons8_add_folder_24px.png'),
        ]
        # 设置按钮
        buttonbar = ttk.Frame(self, bootstyle=PRIMARY)
        buttonbar.pack(fill=BOTH, side=TOP)
        self.setting_bt = ttk.Button(
            master=buttonbar,
            text='设置',
            image=self.images[0],
            compound=LEFT,
            bootstyle=('default', PRIMARY),
            width=5,
            command=lambda: setting_page.SettingFrame(self.root)
        )
        self.setting_bt.pack(side=RIGHT)


class WelcomeFrame(ttk.Frame):
    '''3 显示栏'''

    def __init__(self, master):
        super().__init__(master, width=10, height=10)
        self.root = master
        self.root.showframe.destroy()
        self.root.showframe = self
        self.pack(side=TOP, fill=BOTH, expand=True)
        self.rowconfigure(index=0, weigh=1)
        self.columnconfigure(index=0, weigh=1)
        myname = pd.read_csv(Path.joinpath(SAVE_PATH, 'setting.csv'))['admin'][0]
        if myname == 'admin':
            myname = ''
        else:
            myname = str(myname) + ','
        framet = ttk.Frame(self)
        framet.pack(side=TOP, fill=X, expand=False)
        framet_left = ttk.Frame(framet)
        framet_left.pack(side=LEFT, fill=X, expand=False)
        framet_right = ttk.Frame(framet)
        framet_right.pack(side=RIGHT, fill=X, expand=False)

        ttk.Label(framet, anchor='w', text=myname + "欢迎使用!", font=("黑体", 24), padding=(0, 20)).pack(side=LEFT, fill=X,
                                                                                                    expand=False)

        clock.ClockFrame(framet)

        # 设置背景标签
        file_path = Path.joinpath(IMG_PATH, 'main_page_gif.gif')
        with Image.open(file_path) as im:
            # 创建图片列表
            sequence = ImageSequence.Iterator(im)
            images = [ImageTk.PhotoImage(s) for s in sequence]
            self.image_cycle = cycle(images)
            # 时间间隔
            self.framerate = im.info["duration"]

        self.img_container = ttk.Label(self, anchor='center', image=next(self.image_cycle))
        self.img_container.pack(side=RIGHT, fill=BOTH, expand=True)
        self.after(self.framerate, self.next_frame)

    def next_frame(self):
        '''下一张'''
        self.img_container.configure(image=next(self.image_cycle))
        self.after(self.framerate, self.next_frame)


if __name__ == '__main__':
    App()
