from pathlib import Path
import datetime
import pandas as pd
import os
import sys
import shutil
from tkinter import filedialog

from ttkbootstrap.dialogs import Messagebox
from openpyxl import load_workbook
from xlrd import open_workbook
import initTool

BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
IMG_PATH = Path(BASE_PATH) / 'data' / 'image'

SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'
tempDATA_PATH = Path(BASE_PATH) / 'data' / 'temp'
XlFilesToBeMerged_PATH = Path(tempDATA_PATH) / 'XlFiles'
csvFilesToBeMerged_PATH = Path(tempDATA_PATH) / 'csvFiles'
renshuXlData_PATH = Path(tempDATA_PATH) / 'RS'
renshucsvData_PATH = Path(tempDATA_PATH) / 'RScsv'
kaoqinXlData_PATH = Path(tempDATA_PATH) / 'KQ'
kaoqincsvData_PATH = Path(tempDATA_PATH) / 'KQcsv'
appendXlData_PATH = Path(tempDATA_PATH) / 'AD'
appendcsvData_PATH = Path(tempDATA_PATH) / 'ADcsv'

DATA_PATH = Path(BASE_PATH)
Mer_PATH = Path(BASE_PATH)


def select_files(*master):
    '''先删除原有的excel文件，然后把选中的文件保存下来'''
    for entry in Path(XlFilesToBeMerged_PATH).iterdir():
        if entry.is_file() or entry.is_symlink():
            entry.unlink()
    file_paths = filedialog.askopenfilenames()
    if file_paths != '':
        for file in file_paths:
            shutil.copy2(file, XlFilesToBeMerged_PATH)
        for _ in map(master[0].treeview1.delete, master[0].treeview1.get_children("")):
            pass
        filelist = []
        num = 1
        for entry in Path(XlFilesToBeMerged_PATH).iterdir():
            if str(entry).endswith('.xlsx') or str(entry).endswith('.xls'):
                listtemp = []
                listtemp.append(num)
                listtemp.append(entry.stem)
                filelist.append(listtemp)
                num += 1
        master[0].tableValues1 = filelist
        row = 1
        for data in master[0].tableValues1:
            master[0].treeview1.insert('', 'end', text=row, value=data, iid=str(row))
            row += 1
        Messagebox.show_info_success('待合并文件添加成功！', '温馨提示')


def merge_files(*master):
    '''合并文件'''
    fileslist = []
    num = 1
    # 创建文件名列表
    for entry in Path(XlFilesToBeMerged_PATH).iterdir():
        if str(entry).endswith('.xlsx') or str(entry).endswith('.xls'):
            listtemp = []
            listtemp.append(num)
            listtemp.append(entry.stem)
            fileslist.append(listtemp)
            num += 1
    # 若未添加文件，则显示提示
    if fileslist == []:
        Messagebox.show_info_success('未检测到待合并文件！', '温馨提示')
    # 否则开始进行合并
    else:
        # 清空之前的csv数据
        for entry in Path(tempDATA_PATH).iterdir():
            if entry.is_file() or entry.is_symlink():
                entry.unlink()
        # 把XlFilesToBeMerged里面的excel文件转换成csv文件，以file+n为文件名
        for i in fileslist:
            try:
                myworkbook = load_workbook(Path.joinpath(XlFilesToBeMerged_PATH, str(i[1]) + '.xlsx'), read_only=True)
                myworksheet = myworkbook.worksheets[0]
                initTool.Files().file_to_csv2(myworksheet, csvFilesToBeMerged_PATH, 'file' + str(i[0]) + '.csv')
                myworkbook.close()
            except:
                myworkbook = open_workbook(Path.joinpath(XlFilesToBeMerged_PATH, str(i[1]) + '.xls'))
                myworksheet = myworkbook.sheets()[0]
                initTool.Files().file_to_csv2(myworksheet, csvFilesToBeMerged_PATH, 'file' + str(i[0]) + '.csv')
                myworkbook.release_resources()
                del myworkbook
        # 合并生成的csv文件，保存在total.csv中
        initTool.Files().merge_csv_files()
        # 使序号列按顺序排序
        initTool.Files().fresh_csv(csvFilesToBeMerged_PATH, 'total.csv')
        today = datetime.date.today()
        month = today.month
        day = today.day
        filename = pd.read_csv(Path.joinpath(SAVE_PATH, "setting.csv")).loc[0, "simplified_name"]
        if filename == ' ':
            filename = ''
        # 定义合并文件名并保存文件
        filename = str(month) + "月" + str(day) + "日" + filename + "培训名单" + ".xlsx"
        filepath = master[0].Mer_PATH
        initTool.Files().csv_to_excel(today, filepath, filename)
        master[0].fresh()
        '''把AD文件夹内的文件删除，把选择的文件复制到AD文件夹'''
        for entry in Path(appendXlData_PATH).iterdir():
            if entry.is_file() or entry.is_symlink():
                entry.unlink()
        shutil.copy2(Path.joinpath(filepath, filename), appendXlData_PATH)

        Messagebox.show_info_success('保存成功！', '温馨提示')


def importfile():
    '''把AD文件夹内的文件删除，把选择的文件复制到AD文件夹'''
    for entry in Path(appendXlData_PATH).iterdir():
        if entry.is_file() or entry.is_symlink():
            entry.unlink()
    file_paths = filedialog.askopenfilenames()
    if file_paths != '':
        for file in file_paths:
            shutil.copy2(file, appendXlData_PATH)
        Messagebox.show_info_success('文件添加成功！', '温馨提示')


def select_file_rs(*master):
    '''删除RS文件夹中的所有文件(.xls文件)，然后把选中文件复制到RS文件夹中'''
    for entry in Path(renshuXlData_PATH).iterdir():
        if entry.is_file() or entry.is_symlink():
            entry.unlink()
    file_paths = filedialog.askopenfilename()
    if file_paths != '':
        for file in file_paths:
            shutil.copy2(file, renshuXlData_PATH)
        # 刷新页面
        for _ in map(master[0].treeview1.delete, master[0].treeview1.get_children("")):
            pass
        filelist = []
        num = 1
        for entry in Path(renshuXlData_PATH).iterdir():
            if str(entry).endswith('.xlsx') or str(entry).endswith('.xls'):
                listtemp = []
                listtemp.append(num)
                listtemp.append(entry.stem)
                filelist.append(listtemp)
                num += 1
        master[0].tableValues1 = filelist
        row = 1
        for data in master[0].tableValues1:
            master[0].treeview1.insert('', 'end', text=row, value=data, iid=str(row))
            row += 1
        Messagebox.show_info_success('添加成功！', '温馨提示')


def select_file_kq(*master):
    '''先删除KQ文件夹里的Xl文件,然后把选择的考勤表复制到KQ文件夹中'''
    for entry in Path(kaoqinXlData_PATH).iterdir():
        if entry.is_file() or entry.is_symlink():
            entry.unlink()
    file_path = filedialog.askopenfilename()
    if file_path != '':
        # for file in file_path:
        shutil.copy2(file_path, kaoqinXlData_PATH)
        # 刷新页面
        for _ in map(master[0].treeview1.delete, master[0].treeview1.get_children("")):
            pass
        filelist = []
        num = 1
        for entry in Path(kaoqinXlData_PATH).iterdir():
            if str(entry).endswith('.xlsx'):
                listtemp = []
                listtemp.append(num)
                listtemp.append(entry.stem)
                filelist.append(listtemp)
                num += 1
        master[0].tableValues1 = filelist
        row = 1
        for data in master[0].tableValues1:
            master[0].treeview1.insert('', 'end', text=row, value=data, iid=str(row))
            row += 1
        Messagebox.show_info_success('考勤表添加成功！', '温馨提示')
