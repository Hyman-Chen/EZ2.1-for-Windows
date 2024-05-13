from pathlib import Path
import os
import sys
import pandas as pd

BASE_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))
SAVE_PATH = Path(BASE_PATH) / 'data' / 'save'


def judge(row_val, date, unit, name, GZLB, NLLB, SF):
    '''在总表中查询人员是否存在(日期,单位,姓名,工种类别,年龄类别,身份)'''
    date1 = date.replace('-', '.')
    date2 = date.replace('-', '/')
    try:
        date3 = ('{}.{}.{}').format(date[:4], int(date[5:7]), int(date[8:10]))
        date4 = ('{}/{}/{}').format(date[:4], int(date[5:7]), int(date[8:10]))
    except:
        date3 = ""
        date4 = ""
    date5 = date1 + "00:00:00"
    date6 = date1 + " 00:00:00"
    date7 = date2 + "00:00:00"
    date8 = date2 + " 00:00:00"
    teshugongzhong = [
        '电工',
        '电焊工',
        '焊工',
        '电焊',
        '架子工',
        '吊车司机',
        '塔吊司机',
        '汽吊司机',
        '汽车吊司机',
        '汽吊',
        '人货梯司机',
        '起重司机',
        '汽吊指挥',
        '塔吊指挥',
        '气割',
        '吊篮',
        '指挥',
        '架子',
        '信号工',
    ]
    if row_val[3] in teshugongzhong:
        row_val[3] = "特殊工种"
    else:
        row_val[3] = "普通工种"
    if row_val[4] == "女":
        if int(row_val[5]) >= 45:
            row_val[4] = "超龄"
        else:
            row_val[4] = "未超龄"
    else:
        if int(row_val[5]) >= 55:
            row_val[4] = "超龄"
        else:
            row_val[4] = "未超龄"
    if row_val[3] in ["管理", "管理人员"]:
        row_val[6] = "管理人员"
    else:
        row_val[6] = "作业人员"

    date_list = [row_val[9], ""]
    if unit != "":
        unit_df = pd.read_csv(Path.joinpath(SAVE_PATH, "co_dict.csv"))
        unit_list = list(unit_df.loc[:, unit])
    else:
        row_val[1] = ""
        unit_list = [""]
    name_list = [row_val[2], ""]
    GZLB_list = [row_val[3], ""]
    NLLB_list = [row_val[4], ""]
    SF_list = [row_val[6], ""]
    if (
            (row_val[1] in unit_list) and
            (name in name_list) and
            (GZLB in GZLB_list) and
            (NLLB in NLLB_list) and
            (SF in SF_list)
    ):
        for i in [date1, date2, date3, date4, date5, date6, date7, date8]:
            if i in date_list:
                return True
            else:
                return False
    else:
        return False
