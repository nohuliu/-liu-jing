import pandas as pd
import os
import shutil
from openpyxl import load_workbook
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation


def mkdir(path):
    path = path.strip()     # 去除首位空格
    path = path.rstrip("\\")     # 去除尾部 \ 符号
    isExists = os.path.exists(path)     # 判断路径是否存在
    if not isExists:     # 如果不存在则创建目录,创建目录操作函数
        # os.mkdir(path)与os.makedirs(path)的区别是,当父目录不存在的时候os.mkdir(path)不会创建，os.makedirs(path)则会创建父目录
        # 此处路径最好使用utf-8解码，否则在磁盘中可能会出现乱码的情况
        path_bytes = path.encode("utf-8")
        path_utf8 = path_bytes.decode("utf-8")
        os.makedirs(path_utf8)
        print(path + "创建成功")
        return True
    else:      # 如果目录存在则不创建，并提示目录已存在
        print(path + "目录已存在")
        return False


banzu = ["甲班", "乙班", "丙班"]
for j in banzu:
    path = r"C:\\Users\\电工\\Desktop\\巡检记录\\2023年" + j + "巡检记录"
    mkdir(path)
    for i in range(1, 13):
        path = r"C:\\Users\\电工\\Desktop\\巡检记录\\2023年" + \
            j + "巡检记录\\" + str(i) + "月"
        mkdir(path)
    data = pd.read_excel("C:/Users/电工/Desktop/巡检记录/文件名 - 副本.xls")
    模板 = "C:\\Users\\电工\\Desktop\\巡检记录\\模板.xlsx"
    for i in range(0, len(data.index.values)):
        if j == "甲班":
            x = 0
        if j == "乙班":
            x = 1
        if j == "丙班":
            x = 2
        if data.iloc[i][x] == j:
            fuzhipath = "C:\\Users\\电工\\Desktop\\巡检记录\\2023年" + j + "巡检记录\\" + \
                str(data.iloc[i][6]) + "月\\" + \
                str(data.iloc[i][3]) + \
                str(data.iloc[i][4]) + str(data.iloc[i][5])
            shutil.copyfile(模板, fuzhipath)
    for i in range(0, len(data.index.values)):
        if data.iloc[i][x] == j:
            if j == "甲班":
                xunjianren = '"李国卿、白金钟,李国卿、刘晶,白金钟、刘晶"'
            if j == "乙班":
                xunjianren = '"赵贯策、张达,赵贯策、刘业,张达、刘业"'
            if j == "丙班":
                xunjianren = '"边永利、甄一昶,边永利、李建绰,甄一昶、李建绰"'
            dakaipath = "C:\\Users\\电工\\Desktop\\巡检记录\\2023年" + j + "巡检记录\\" + \
                str(data.iloc[i][6]) + "月\\" + \
                str(data.iloc[i][3]) + \
                str(data.iloc[i][4]) + str(data.iloc[i][5])
            wb = load_workbook(dakaipath)
            s1 = wb.get_sheet_by_name("路线1变频器巡检记录一")
            s1["O2"].value = str(data.iloc[i][4])
            s1["G2"].value = j
            dv1 = DataValidation(
                type="list", formula1=xunjianren, allow_blank=True)
            temp1 = s1["K2"]
            dv1.add(temp1)
            s1.add_data_validation(dv1)
            s2 = wb.get_sheet_by_name("路线1高低压室巡检记录二")
            s2["R2"].value = str(data.iloc[i][4])
            s2["I2"].value = j
            dv2 = DataValidation(
                type="list", formula1=xunjianren, allow_blank=True)
            temp2 = s2["N2"]
            dv2.add(temp2)
            s2.add_data_validation(dv2)
            s3 = wb.get_sheet_by_name("路线1水阻柜巡检记录三")
            s3["H2"].value = str(data.iloc[i][4])
            s3["C2"].value = j
            dv3 = DataValidation(
                type="list", formula1=xunjianren, allow_blank=True)
            temp3 = s3["E2"]
            dv3.add(temp3)
            s3.add_data_validation(dv3)
            s4 = wb.get_sheet_by_name("路线2变频器巡检记录一")
            s4["O2"].value = str(data.iloc[i][4])
            s4["G2"].value = j
            dv4 = DataValidation(
                type="list", formula1=xunjianren, allow_blank=True)
            temp4 = s4["K2"]
            dv4.add(temp4)
            s4.add_data_validation(dv4)
            wb.save(dakaipath)
    print("2023年" + j + "巡检记录已完成")
