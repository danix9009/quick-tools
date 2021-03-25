# -*- coding:utf-8 -*
import pymysql
import xlrd,xlutils
from xlrd import xldate_as_tuple
import re
import json

class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
        #定义一个属性接收文件路径
        self.data_path = data_path
        # 定义一个属性接收工作表名称
        self.sheetname = sheetname
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_name(self.sheetname)
        # 根据工作表的索引获取工作表的内容（方式②）
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(0)
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        # 获取工作表的有效列数
        self.colNum = self.table.ncols

    def readExcel(self):
        # 定义一个空列表
        datas = []
        for i in range(1, self.rowNum):
            # 定义一个空字典
            sheet_data = {}

            for j in range(self.colNum):
                c_type = self.table.cell(i,j).ctype
                # 获取单元格数据
                c_cell = self.table.cell_value(i, j)
                if c_type == 2 and c_cell % 1 == 0:  # 如果是整形
                    c_cell = int(c_cell)
                sheet_data[self.keys[j]] = c_cell
            
            if sheet_data:
                datas.append(sheet_data)
        return datas

if __name__ == "__main__":
    data_path = "xxxx.xlsx"
    sheetname = "sheet1"
    m = {}
    m[1] = "TEST012"
    m[2] = 1256
    m[3] = "2020-08-17T17:16:36.517+08:00"
    m[4] = ["a","b","c"]
    m[5] = [1, 2, 3]
    m[9] = [116.397128,39.916527]
    get_data = ExcelData(data_path, sheetname)
    datas = get_data.readExcel()
    # 添加所需要的其他字段
    print("{\"categoryId\":3632,\n\"data\":{")
    for d in datas:
        t = d['fieldType']
        if t ==1 or t==3 :
            print("\"{}\":\"{}\", //{}".format(d['field'],m[t],d['fieldName']))
        else:
            print("\"{}\":{}, //{}".format(d['field'],m[t],d['fieldName']))
    # 添加提所需要的其他字段
    print("},\n\"factId\":\"test123\"}")
