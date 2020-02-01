# -*- coding: utf-8 -*-

import xlwings as xw
import tushare as ts


@xw.sub
def get_industry_classified():
    wb = xw.Book.caller()
    sht = wb.sheets[0]
    df = ts.get_industry_classified()

    # wb = xw.Book('') # wb = xw.Book(filename) would open an existing file
    wb=xw.Book.caller()

    # 删除特定名称的Sheet表
    # wb.sheets("行业数据").delete()

    try:
        # 删除特定名称的Sheet表
        wb.sheets("行业数据").delete()
    except:
        print("Sheet does NOT exist!!!")
        # 如果某个Sheet已经存在，则删除！ 对于不存在的表单删除，则会报错！！！


    # 新建一个表单，并且在新的表单中进行操作
    ws = wb.sheets.add("行业数据", after="数据工作台")
    # 选择已经创建的表单
    # ws = wb.sheets["Sheet1"]

    ws.range("A1").value = df

    # 重新回到主控制台的工作簿
    wb.sheets("数据工作台").activate()

