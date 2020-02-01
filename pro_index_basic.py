# -*- coding: utf-8 -*-

import xlwings as xw
import tushare as ts


@xw.sub
def pro_index_basic():
    wb = xw.Book.caller()
    sht = wb.sheets[0]

    pro = ts.pro_api('a9eb78ead4fb3fb07fb17e9dd6fc7615d734eac70ebfe7634bee30d8')

    df = pro.index_basic(market='SW')

    # wb = xw.Book('') # wb = xw.Book(filename) would open an existing file
    wb=xw.Book.caller()

    try:
        # 删除特定名称的Sheet表
        wb.sheets("指数列表").delete()
    except:
        print("Sheet does NOT exist!!!")
        # 如果某个Sheet已经存在，则删除！ 对于不存在的表单删除，则会报错！！！


    # 新建一个表单，并且在新的表单中进行操作
    ws = wb.sheets.add("指数列表", after="数据工作台")
    # 选择已经创建的表单
    # ws = wb.sheets["Sheet1"]

    ws.range("A1").value = df

    # 重新回到主控制台的工作簿
    wb.sheets("数据工作台").activate()
