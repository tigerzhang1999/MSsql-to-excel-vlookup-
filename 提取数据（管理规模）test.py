import pyodbc
import xlrd
from xlutils.copy import copy
from xlrd import *

# 创建元组
regnum = ()
# 打开excel文件，创建一个workbook对象
rbook = xlrd.open_workbook("D:/PEdata/Hedgefundtest.xls")
# sheets方法返回对象列表
rbook.sheets()
# 取第一个工作簿
rsheet = rbook.sheet_by_index(0)

# 循环工作簿的所有行
for row in rsheet.get_rows():  # 循环创建所有备案号的元组
    product_column = row[1]  # 备案号所在列
    product_value = product_column.value  # 提取赋值备案号
    if product_value != '产品备案编号':
        regnumtemp = (product_value,)
        regnum = regnum + regnumtemp

print(regnum) # 打印备案号元组确认是否正确

#  连接服务器
cnxn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};"
                            "SERVER="
                            "Database="
                            "PWD=")

cursor = cnxn.cursor()  # 创建cursor查询

for index in range(len(regnum)):   # 进行对于相匹配的备案号的私募名称循环查询
    inx = regnum[index]  # 提取当次循环的备案号
    # 定义sql查询语句，？为待补充参数
    sql = """
    SELECT issuing_scale FROM t_fund_info where reg_code = ?
    """
    cursor.execute(sql,inx)  # inx为补充参数
    data = cursor.fetchall()  # 提取数据
    data2 = "".join('%s' %a for a in data)
    data3 = data2.strip('()[]')  # 去除首尾无用字符
    data4 = data3.strip('Decimal')
    print(data4)  #打印查看是否正确
    w = copy(open_workbook('D:/PEdata/hedgefund.xls'))  # 复制一份要写入的文档
    w.get_sheet(0).write(index+1, 8, str(data4))  # 定义写入行，列，数据
    w.save('D:/PEdata/hedgefund.xls')  #保存


