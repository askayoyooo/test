"""
Function: 将Excel 数据导入到Mysql数据库中
"""
import xlrd
import pymysql


def insert(sheet, db, database):
    print("正在插入" + sheet)
    sheeti = book.sheet_by_name(sheet)
    cursor = db.cursor()
    query = """INSERT INTO """+database+""" (sequence, testprocedure, passcriteria, sku1, sku2, 
    sku3, sku4, notes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
    for i in range(0, sheeti.nrows):
        if i >= 0:
            sequence = sheeti.cell(i, 0).value
            testprocedure = sheeti.cell(i, 1).value
            passcriteria = sheeti.cell(i, 2).value
            sku1 = sheeti.cell(i, 3).value
            sku2 = sheeti.cell(i, 4).value
            sku3 = sheeti.cell(i, 5).value
            sku4 = sheeti.cell(i, 6).value
            notes = sheeti.cell(i, 7).value
            values = (sequence, testprocedure, passcriteria, sku1, sku2, sku3, sku4, notes)
            cursor.execute(query, values)
    cursor.close()
    db.commit()
    print("")
    print("Done!"+sheet)
    print("")
    columns = str(sheeti.ncols)
    rows = str(sheeti.nrows)
    print("我刚导入了 %s 列 和 %s 行数据到数据库." % (columns, rows))


# db = pymysql.connect("localhost", "root", "tdemlozpdeq", "test", use_unicode=True, charset="utf8")  # 家中开发机
db = pymysql.connect("localhost", "root", "123456", "test", use_unicode=True, charset="utf8")  # 办公开发机
# book = xlrd.open_workbook("C:/Users/95/Desktop/testplan/testplan.xlsx")  # 家中开发机
book = xlrd.open_workbook("C:/Users/Eiffel238i sku2-3/Documents/testplan451.xlsx")  # 办公开发机

sheet_names = book.sheet_names()


for i in range(len(sheet_names)):
    database_name = "sheet_"+str(10000+i)+"_"+sheet_names[i].lower()
    sheet_name = sheet_names[i]
    insert(sheet_name, db, database_name)



