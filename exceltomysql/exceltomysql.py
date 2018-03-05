"""
Function: 将Excel 数据导入到Mysql数据库中

"""
import xlrd
import pymysql

def insert(sheet,db):
    sheeti = book.sheet_by_name(sheet)
    cursor = db.cursor()
    query = """INSERT INTO """+sheet+""" (sequence, testprocedure, passcriteria, sku1, sku2, 
    sku3, sku4, notes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
    for i in range(1, sheeti.nrows):
        if i > 7:
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
    print("Done!")
    print("")
    columns = str(sheeti.ncols)
    rows = str(sheeti.nrows)
    print("我刚导入了 %s 列 和 %s 行数据到数据库." % (columns, rows))


db = pymysql.connect("localhost", "root", "tdemlozpdeq", "test", use_unicode=True, charset="utf8")
book = xlrd.open_workbook("C:/Users/95/Desktop/testplan/testplan.xlsx")


sheet_names = book.sheet_names()
name_list = []
for j in sheet_names:
    insert(j,db)


# def insert(sheet):
#     sheeti = book.sheet_by_name(sheet)
#     cursor = db.cursor()
#     query = """INSERT INTO """+sheet+""" (sequence, testprocedure, passcriteria, sku1, sku2,
#     sku3, sku4, notes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
#     for i in range(1, sheet.nrows):
#         if i > 7:
#             sequence = sheeti.cell(i, 0).value
#             testprocedure = sheeti.cell(i, 1).value
#             passcriteria = sheeti.cell(i, 2).value
#             sku1 = sheeti.cell(i, 3).value
#             sku2 = sheeti.cell(i, 4).value
#             sku3 = sheeti.cell(i, 5).value
#             sku4 = sheeti.cell(i, 6).value
#             notes = sheeti.cell(i, 7).value
#
#             values = (sequence, testprocedure, passcriteria, sku1, sku2, sku3, sku4, notes)
#             cursor.execute(query, values)
#     cursor.close()
#     db.commit()
#     print("")
#     print("Done!")
#     print("")
#     columns = str(sheeti.ncols)
#     rows = str(sheeti.nrows)
#     print("我刚导入了 %s 列 和 %s 行数据到数据库." % (columns, rows))




# cursor = db.cursor()
# query = """INSERT INTO mst01 (sequence, testprocedure, passcriteria, sku1, sku2, sku3, sku4, notes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
#
# for i in range(1,sheet.nrows):
#     if i > 7:
#         sequence = sheet.cell(i,0).value
#         testprocedure = sheet.cell(i,1).value
#         passcriteria = sheet.cell(i,2).value
#         sku1 = sheet.cell(i,3).value
#         sku2 = sheet.cell(i,4).value
#         sku3 = sheet.cell(i,5).value
#         sku4 = sheet.cell(i,6).value
#         notes = sheet.cell(i,7).value
#
#         values = (sequence, testprocedure, passcriteria, sku1, sku2, sku3, sku4, notes)
#         cursor.execute(query, values)
#
#
# cursor.close()
# db.commit()
# print("")
# print("Done!")
# print("")
# columns = str(sheet.ncols)
# rows = str(sheet.nrows)
# print("我刚导入了 %s 列 和 %s 行数据到数据库." % (columns, rows))
