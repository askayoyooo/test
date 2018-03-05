# __author: "Tim Zhong"
# date: 2018/3/4


import xlrd

book = xlrd.open_workbook("C:/Users/Eiffel238i sku2-3/Documents/testplan451.xlsx")  # 办公开发机
name_list = book.sheet_names()

# name_list = ['MST01', 'DRV01', 'Aud01a', 'Aud01b', 'Aud01c ',
# 'Aud01d', 'Aud01e', 'Aud01f', 'Aud02', 'Vdo01a', 'Vdo01b', 'Vdo01c', 'Vdo01d',
# 'Vdo01e', 'Vdo01f', 'Vdo01g', 'Vdo01h', 'Vdo01i', 'Vdo01j', 'Vdo01k', 'Vdo02',
# 'Vdo03', 'Vdo04', 'Vdo05', 'Vdo06', 'Vdo07', 'Vdo08', 'Vdo09', 'Vdo10', 'Vdo11',
# 'Vdo12', 'SW01', 'SW02', 'SW03', 'SW05', 'SW07', 'SW09', 'HW01a', 'HW01b', 'HW01c',
# 'HW01d', 'HW01e', 'HW01f', 'HW01g', 'HW01h', 'HW01i', 'HW02a', 'HW02b', 'HW02c', 'HW02d',
# 'HW02e', 'HW02f', 'HW02g', 'HW04', 'HW05', 'HW06', 'HW07', 'NET01', 'NET02', 'NET03', 'NET06',
# 'Net07', 'MEM01', 'USB01', 'USB02', 'USB03', 'TV01', 'FNC01a', 'FNC01b', 'FNC01c', 'FNC02a',
# 'FNC02b', 'FNC02c', 'FNC02d', 'FNC03', 'FNC04', 'FNC05', 'FNC06', 'FNC07', 'FNC08', 'FNC09',
# 'FNC10', 'FNC11', 'FNC12', 'FNC13', 'FNC14', 'FNC17', 'FNC18', 'FNC19', 'FNC20', 'FNC21',
# 'FNC27', 'FNC29', 'FNC31', 'FNC35', 'FNC36', 'FNC37', 'FNC38', 'FNC39', 'FNC40', 'FNC41',
# 'FNC42', 'FNC44', 'FNC46', 'FNC47', 'FNC48', 'FNC49', 'FNC50', 'FNC51', 'FNC52', 'FNC53',
#  'FNC54', 'FNC55', 'B1_1', 'B1_2', 'B1_3', 'B1_4', 'B1_5', 'B2', 'B3', 'B4', 'B5', 'B6',
# 'B7', 'B9', 'B10', 'B11', 'B12', 'B13', 'B15', 'B16', 'B18', 'B19', 'B19_2', 'B20', 'P1',
# 'UXP01', 'UXP02', 'UXP03', 'W1', 'W2', 'W3', 'W4', 'W6', 'W7', 'W8', 'L1', 'New01']
str_list = []
for i in range(len(name_list)):
    string = "create table if not exists sheet_"+str(10000+i)+"_"+name_list[i]+" like mst01;"
    str_list.append(string)
for i in str_list:
    print(i)



