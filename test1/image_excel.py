#!/usr/bin/python
# -*- codding: cp936 -*-
import xlsxwriter
book = xlsxwriter.Workbook('pict.xls')
sheet = book.add_worksheet('demo')
sheet.insert_image('D4','a.png',{"x_scale":0.02,"y_scale":0.02})
sheet.insert_image('D5','a.png',{"x_scale":0.02,"y_scale":0.02})
book.close()


# import xlwt
# file = xlwt.Workbook()                # 注意这里的Workbook首字母是大写
# table = file.add_sheet('sheet1')  # 新建一个sheet
# table.insert_bitmap('a.png', 2, 2)          # 写入数据table.write(行,列,value)
#
# # 如果对一个单元格重复操作，会引发
# # returns error:
# # Exception: Attempt to overwrite cell:
# # sheetname=u'sheet 1' rowx=0 colx=0
# # 所以在打开时加cell_overwrite_ok=True解决
#
# table = file.add_sheet('sheet name',cell_overwrite_ok=True)
# file.save('demo.xls')     # 保存文件
