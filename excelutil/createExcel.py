import xlwt
import xlrd
from datetime import datetime
from xlutils.copy import copy

# 实例化一个Workbook()对象(即excel文件)
wbk = xlwt.Workbook()
# 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
sheet = wbk.add_sheet('sheet1', cell_overwrite_ok=True)
sheet.write(0,0,'测试脚本')
sheet.write(0,1,'执行数据')
sheet.write(0,2,'开始时间')
sheet.write(0,3,'结束时间')
sheet.write(0,4,'执行时长(秒)')
sheet.write(0,5,'执行结果')
sheet.write(0,6,'错误消息')
sheet.write(0,7,'错误截图')
# sheet.write_merge(2,2,0,0,1)
# 获取当前日期，得到一个datetime对象如：(2016, 8, 9, 23, 12, 23, 424000)
today = datetime.today()
# 将获取到的datetime对象仅取日期如：2016-8-9
today_date_str = str(datetime.date(today))
today_time_str = str(datetime.time(today))
today_time_str = today_time_str.replace(":", "").split(".", 1)[0]
# 以当前日期作为excel名称保存。
excel_name = today_date_str+"_"+today_time_str + '.xls'
wbk.save(excel_name)

# 打开想要更改的excel文件
old_excel = xlrd.open_workbook('2018-04-23_155602.xls', formatting_info=True)
# 将操作文件对象拷贝，变成可写的workbook对象
new_excel = copy(old_excel)
# 获得第一个sheet的对象
ws = new_excel.get_sheet(0)
# 写入数据
ws.write(2, 0, '第一行，第一列')
ws.write(2, 1, '第一行，第二列')
ws.write(2, 2, '第一行，第三列')
ws.write(3, 0, '第二行，第一列')
ws.write(3, 1, '第二行，第二列')
ws.write(3, 2, '第二行，第三列')
# 另存为excel文件，并将文件命名
new_excel.save('2018-04-23_155602.xls')

