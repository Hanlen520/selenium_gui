import xlwt
from datetime import datetime

# 实例化一个Workbook()对象(即excel文件)
wbk = xlwt.Workbook()
# 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)
# 获取当前日期，得到一个datetime对象如：(2016, 8, 9, 23, 12, 23, 424000)
today = datetime.today()
# 将获取到的datetime对象仅取日期如：2016-8-9
today_date_str = str(datetime.date(today))
today_time_str = str(datetime.time(today))
today_time_str = today_time_str.replace(":", "").split(".", 1)[0]
# 以当前日期作为excel名称保存。
excel_name = today_date_str+"_"+today_time_str + '.xls'
wbk.save(excel_name)