import xlrd

'''excel工具方法'''
def getRowIndex(table, rowName):
    rowIndex = None
    for i in range(table.nrows):
        if(table.cell_value(i,0) == rowName):
            rowIndex = i
            break
    return rowIndex


def getColumnIndex(table, columnName):
    columnIndex = None
    # 同时配合也要修改这里为table.ncols或table.nrows
    for i in range(table.ncols):
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex


def readExcelDataByName(fileName, sheetName):
    table = None
    errorMsg = ""
    try:
        data = xlrd.open_workbook(fileName)
        table = data.sheet_by_name(sheetName)
    except Exception as msg:
        errorMsg = msg
    return table, errorMsg


def readByColName(path,sheetName,colName):
    try:
        table = readExcelDataByName(path, sheetName)[0]
        # 修改这里的两个参数的位置可以确定是根据行名还是根据列表读取
        value = table.cell_value(1,getColumnIndex(table, colName),)
    except Exception as msg:
        raise RuntimeError('readByColName()方法为以第一行数据为键取第二行数据的值！')
    return value


def readByRowName(path,sheetName,rawName):
    try:
        table = readExcelDataByName(path, sheetName)[0]
        # 修改这里的两个参数的位置可以确定是根据行名还是根据列表读取
        value = table.cell_value(getRowIndex(table, rawName),1)
    except Exception as msg:
        raise RuntimeError('readByRowName()方法为以第一列数据为键取第二列数据的值')
    return value


if __name__ == '__main__':
    path = "C:\\Users\\Administrator\\Desktop\\epicc.xls"
    value1 = readByColName(path, "老非车险接口", "EAU_DY")
    print(value1)