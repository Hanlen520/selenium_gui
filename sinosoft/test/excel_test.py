import xlrd


def getExcelList():
    workbook = xlrd.open_workbook('测试数据.xls')
    sheet_names = workbook.sheet_names()
    temp_list = []
    t_list = []
    t = ""
    tep_list = []
    final_list = []
    list_sheet = workbook.sheet_by_name("list")
    for i in range(list_sheet.nrows):
        rows = list_sheet.row_values(i)  # 获取行内容
        temp_list.append(rows)
        if "_" in rows[9]:
            row9 = rows[9].split("_", 1);
            risk_sheet = workbook.sheet_by_name(row9[0])
            for k in range(risk_sheet.nrows):
                rows1 = risk_sheet.row_values(k)  # 获取行内容
                if rows1[1] == row9[1]:
                    t = t + ','.join(rows1) + "|"
                    temp_list[i][9] = t
            t = ""
    # print(list)
    for index in range(len(temp_list)):
        if index == 0:
            continue
        # print(temp_list[index])
        risk_list = temp_list[index][9].split("|", 30)
        for j in range(len(risk_list) - 1):
            l = risk_list[j].split(",")
            t_list.append(l)
            # print(l)
        for i in range(len(temp_list[index])):
            if i == 9:
                tep_list.append(t_list)
            else:
                tep_list.append(temp_list[index][i])
        final_list.append(tep_list)
        tep_list = []
        t_list = []
    # print(len(final_list))
    # print(final_list)
    return final_list


if __name__ == '__main__':
    my_list = getExcelList()
    for i in range(len(my_list)):
        print(my_list[i])
        for j in range(len(my_list[i][9])):
            # print(j)
            print(my_list[i][9][j])