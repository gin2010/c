# -*- coding: utf-8 -*-
# @Date : 2020-01-22
# @Author : water
# @Version  : v1.0
# @Desc  : 合并excel表

import xlrd,xlwt
import time,os

TIME_STR = time.strftime("%Y%m%d")
PATH = os.path.dirname(os.path.abspath(__file__))

def get_name():
    excel_lists = []

    for root, dirs, files in os.walk(PATH, topdown=False):
        for name in files:
            each_file =  name
            if each_file.split('.')[-1] in ['xlsx','xls']:
                excel_lists.append(each_file)
    return excel_lists


def merge_excel(excel_files):
    datas = dict()
    for excel_file in excel_files:
        # 打开每一个excel
        wb = xlrd.open_workbook(os.path.join(PATH,excel_file))
        name = excel_file.split('_')[0]  # 取名字
        sheets = wb.sheet_names()
        for sheet in sheets:
            # 打开一个sheet
            ws = wb.sheet_by_name(sheet)
            if datas.get(sheet) == None:
                # 存放对应sheet的字典key初始化
                datas[sheet] = list()
                title = ws.row_values(0)
                # 如果sheet里没有姓名，插入姓名列
                if '姓名' not in title:
                    title.insert(0,"姓名")
                datas[sheet].append(title)
            # 读取每一行数据并添加到datas对应key的列表中
            for r in range(1,ws.nrows):
                each_data = ws.row_values(r)
                # 将名字添加到数据表中
                if name not in each_data:
                    each_data.insert(0,name)
                datas[sheet].append(each_data)
    return datas


def main():
    excel_lists = get_name()
    wb_name = excel_lists[0].split('_', 1)[1].replace('xlsx','xls')
    print(wb_name)
    out_excel = xlwt.Workbook()
    datas = merge_excel(excel_lists)
    # print(datas)
    # 循环datas每个key，并将值写入对应的sheet中
    for data in datas.keys():
        # print(datas[data])
        # 根据key的值新建sheet
        sheet = out_excel.add_sheet(data,cell_overwrite_ok=True)
        for r in range(len(datas[data])):
            for c in range(len(datas[data][r])):
                sheet.write(r,c,datas[data][r][c])
                if datas[data][0][0] == "序号" and r > 0:
                    sheet.write(r, 0, r)
    out_excel.save(os.path.join(PATH,'result' + "_" + wb_name))


if __name__ == "__main__":
    main()

