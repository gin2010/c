# -*- coding: utf-8 -*-
# @Date : 2020-01-22
# @Author : water
# @Version  : v1.0
# @Desc  : 合并excel表
#          1.修改日期显示为浮点问题，在xlwt写入表格的时候增加格式；
#          2.修改读取空sheet时的错误提示；
#          3.增加内容列比标题列多的情况

import xlrd,xlwt
import time,os
# from datetime import datetime
# from xlrd import xldate_as_tuple

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
        # print(excel_file)
        wb = xlrd.open_workbook(os.path.join(PATH,excel_file))
        name = excel_file.split('_')[0]  # 取名字
        sheets = wb.sheet_names()
        for sheet in sheets:
            # 打开一个sheet
            ws = wb.sheet_by_name(sheet)
            # 判断一个sheet是否为空
            if ws.nrows <= 1 :
                print("错误：{}表中 <{}> 内容为空，请确认是否填写正确！！".format(excel_file,sheet))
                continue
            elif datas.get(sheet) == None:
                # 存放对应sheet的字典key初始化
                datas[sheet] = list()
                title = ws.row_values(0)
                # 如果sheet的列里没有 姓名，插入姓名列
                if '姓名' not in title:
                    title.insert(0,"姓名")
                datas[sheet].append(title)
            # 读取每一行数据并添加到datas对应key的列表中
            for r in range(1,ws.nrows):
                each_data = ws.row_values(r)
                # print(each_data)
                # 增加对日期类型的判断，由于有些日期(如2020年2月2日)读进来以后ctype是2，所以无法处理
                # each_data = list()
                # for c in range(ws.ncols):
                #     if ws.cell(r, c).ctype == 3 :
                #         date_set = (xldate_as_tuple(ws.cell(r,c).value,0))
                #         if date_set[0] == 0:
                #             each_data.append(":".join(map(lambda x: str(x), date_set[3:])))
                #         else:
                #             each_data.append("a"+'/'.join(map(lambda x: str(x), date_set[:3])))
                #     else:
                #         each_data.append(ws.cell(r, c).value)
                # 将名字添加到数据表中
                if name not in each_data[:2]:
                    each_data.insert(0,name)
                datas[sheet].append(each_data)
                # print(each_data)
    return datas


def main():
    excel_lists = get_name()
    wb_name = excel_lists[0].split('_', 1)[1].replace('xlsx','xls')
    # print(wb_name)
    out_excel = xlwt.Workbook()
    # 设置日期的写入格式
    style_date = xlwt.XFStyle()
    style_date.num_format_str = 'YYYY.M.D'
    # 设置时间的写入格式
    style_time = xlwt.XFStyle()
    style_time.num_format_str = 'hh:mm:ss'

    datas = merge_excel(excel_lists)
    # print(datas)
    # 循环datas每个key，并将值写入对应的sheet中
    for data in datas.keys():
        # print(datas[data])
        # 根据key的值新建sheet
        sheet = out_excel.add_sheet(data,cell_overwrite_ok=True)

        for r in range(len(datas[data])):
            for c in range(len(datas[data][r])):
                # 根据标题的"日期"、"时间"来格式化写入该列的数据
                try:
                    title = datas[data][0][c] #标题列数可能会小于数据实际列
                except:
                    sheet.write(r, c, datas[data][r][c])
                else:
                    if "日期" in title:
                        sheet.write(r, c, datas[data][r][c], style_date)
                    elif "时间" in datas[data][0][c]:
                        sheet.write(r, c, datas[data][r][c], style_time)
                    else:
                        sheet.write(r,c,datas[data][r][c])
                # 单独写入序号
                if datas[data][0][0] == "序号" and r > 0:
                    sheet.write(r, 0, r)
    out_excel.save(os.path.join(PATH,'result' + "_" + wb_name))
    print("合并表格为：",'result' + "_" + wb_name)

if __name__ == "__main__":
    main()
    input("------合并成功-------")
