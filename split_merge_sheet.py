#!/usr/bin/env python
# -*- encoding: utf-8 -*-

'''
@File    :   split_merge_sheet.py
@Time    :   2023/07/03
@Author  :   HDUZN
@Version :   1.0
@Contact :   hduzn@vip.qq.com
@License :   (C)Copyright 2023-2024
@Desc    :   1.把Excel文件按Sheet工作表保存成n个Excel文件
             2.把n个Excel文件作为单独的sheet工作表合成到一个Excel文件 merged.xlsx
             3.把n个Excel文件内容合成到一个Excel文件一个sheet工作表中 merged_one_sheet.xlsx
             4.表格按A列标题“type”拆分成n个表格,保留标题行
'''

import os
import pandas as pd

# 创建文件夹
def mkdir(path):
    outdir = os.path.exists(path)
    if not outdir:
        os.makedirs(path)

# 1.把Excel文件按Sheet工作表保存成n个Excel文件
def sheets_to_files(ex_file, output_path):
    # 读取Excel文件
    xlsx = pd.ExcelFile(ex_file)

    # 遍历每个工作表
    for sheet_name in xlsx.sheet_names:
        print('Sheet名称: ' + sheet_name)
        df = pd.read_excel(ex_file, sheet_name=sheet_name)
        # 将当前工作表写入新文件
        file_name = output_path + '\\' + sheet_name + '.xlsx'
        df.to_excel(file_name, index=False, sheet_name=sheet_name)

# 2.把n个Excel文件作为单独的sheet工作表合成到一个Excel文件 merged.xlsx
def merge_to_excel(ex_file_path):
    # 将合并后的数据保存为 merged.xlsx
    merged_file = ex_file_path + '\\' + 'merged.xlsx'
    if os.path.exists(merged_file):
        os.remove(merged_file)

    # 遍历ex_file_path下所有excel文件
    ex_files = []
    for file_name in os.listdir(ex_file_path):
        # 判断文件是否为Excel文件
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # 构造文件的绝对路径
            file_path = os.path.join(ex_file_path, file_name)
            # 将文件路径添加到List中
            ex_files.append(file_path)
    # print(ex_files)

    writer = pd.ExcelWriter(merged_file, engine='xlsxwriter')
    # 创建一个空的DataFrame用于存储合并后的数据
    # merged_data = pd.DataFrame()

    # 遍历每个Excel文件
    for file in ex_files:
        # 提取文件名（不带后缀）作为Sheet名称
        ex_name = os.path.basename(file)
        sheet_name = ex_name.split('.')[0]

        df = pd.read_excel(file) # 读取Excel文件
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        # 将数据添加到合并后的DataFrame中，同时指定Sheet名称
        #merged_data = merged_data.append(df, ignore_index=True, sort=False)
        #merged_data.sheet_name = sheet_name

    writer.save()
    # merged_data.to_excel(merged_file, index=False)

# 3.把n个Excel文件内容合成到一个Excel文件一个sheet工作表中 merged_one_sheet.xlsx
def merge_to_one_sheet(ex_file_path):
    # 将合并后的数据保存为 merged_one_sheet.xlsx
    merged_file = ex_file_path + '\\' + 'merged_one_sheet.xlsx'
    if os.path.exists(merged_file):
        os.remove(merged_file)

    # 遍历ex_file_path下所有excel文件
    ex_files = []
    for file_name in os.listdir(ex_file_path):
        # 判断文件是否为Excel文件
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # 构造文件的绝对路径
            file_path = os.path.join(ex_file_path, file_name)
            # 将文件路径添加到List中
            ex_files.append(file_path)
    # print(ex_files)

    # 创建一个空的DataFrame用于存储合并后的数据
    merged_data = pd.DataFrame()
    # 遍历每个Excel文件
    for file in ex_files:
        # 提取文件名（不带后缀）作为Sheet名称
        ex_name = os.path.basename(file)
        sheet_name = ex_name.split('.')[0]

        df = pd.read_excel(file) # 读取Excel文件
        # 将数据添加到合并后的DataFrame中，同时指定Sheet名称
        # merged_data = merged_data.append(df, ignore_index=True, sort=False)
        merged_data = pd.concat([merged_data, df], ignore_index=True, sort=False)
        merged_data.sheet_name = sheet_name

    merged_data.to_excel(merged_file, index=False)

# 4.表格按A列标题“type”拆分成n个表格,保留标题行
def split_by_type(ex_file, output_path):
    data = pd.read_excel(ex_file)
    rows = data.shape[0] # 获取内容行数,不包含标题 shape[1]获取列数
    print(rows)
    col_a_title = data.columns[0] # 获取A列的标题
    # print(col_a_title)
    type_list = data[col_a_title].unique().tolist() # 获取A列的所有内容并去重
    # print(type_list)

    for type in type_list:
        df = data.loc[data[col_a_title] == type]
        output_file = output_path + '\\' + str(type) + '.xlsx'
        df.to_excel(output_file, index=False)

# 5.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
def split_by_type_to_one(ex_file, output_file):
    if os.path.exists(output_file):
        os.remove(output_file)

    data = pd.read_excel(ex_file)
    col_a_title = data.columns[0] # 获取A列的标题
    type_list = data[col_a_title].unique().tolist() # 获取A列的所有内容并去重

    # 创建一个ExcelWriter对象，用于写入多个Sheet
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    # writer.book = openpyxl.Workbook()

    for type in type_list:
        df = data.loc[data[col_a_title] == type]
        # 将数据写入到一个新的Sheet中
        df.to_excel(writer, sheet_name=str(type), index=False)

    # 保存Excel文件
    writer.save()

def main():
    menu = """
    ******************************* 功能菜单 ***********************************
    *  1.把Excel文件按Sheet工作表保存成n个Excel文件
    *  2.把n个Excel文件作为单独的sheet工作表合成到一个Excel文件
    *  3.把n个Excel文件内容合成到一个Excel文件一个sheet工作表中
    *  4.表格按A列标题“type”拆分成n个表格,保留标题行
    *  5.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
    ***************************************************************************
    """
    print(menu)
    select = input('请输入需要的功能对应的序号(例如：1)后按回车键(Enter)：')
    
    if(select=='1'):
        # 1.把Excel文件按Sheet工作表保存成n个Excel文件
        ex_path = r'.\ex_files'
        ex_file = ex_path + '\\' + '分sheet为新文件-模板.xlsx'
        output_path = r'.\output'
        mkdir(output_path)
        sheets_to_files(ex_file, output_path)
    elif(select=='2'):
        # 2.把n个Excel文件作为单独的sheet工作表合成到一个Excel文件
        ex_file_path = r'.\ex_files\sheets' # n个Excel文件目录
        merge_to_excel(ex_file_path)
    elif(select=='3'):
        # 3.把n个Excel文件内容合成到一个Excel文件一个sheet工作表中
        ex_file_path = r'.\ex_files\sheets' # n个Excel文件
        merge_to_one_sheet(ex_file_path)
    elif(select=='4'):
        # 4.表格按A列标题“type”拆分成n个表格,保留标题行
        ex_path = r'.\ex_files'
        ex_file = ex_path + '\\' + '按标题type拆分-模板.xlsx'
        output_path = r'.\output'
        mkdir(output_path)
        split_by_type(ex_file, output_path)
    elif(select=='5'):
        # 5.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
        ex_path = r'.\ex_files'
        ex_file = ex_path + '\\' + '按标题type拆分-模板.xlsx'
        output_path = r'.\output'
        mkdir(output_path)
        output_file = output_path + '\\' + 'new_file.xlsx'
        split_by_type_to_one(ex_file, output_file)
    else:
        print('No this function num!')

    print('------------------------fine!')

main()