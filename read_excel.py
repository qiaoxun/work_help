# -*- coding: utf-8 -*-
import xlrd

from datetime import date,datetime

file1 = 'test.xlsx'
file2 = 'test1.xlsx'

def read_excel():
  
    wb1 = xlrd.open_workbook(filename=file1)
    wb2 = xlrd.open_workbook(filename=file2)

    sheet1 = wb1.sheet_by_index(0)
    sheet2 = wb2.sheet_by_index(0)
    print('sheet1 rows={}'.format(sheet1.nrows))
    print('sheet2 rows={}'.format(sheet2.nrows))
    sheet1_rows = sheet1.nrows
    sheet2_rows = sheet2.nrows

    data_list1 = []

    for i in range(1, sheet1_rows):
        tel = sheet1.cell(i,0).value
        money = sheet1.cell(i,1).value
        # print(sheet1.cell(i,0).value)
        # print(sheet1.cell(i,1).value)
        data_list1.append({'tel': tel, 'money': money})
    
    data_list2 = []

    for i in range(1, sheet2_rows):
        tel = sheet2.cell(i,0).value
        money = sheet2.cell(i,1).value
        # print(sheet2.cell(i,0).value)
        # print(sheet2.cell(i,1).value)
        data_list2.append({'tel': tel, 'money': money})

    # print(data_list1)
    # print(data_list2)


    for i, each_data in enumerate(data_list1):
        tel = each_data['tel']
        money = each_data['money']
        result = list(filter(lambda x:x['tel'] == tel, data_list2))
        if len(result) > 0:
          data2 = result[0]
          if (money != data2['money']):
              print('发现不同数据 电话号码：{} 数据为 {} <> {}'.format(tel, money, data2['money']))


read_excel()
