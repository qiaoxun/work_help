# -*- coding: utf-8 -*-
from xlutils.copy import copy
import xlrd
import xlwt

file1 = 'qiaoxun.xlsx'

def read_excel():
  
    wb1 = xlrd.open_workbook(filename=file1)

    sheet1 = wb1.sheet_by_index(0)
    sheet2 = wb1.sheet_by_index(1)
    print('sheet1 rows={}'.format(sheet1.nrows))
    print('sheet2 rows={}'.format(sheet2.nrows))
    sheet1_rows = sheet1.nrows
    sheet2_rows = sheet2.nrows

    data_list1 = []
    for i in range(1, sheet1_rows):
        tel = sheet1.cell(i,0).value
        money = sheet1.cell(i,1).value
        data_list1.append({'tel': tel, 'money': money})
    
    data_list2 = []
    for i in range(1, sheet2_rows):
        tel = sheet2.cell(i,0).value
        money = sheet2.cell(i,1).value
        data_list2.append({'tel': tel, 'money': money})

    for each_data in data_list1:
        tel = each_data['tel']
        money = each_data['money']
        result = list(filter(lambda x:x['tel'] == tel, data_list2))
        if len(result) > 0:
          data2 = result[0]
          if (money != data2['money']):
              print('发现不同数据 电话号码：{} 数据为 {} <> {}'.format(tel, money, data2['money']))
    
    
    extra_data_in_sheet1, extra_data_in_sheet2 = find_extra_data_in_both_sheet(data_list1, data_list2)
    print('extra_data_in_sheet1')
    print(extra_data_in_sheet1)
    print('extra_data_in_sheet2')
    print(extra_data_in_sheet2)

    return combine_two_data_list(data_list1, data_list2, extra_data_in_sheet1, extra_data_in_sheet2)


def find_extra_data_in_both_sheet(data_list1, data_list2):
    all_data_list = data_list1 + data_list2
    extra_data_in_sheet1 = []
    extra_data_in_sheet2 = []

    for each_data in all_data_list:
        tel = each_data['tel']
        result_list1 = list(filter(lambda x: x['tel'] == tel, data_list1))
        result_list2 = list(filter(lambda x: x['tel'] == tel, data_list2))

        if len(result_list1) == 0:
            extra_data_in_sheet2.append(each_data)
        if len(result_list2) == 0:
            extra_data_in_sheet1.append(each_data)

    return extra_data_in_sheet1, extra_data_in_sheet2


def combine_two_data_list(data_list1, data_list2, extra_data_in_sheet1, extra_data_in_sheet2):
    all_data = []

    for each_data in data_list1:
        tel = each_data['tel']
        money = each_data['money']
        result_list = list(filter(lambda x: x['tel'] == tel, extra_data_in_sheet1))
        if len(result_list) == 0:
            data_in_list2_result = list(filter(lambda x: x['tel'] == tel, data_list2))
            if len(data_in_list2_result) > 0:
                data_in_list2 = data_in_list2_result[0]
                combine_data = {'tel1': tel, 'money1': money, 'tel2': data_in_list2['tel'], 'money2': data_in_list2['money']}
                all_data.append(combine_data)
    
    for each_data in extra_data_in_sheet1:
        tel = each_data['tel']
        money = each_data['money']
        combine_data = {'tel1': tel, 'money1': money, 'tel2': '', 'money2': ''}
        all_data.append(combine_data)

    for each_data in extra_data_in_sheet2:
        tel = each_data['tel']
        money = each_data['money']
        combine_data = {'tel1': '', 'money1': '', 'tel2': tel, 'money2': money}
        all_data.append(combine_data)

    # print('===all_data===')
    # print(all_data)
    return all_data


def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style



def write_excel(all_data):
	f = xlwt.Workbook()
	sheet1 = f.add_sheet('结果',cell_overwrite_ok=True)
	row0 = ["充值号码","商品总价(元)","充值号码","商品总价(元)"]
	#写第一行
	for i in range(0, len(row0)):
		sheet1.write(0, i, row0[i], set_style('Times New Roman',220,True))
	
	cell_style = set_style('Times New Roman', 220, True)

	for index, each in enumerate(all_data):
		row = index + 1
		sheet1.write(row, 0, each['tel1'], cell_style)
		sheet1.write(row, 1, each['money1'], cell_style)
		sheet1.write(row, 2, each['tel2'], cell_style)
		sheet1.write(row, 3, each['money2'], cell_style)

	f.save('result.xls')

all_data = read_excel()

# write_excel(all_data)
