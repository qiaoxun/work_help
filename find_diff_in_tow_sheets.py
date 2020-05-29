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
        duplicate_data = is_data_in_list(tel, data_list1)
        if len(duplicate_data) > 0:
            dup_data = duplicate_data[0]
            dup_data['money'] = float(dup_data['money']) + float(money)
            dup_data['duplicate'] = dup_data['duplicate'] + ' + '  + str(money)
        else:
            data_list1.append({'tel': tel, 'money': money, 'duplicate': str(money)})
    
    data_list2 = []
    for i in range(1, sheet2_rows):
        tel = sheet2.cell(i,0).value
        money = sheet2.cell(i,1).value
        duplicate_data = is_data_in_list(tel, data_list2)
        if len(duplicate_data) > 0:
            dup_data = duplicate_data[0]
            dup_data['money'] = float(dup_data['money']) + float(money)
            dup_data['duplicate'] = dup_data['duplicate'] + ' + '  + str(money)
        else:
            data_list2.append({'tel': tel, 'money': money, 'duplicate': str(money)})


    for each_data in data_list1:
        tel = each_data['tel']
        money = each_data['money']
        result = list(filter(lambda x:x['tel'] == tel, data_list2))
        if len(result) > 0:
          data2 = result[0]
          if (money != data2['money']):
              print('发现不同数据 电话号码：{} 数据为 {} <> {}'.format(tel, money, data2['money']))
    
    
    extra_data_in_sheet1, extra_data_in_sheet2 = find_extra_data_in_both_sheet(data_list1, data_list2)
    # print('extra_data_in_sheet1')
    # print(extra_data_in_sheet1)
    # print('extra_data_in_sheet2')
    # print(extra_data_in_sheet2)

    return combine_two_data_list(data_list1, data_list2, extra_data_in_sheet1, extra_data_in_sheet2)

def is_data_in_list(tel, data_list):
    return list(filter(lambda x:x['tel'] == tel, data_list))


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
        duplicate = each_data['duplicate']
        result_list = list(filter(lambda x: x['tel'] == tel, extra_data_in_sheet1))
        if len(result_list) == 0:
            data_in_list2_result = list(filter(lambda x: x['tel'] == tel, data_list2))
            if len(data_in_list2_result) > 0:
                data_in_list2 = data_in_list2_result[0]
                combine_data = {'tel1': tel, 'money1': money, 'duplicate1': duplicate, 'tel2': data_in_list2['tel'], 'money2': data_in_list2['money'], 'duplicate2': data_in_list2['duplicate']}
                all_data.append(combine_data)
    
    for each_data in extra_data_in_sheet1:
        tel = each_data['tel']
        money = each_data['money']
        duplicate = each_data['duplicate']
        combine_data = {'tel1': tel, 'money1': money, 'duplicate1': duplicate, 'tel2': '', 'money2': '', 'duplicate2': ''}
        all_data.append(combine_data)

    for each_data in extra_data_in_sheet2:
        tel = each_data['tel']
        money = each_data['money']
        duplicate = each_data['duplicate']
        combine_data = {'tel1': '', 'money1': '', 'duplicate1': '', 'tel2': tel, 'money2': money, 'duplicate2': duplicate}
        all_data.append(combine_data)
    
    # print('===all_data===')
    # print(all_data)
    return all_data

def deal_with_duplicate_data(all_data):
    data_list = []
    duplicate_data_list = []
    money_diff_data = []
    for each_data in all_data:
        money1 = each_data['money1']
        money2 = each_data['money2']
        duplicate1 = each_data['duplicate1']
        duplicate2 = each_data['duplicate2']
        if money1 != money2:
            money_diff_data.append(each_data)
        elif duplicate1.find('+') > 0 or duplicate2.find('+') > 0:
            duplicate_data_list.append(each_data)
        else:
            data_list.append(each_data)
    return data_list, money_diff_data, duplicate_data_list


def set_style(name, height, bold=False, color='white'):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	if color != 'white':
		pattern = xlwt.Pattern()
		pattern.pattern = xlwt.Pattern.SOLID_PATTERN
		pattern.pattern_fore_colour = xlwt.Style.colour_map[color]
		style.pattern = pattern
	return style



def write_excel(all_data):
	print('all_data length = {}'.format(len(all_data)))
	f = xlwt.Workbook()
	sheet1 = f.add_sheet('结果',cell_overwrite_ok=True)
	row0 = ["充值号码","商品总价(元)", "重复数据", "充值号码","商品总价(元)", "重复数据"]
	#写第一行
	for i in range(0, len(row0)):
		sheet1.write(0, i, row0[i], set_style('Times New Roman',220,True))
	
	cell_style = set_style('Times New Roman', 220, True, 'white')
	yellow_cell_style = set_style('Times New Roman', 220, True, 'yellow')
	green_cell_style = set_style('Times New Roman', 220, True, 'green')

	row_index = 0
	data_list, money_diff_data, duplicate_data_list = deal_with_duplicate_data(all_data)

	print('row_index = {}'.format(row_index))
	for index, each in enumerate(data_list):
		row_index = row_index + 1
		sheet1.write(row_index, 0, each['tel1'], cell_style)
		sheet1.write(row_index, 1, each['money1'], cell_style)
		sheet1.write(row_index, 2, each['duplicate1'], cell_style)
		sheet1.write(row_index, 3, each['tel2'], cell_style)
		sheet1.write(row_index, 4, each['money2'], cell_style)
		sheet1.write(row_index, 5, each['duplicate2'], cell_style)

	print('row_index = {}'.format(row_index))
	for index, each in enumerate(duplicate_data_list):
		row_index = row_index + 1
		sheet1.write(row_index, 0, each['tel1'], cell_style)
		sheet1.write(row_index, 1, each['money1'], cell_style)
		sheet1.write(row_index, 2, each['duplicate1'], yellow_cell_style)
		sheet1.write(row_index, 3, each['tel2'], cell_style)
		sheet1.write(row_index, 4, each['money2'], cell_style)
		sheet1.write(row_index, 5, each['duplicate2'], yellow_cell_style)

	print('row_index = {}'.format(row_index))
	for index, each in enumerate(money_diff_data):
		row_index = row_index + 1
		sheet1.write(row_index, 0, each['tel1'], cell_style)
		sheet1.write(row_index, 1, each['money1'], green_cell_style)
		sheet1.write(row_index, 2, each['duplicate1'], cell_style)
		sheet1.write(row_index, 3, each['tel2'], cell_style)
		sheet1.write(row_index, 4, each['money2'], green_cell_style)
		sheet1.write(row_index, 5, each['duplicate2'], cell_style)
	print('row_index = {}'.format(row_index))

	f.save('result2.xls')

all_data = read_excel()

write_excel(all_data)
