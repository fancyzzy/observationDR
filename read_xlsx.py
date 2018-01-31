#!/usr/bin/env python3

'''
读取xlsx数据源
'''
import openpyxl


class MyXlsx(object):
	def __init__(self, xlsx_path):
		self.path = xlsx_path

		print("DEBUG, start to load datasource...")
		self.xlsx = openpyxl.load_workbook(xlsx_path)
		print("DEBUG, load finished.")

		#获得所有的sheet页名单, 即观测项目
		self.sheets_list =  self.xlsx.sheetnames[:]
		print("DEBUG observation_list=", self.sheets_list)

		sheet_ranges = self.xlsx['地表沉降']
		print("DEBUG sheet_ranges", sheet_ranges)
		print("DEBUG sheet_ranges B3", sheet_ranges['B3'].value)

		#test one sheet_area_range
		ground_sheet_areas = self.get_one_sheet_areas_range('地表沉降')
		print("DEBUG test, ground_area: ", ground_sheet_areas)


		'''
		#获取所有sheet的区间的行范围, 以字典形式为数据索引
		#{'sheet1':{'area1':(1,10), 'area2':(11,15),...}, 'sheet2':{'area4':(1,23)}}
		self.all_row_range = self.get_all_sheets_areas_range()
		#获得'地表沉降'页的A列元素, 作为总区间汇总
		#['area1','area2','area3',...'area8']
		#疑问，是否可以用地表沉降的区间作为全部区间?
		sheet_name = '地表沉降'
		d_areas = self.all_row_range(sheet_name)
		self.total_areas = list(d_areas.keys())
		'''
	##################__init__()##############################	


	def get_all_sheets_areas_range(self):
		'''
		获取所有sheet的area名和其观测点的行数范围
		用一个字典嵌套字典作为将来获取表信息的索引数据库，例如下:
		all_sheets_areas_range = {'sheet1':{'area1':(1,10), 'area2':(11,15),...}, 'sheet2':{'area4':(1,23)}}

		'''
		all_sheets_areas_range = {}
		sheet_areas_range = {}

		for sheet in self.sheets_list:
			sheet_areas_range = self.get_one_sheet_areas_range(sheet)
			all_sheets_areas_range[sheet] = sheet_areas_range

		return all_sheets_areas_range
	#############get_all_sheets_areas_range()#################		


	def get_one_sheet_areas_range(self, sheet_name):
		'''
		获取一个sheet的所有区间的行范围
		返回值: sheet_areas_range = {'area1':(1,10), 'area2':(11,15),...'area4':(30,35)} 
		含义是{区间名:(起始行数,结束行数)

		'''
		sheet_areas_range = {}
		area_name = ''
		print("DEBUG get_one_sheet_areas_range start, sheet_name=", sheet_name)
		sheet = self.xlsx.get_sheet_by_name(sheet_name)
		print("DEBUG sheet got: ", sheet)
		print("DEBUG sheet got type: ", type(sheet))
		start = 0
		start_count = False
		#最大支持500行的观测点个数
		for i in range(1, 500):
			#表格格式注意, 区间必须是在A列, A列开始为空
			v_1_col = sheet.cell(row=i, column=1).value
			v_2_col = sheet.cell(row=i, column=2).value
			if v_1_col != None and (not start_count):
				area_name = v_1_col
				start = i
				start_count = True

			#发现新的area, 保存之前area的信息
			elif v_1_col != None and start_count:
				sheet_areas_range[area_name] = (start,i)
				area__name = v_1_col
				start = i

			#最后一个的后一个
			#表格格式注意, 观测点之间不能有空行
			elif v_1_col == None and v_2_col == None:
				sheet_areas_range[area_name] = (start,i)
				break

		print("DEBUG sheet '{}' has areas_range: {}".format(\
			sheet_name, sheet_areas_range))
	#######get_one_sheet_areas_range()###########################



if __name__ == '__main__':

	print("start")
	xlsx_path = r"C:\Users\tarzonz\Desktop\演示工程A\一二工区计算表2018.1.1.xlsx"

	my_xlsx = MyXlsx(xlsx_path)

	#print("DEBUG ground_sheet_areas=",my_xlsx.total_areas)



	print("DEBUG done")

