#!/usr/bin/env python3

'''
读取xlsx数据源
解析所有sheet名称
获取每个sheet的所有区间的名字和行范围
'''
import openpyxl
from datetime import datetime


class MyXlsx(object):
	def __init__(self, xlsx_path):
		self.path = xlsx_path

		print("init to load datasource...")
		self.wb = openpyxl.load_workbook(xlsx_path)
		print("xlsx load finished.")

		#获得所有的sheet页名单, 即观测项目
		self.sheets =  self.wb.sheetnames[:]

		#预先保存所有sheet的最大列数
		self.d_maxcol = {}
		for sheet in self.sheets:
			self.d_maxcol[sheet] = len(tuple(self.wb[sheet].columns))


		#获取所有sheet的区间的行范围, 以字典形式为数据索引
		#{'sheet1':{'area1':(1,10), 'area2':(11,15),...}, 'sheet2':{'area4':(1,23)}}
		self.all_areas_row_range = self.get_all_sheets_areas_range()

		#获得'地表沉降'页的A列元素, 作为总区间汇总
		#['area1','area2','area3',...'area8']
		#疑问，是否可以用地表沉降的区间作为全部区间?
		sheet_name = '地表沉降'
		d_areas = self.all_areas_row_range[sheet_name]
		self.areas = list(d_areas.keys())
	##################__init__()##############################	


	def get_all_sheets_areas_range(self):
		'''
		获取所有sheet的area名和其观测点的行数范围
		用一个字典嵌套字典作为将来获取表信息的索引数据库，例如下:
		all_sheets_areas_range = {'sheet1':{'area1':(1,10), 'area2':(11,15),...}, 'sheet2':{'area4':(1,23)}}

		'''
		all_sheets_areas_range = {}
		sheet_areas_range = {}

		for sheet in self.sheets:
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
		sheet = self.wb[sheet_name]
		start = 0
		start_count = False
		#最大支持500行的观测点个数
		for i in range(1, 500):
			#表格格式注意, 区间必须是在A列, A列开始为空
			v_1_col = sheet.cell(row=i, column=1).value
			#print("DEBUG i:{}, v_1_col:{}".format(i, v_1_col))
			v_2_col = sheet.cell(row=i, column=2).value
			if v_1_col != None and (not start_count):
				area_name = v_1_col
				#print("DEBUG found an area area_name=", v_1_col)
				start = i
				start_count = True

			#发现新的area, 保存之前area的name,和上一行的行号i-1
			elif v_1_col != None and start_count:
				sheet_areas_range[area_name] = (start,i-1)
				#print("DEBUG added an area{},({})".format(area_name, (start,i-1)))
				#start 重新开始记录
				area_name = v_1_col
				start = i

			#最后一行结束以2列的值全为空，为结束，并且已经开始计数
			#表格格式注意, 观测点之间不能有空行
			elif v_1_col == None and v_2_col == None and start_count:
				sheet_areas_range[area_name] = (start,i-1)
				#print("DEBUG added an area{},({})".format(area_name, (start,i-1)))
				break

			else:
				#继续寻找
				continue

		return sheet_areas_range
	#######get_one_sheet_areas_range()###########################


	def get_item_col(self, sheet, item):
		'''
		寻找第一第二排的某一项的在sheet里的列坐标
		'''
		sh = self.wb[sheet]
		#从后往前找
		for i in range(self.d_maxcol[sheet], 0, -1):
			#查找前两排，找到这个值，返回这个值的列坐标
			#表格格式注意，日期用日期格式，python里面是
			#datetime.datetime类型
			#每个sheet的行表头在row1和row2
			if item == sh.cell(1,i).value or item == sh.cell(2,i).value:
				print("find column index! i=", i)
				return i

		return None

	#########get_item_col()##########################################


	def get_range_values(self, sheet, area_name, col):
		'''
		通过sheet，area和列坐标
		返回area这一列的所有值, 到一个列表[]
		列入返回1月1日这一列的衡山路站的测量值
		'''
		sh = self.wb[sheet]
		range_values = []

		#获取area的测量点行范围row_range
		start_row, end_row = self.all_areas_row_range[sheet][area_name]

		for i in range(start_row, end_row+1):
			range_values.append(sh.cell(i, col).value)

		return
	#########get_values()######################################


if __name__ == '__main__':

	print("start")
	xlsx_path = r"C:\Users\tarzonz\Desktop\演示工程A\一二工区计算表2018.1.1.xlsx"

	my_xlsx = MyXlsx(xlsx_path)

	#测试获得所有列的数目
	#all_column = my_xlsx.wb['地表沉降'].columns
	#print("DEBUG all_row=",len(all_row))
	#print("DEBUG all_col=",len(tuple(all_column)))

	#print("DEBUG ground_sheet_areas=",my_xlsx.total_areas)



	#测试找某一个日期的列坐标
	ws = my_xlsx.wb['地表沉降']

	ss = '2018/1/1'
	sd = datetime.strptime(ss, '%Y/%m/%d')
	print("DEBUG sd=",sd)

	i = my_xlsx.get_item_col('地表沉降', sd)
	print("i=",i)

	'''
	#获取某一个单元格的值
	dd = ws['WE2'].value
	ddd = ws.cell(2, 603).value
	print("Debug ddd=",ddd)
	print("DEBUG type(ddd)=",type(ddd))
	'''



	print("DEBUG done")
