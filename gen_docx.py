#!/usr/bin/env python3

'''
写docx
'''

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from docx.shared import Cm
import os
from datetime import datetime
from collections import namedtuple
from numpy import array
import read_xlsx
import draw_plot

ProInfo = namedtuple("ProInfo", ['name', 'area', 'code', 'contract', 'builder',\
		'supervisor', 'third_observer', 'builder_observer', 'xlsx_path', 'date'])



#日报信息头页，总体监测分析表， 现场巡查表， 沉降监测表(地表，建筑物，管线),
#测斜监测表，爆破振动监测表，平面布点图
PAGE_CATEGORY = ['header', 'overview', 'security', 'settlement_ground',\
	'settlement_buidling', 'settlement_pipeline', 'inclinometer', 'blasting',\
	'floor_layout']


def date_to_str(date_str):
	ds = date_str.strftime("%Y/%m/%d")
	return ds.split('/')[0] + '年' + ds.split('/')[1] + '月' + ds.split('/')[2] + '日'

def d2s(date_str):
	ds = date_str.strftime("%Y/%m/%d")
	return ds.split('/')[1] + '月' + ds.split('/')[2] + '日'

def d_s(date_str):
	ds = date_str.strftime("%Y/%m/%d")
	return ds

def get_file_list(dir,file_list):
	'''
	获取目录dir下的所有文件名(文件路径)
	略过隐藏的特殊文件
	支持子目录
	'''
	try:
		new_dir = dir
		if os.path.isfile(dir):
			file_list.append(dir)
		elif os.path.isdir(dir):
			for s in os.listdir(dir):
				#略过特殊字符开头的文件或者文件夹
				if not s[0].isdigit() and not s[0].isalpha():
					#logger.warning("Hidden file:%s"%(s))
					#logger.warning("Hidden file:{}".format(s))
					if s != '.':
						continue
				new_dir = os.path.join(dir,s)
				get_file_list(new_dir,file_list)
		else:
			pass
	except Exception as e:
		#logger.warning(e)
		print("warning,e:",e)
	return file_list

class MyDocx(object):
	def __init__(self, docx_path, proj_info, my_xlsx):
		print("__init__ MyDocx")

		self.proj = ProInfo(*proj_info)
		self.docx = None
		self.path = docx_path
		self.date = proj_info[-1]
		self.xlsx_path = os.path.dirname(proj_info[-2])
		print("DEBUG xlsx path=",self.xlsx_path)
		self.str_date = date_to_str(self.date)
		#xlsx实例
		self.my_xlsx = my_xlsx
		self.my_plot = draw_plot.MyPlot()
	#########__init__()#####################################


	def gen_docx(self):
		'''
		生成docx文件
		'''
		print("start 'gen_docx'")

		#if not self.path or not os.path.exists(self.path):
		if not self.path:
			print("error, no available docx path")
			return
		
		self.docx = Document()

		#创建首页
		if not self.make_header_pages():
			print("DEBUG make_head_page error")
		else:
			pass

		#创建数据分析页****
		if not self.make_overview_pages():
			print("DEBUG make_overview_pages error")
		else:
			pass

		#创建现场安全巡视页
		#new section landscape
		#页面布局为横向
		pass
		if not self.make_security_pages():
			print("DEBUG make_security_pages error")
		else:
			pass


		#创建沉降监测表页*****
		#回复页面布局为纵向
		pass
		if not self.make_settlement_pages():
			print("DEBUG make_settlement_pages error")
		else:
			pass

		#测斜监测报表***
		pass


		#爆破振动监测报表
		#new section landscape
		#页面布局为横向
		new_section = d.add_section(WD_SECTION.ODD_PAGE)
		new_section.orientation = WD_ORIENT.LANDSCAPE
		new_section.page_width = Cm(27.94)
		new_section.page_height = Cm(21.59)
		pass

		#链接布点图word文件
		if not self.make_layout_pages():
			print("DEBUG make_layout_pages error")


		print("All pages done, saving docx file...")
		self.docx.save(self.path)
		return True
	#######gen_docx()####################################


	def write_header(self):
		'''
		项目信息
		'''
		d = self.docx
		d.add_paragraph("%s" % self.proj.name)

		p = d.add_paragraph("施工单位: ")
		p.add_run("%s" % self.proj.builder).underline = True
		p.add_run("    合同号: ")
		p.add_run("%s" % self.proj.contract).underline = True

		p = d.add_paragraph("监理单位: ")
		p.add_run("%s" % self.proj.supervisor).underline = True
		p.add_run("    编号: ")
		p.add_run("%s" % self.proj.code).underline = True

		p = d.add_paragraph("第三方检测单位: ")
		p.add_run("%s" % self.proj.third_observer).underline = True
	################write_header()########################


	def make_header_pages(self):
		'''
		首页
		'''
		print("start 'make_hearder_pages'")

		result = False
		d = self.docx
		d.add_heading(self.proj.name, 0)
		d.add_heading(self.proj.area, 0)

		d.add_paragraph("第三方检测日报")
		d.add_paragraph("(第%s次)" % self.proj.code.split('-')[-1])

		p = d.add_paragraph("编号: ")
		p.add_run("%s" % self.proj.code).underline = True
		p = d.add_paragraph("检测日期: ")
		p.add_run("%s" % self.str_date).underline = True

		d.add_paragraph("报警: 是      否")
		d.add_paragraph("报警内容: ")
		d.add_paragraph("")
		d.add_paragraph("")
		d.add_paragraph("")

		p = d.add_paragraph("项目负责人: ")
		p.add_run("      ").underline = True
		p = d.add_paragraph("签发日期: ")
		p.add_run("      ").underline = True
		p = d.add_paragraph("单位名称: ")
		p.add_run("  (盖章)   ").underline = True

		d.add_paragraph("")
		p = d.add_paragraph("%s" %self.str_date)

		###new page###########
		d.add_page_break()
		self.write_header()

		d.add_paragraph("第三方检测审核单")
		t = d.add_table(rows=1, cols=1, style = 'Table Grid')
		t.cell(0, 0).text = "审核意见:\n\n\n\n\n" + " "*80 +"监理工程师:"\
		 + " "*30 + "日期:" 


		result = True
		return result
	##################make_header_pages()################	


	def one_overview_table(self, area_name):
		'''
		一个区间的监测数据分析表
		'''
		print("Start 'one_overview_table' for area_name:",area_name,self.str_date)

		d = self.docx
		px = self.my_xlsx
		t = d.add_table(rows=1, cols=8, style='Table Grid')
		t.cell(0,0).text = '监测项目'
		t.cell(0,1).text = '本次\n变化\n最大点'
		t.cell(0,2).text = '日变化\n速率\n(mm/d)'
		t.cell(0,3).text = '日变量\n报警值\n(mm/d)'
		t.cell(0,4).text = '累计\n变化量\n最大点'
		t.cell(0,5).text = '累计\n变化量\n/mm'
		t.cell(0,6).text = '累计\n变量\n报警值/mm'
		t.cell(0,7).text = '累计\n变量\n控制值/mm'

		#找到这个area的所有观测项目，作为首列内容
		#all_areas_row_range = {'sheet1':{'area1':(1,10), 'area2':(11,15),...}, \
		#'sheet2':{'area4':(1,23)}}
		related_sheets = []
		for sheet in px.sheets:
			#表格格式注意，每个sheet的第一列的区间名字要一一对应
			if area_name in px.all_areas_row_range[sheet].keys():
					#还要考虑这一天的测量点有值:
					#pass
					#related_sheets.append = [sheet1,sheet2,...]
					related_sheets.append(sheet)

		print("DEBUG {}涵盖这些观测项目:{}".format(area_name,related_sheets))

		#遍历这个站所有有关的测量数据	
		for sheet in related_sheets:
		#略过这几个观测sheet，excel表格有疑问
			if sheet == '建筑物倾斜' or sheet == '安薛区间混撑' or\
			 sheet == '支撑轴力':
				print("由于excel表格以为，暂时略过 {}".format(sheet))
				continue
			print("------DEBUG, '{}, {}' 数据分析表-------".format(area_name, sheet))
			#获取这个sheet，这个日期的列坐标
			col_index,_ = px.get_item_col(sheet, self.date)
			if col_index == None:
				print("DEBUG error, col_index not found!")
				continue
			today_range_values = px.get_range_values(sheet, area_name, col_index)
			print("DEBUG '当天值列':{}".format(today_range_values))

			#获取前一天的值, 这里是否应该找到有测量值的上一次？
			last_range_values = px.get_range_values(sheet, area_name, col_index-1)
			print("DEBUG '昨天值列':{}".format(last_range_values))

			#找到其中最大变化的
			#对应位进行相减，放到新的def中，然后找到绝对值最大的，作为变换最大量
			#None的位算0
			diff_original_values = []
			diff_abs_values = []
			ln = len(today_range_values)
			for i in range(ln):
				new_v = today_range_values[i]
				last_v = last_range_values[i]
				if new_v != None and last_v != None:
					diff_original_values.append((float(new_v)-float(last_v))*1000)
				else:
					diff_original_values.append(0)

			#求出绝对值最大的值
			diff_abs_values = list(map(abs,diff_original_values))
			max_change = max(diff_abs_values)
			print("DEBUG '最大变化值'是:{}".format(max_change))
			#如果有最大值，且不为0
			if max_change != 0:
				#通过变化最大量的index和area的range找到行坐标
				#疑惑，这里有相同的最大值怎么办? 目前只找最前面的一个
				max_idx = diff_abs_values.index(max_change)
				#找到区间的行范围, 加上最大值的相对index就是最大值的row_index
				row_start, row_end = px.all_areas_row_range[sheet][area_name]
				row_index = max_idx+ row_start
				#通过行坐标，找到测量点列的测量点id
				s_index = 'B%d'%row_index
				obser_id = px.wb[sheet][s_index].value
				print("DEBUG '本次变化最大点'是:{}".format(obser_id))

				#新加一行，写入测量项目sheet，写入这个测量点id
				row = t.add_row()
				#监测项目
				row.cells[0].text = sheet
				#本次变化最大点
				row.cells[1].text = obser_id
				#日变化速率
				#保留两位小数
				row.cells[2].text = str(round(diff_original_values[max_idx],2))

				#日变量报警值空着
				row.cells[3].text = ' '

				#求本次累计值 = 当前值-初值+旧累计值
				acc_values = []
				acc_abs_values = []
				#获取'初值'这一列，在第3列
				initial_range_values = px.get_range_values(sheet, area_name, 3)
				print("DEBUG '初始值列':{}".format(initial_range_values))
				#获取'旧累计'这一列，在第4列
				old_acc_range_values = px.get_range_values(sheet, area_name, 4)
				print("DEBUG '旧累计值列':{}".format(old_acc_range_values))
				for i in range(ln):
					new_v = today_range_values[i]
					init_v = initial_range_values[i]
					oacc_v = old_acc_range_values[i]
					if oacc_v == None:
						oacc_v = 0
					if new_v != None and init_v != None:
						acc_values.append((float(new_v)-float(init_v))*1000+float(oacc_v))
					else:
						acc_values.append(0)
				print("DEBUG '本次累计值列':{}".format(acc_values))
				acc_abs_values = list(map(abs,acc_values))
				max_acc = max(acc_abs_values)
				print("DEBUG '最大累计值'是:{} ".format(max_acc))
				max_acc_idx = acc_abs_values.index(max_acc)
				row_index = max_acc_idx + row_start
				s_index = 'B%d'%row_index
				obser_id = px.wb[sheet][s_index].value
				print("DBUGG '本次累计变化最大点'是:{}".format(obser_id))
				row.cells[4].text = obser_id
				row.cells[5].text = str(round(acc_values[max_acc_idx],2))

				#累计变量报警值 空着
				row.cells[6].text = ' '
				#累计变量控制值 空着
				row.cells[7].text = ' '
			else:
				print("Debug warning, 无最大点！")

		#遍历这个站所有有关的测量数据	
		#end for sheet in related_sheets

		#爆破振动 行
		row = t.add_row()
		row.cells[0].text = '爆破振动'
		second_cell = row.cells[1]
		second_cell.merge(row.cells[7])
		#巡检 行
		row = t.add_row()
		row.cells[0].text = '巡检'
		second_cell = row.cells[1]
		second_cell.merge(row.cells[7])
		row.cells[1].text = '现场无异常情况。'
		#数据分析 行
		row = t.add_row()
		first_cell = row.cells[0]
		first_cell.merge(row.cells[7])
		s = '今日各监测项目数据变化量较小，数据在可控范围内；监测频率为1次/1d。'
		row.cells[0].text = '数据分析: ' + s

		return
	##########one_overview_table()###############################


	def make_overview_pages(self):
		'''
		监测数据分析表
		'''
		print("Start 'make_overview_pages'")

		result = False
		d = self.docx
		areas = self.my_xlsx.areas

		###new page###########
		d.add_page_break()
		p = d.add_paragraph()
		p.add_run("检测分析报告")
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p = d.add_paragraph()
		p.add_run("一、施工概况")

		###new page###########
		d.add_page_break()
		p = d.add_paragraph()
		p.add_run("二、数据分析")

		#表标题
		table_cap = "监测数据分析表"
		i = 0
		print("DEBUG start all areas:",areas)
		for area_name in areas:
			print("###开始生成 {} 数据分析表###".format(area_name))
			#test debug only one area
			if '衡山路站' in area_name:
				i += 1
				ss = '表' + '%d'%i + area_name + table_cap
				d.add_paragraph(ss).paragraph_format.alignment = \
				WD_ALIGN_PARAGRAPH.CENTER
				self.one_overview_table(area_name)
			#Test open to all	
			else:
				pass
				'''
				i += 1
				ss = '表' + '%d'%i + area_name + table_cap
				d.add_paragraph(ss).paragraph_format.alignment = \
				WD_ALIGN_PARAGRAPH.CENTER
				self.one_overview_table(area_name)
				'''

		###new page###########
		d.add_page_break()
		p = d.add_paragraph()
		p.add_run("三、结论")
		###new page###########
		d.add_page_break()
		p = d.add_paragraph()
		p.add_run("四、建议")

		ss = "监测单位:             （盖章）"
		p = d.add_paragraph()
		p.add_run(ss)
		ss = "负责人：              年  月  日 "
		p = d.add_paragraph()
		p.add_run(ss)

		result = True
		return result
	##################make_overview_pages()##############################


	def one_security_table(self, area_name):
		'''
		一个区间的现场巡查报表
		'''
		print("Start 'one_security_table' for area_name:",area_name,self.str_date)

		d = self.docx
		proj = self.proj
		ds = self.str_date

		t = d.add_table(rows=10, cols=6, style='Table Grid')
		t.cell(0,0).text = '线路名称'
		t.cell(0,1).text = proj.name
		t.cell(0,2).text = '监测标段'
		t.cell(0,3).text = ''
		t.cell(0,4).text = '工点名称'
		t.cell(0,5).text = area_name

		t.cell(1,0).text = '重点风险源'
		t.cell(1,1).merge(t.cell(1,3))
		t.cell(1,1).text = ''
		t.cell(1,2).text = '第三方监测单位'
		t.cell(1,3).text = proj.third_observer

		t.cell(2,0).text = '施工部位'
		t.cell(2,1).text = proj.name + '主体'
		t.cell(2,2).text = '天气'
		t.cell(2,3).text = ''
		t.cell(2,4).text = '施工方监测单位'
		t.cell(2,5).text = proj.builder_observer

		t.cell(3,0).text = '巡视内容'
		t.cell(3,1).text = '存在的问题描述'
		t.cell(3,2).text = '原因分析'
		t.cell(3,3).text = '可能导致的后果'
		t.cell(3,4).text = '安全状态评价'
		t.cell(3,5).text = '处置措施建议'

		t.cell(4,0).text = '开挖面地质状况'
		t.cell(4,1).text = '--'
		t.cell(4,2).text = '地质条件'
		t.cell(4,3).text = '--'
		t.cell(4,4).text = '安全可控状态'
		t.cell(4,5).text = '--'

		t.cell(5,0).text = '支护结构体系'
		t.cell(5,1).text = '--'
		t.cell(5,2).text = '--'
		t.cell(5,3).text = '--'
		t.cell(5,4).text = '安全可控状态'
		t.cell(5,5).text = '--'

		t.cell(6,0).text = '周边环境'
		t.cell(6,1).text = 'xx附近有高层建筑群'
		t.cell(6,2).text = '自身结构较稳定'
		t.cell(6,3).text = '可能导致房屋出现裂缝'
		t.cell(6,4).text = '安全可控状态'
		t.cell(6,5).text = '控制爆破药量进尺'

		t.cell(7,0).text = '监测设施'
		t.cell(7,1).merge(t.cell(7,5))
		t.cell(7,1).text = '良好'

		t.cell(8,0).text = '现场巡视人'
		t.cell(8,1).merge(t.cell(8,2))
		t.cell(8,1).text = '          '+ ds
		t.cell(8,3).text = '项目技术负责人'
		t.cell(8,4).merge(t.cell(8,5))
		t.cell(8,4).text = '          '+ ds

		t.cell(9,0).merge(t.cell(9,5))
		t.cell(9,0).text = '备注: '

	##################one_security_table()###############################


	def make_security_pages(self):
		'''
		现场巡查报表
		'''
		print("Start 'make_security_pages'")

		result = False
		d = self.docx
		areas = self.my_xlsx.areas
		proj = self.proj

		table_cap = '现场巡查报表'
		i = 0
		for area_name in areas:
			print("###开始生成 {} 现场巡查报表###".format(area_name))
			#加docx session，使用更加宽的页面布局
			#pass

			#test debug only one area
			if '衡山路站' in area_name:
				###new page###########
				d.add_page_break()
				i += 1
				ss = '表' + '%d'%i + ' 现场安全巡视表'
				p = d.add_paragraph()
				p.add_run(ss)

				p = d.add_paragraph()
				p.add_run(area_name).underline = True
				p.add_run(table_cap)
				p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
				p = d.add_paragraph()
				p.add_run('编号: ')
				p.add_run(proj.code).underline = True
				p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
				self.one_security_table(area_name)
			#Test open to all	
			else:
				pass

		result = True
		return result
	#################make_security_pages()###############################


	def find_avail_rows_dates_values(self, sheet, area_name, needed_num):
		'''
		找到7天的有效值列
		返回三个列表:
		row_list = [row_index1, row_index2,...,row_indexy]
		date_list = [date7, date6, date4,...date1]
		value_list = [[date7_v1, date7_v2,...], [date6_v1, date6_v2,...],...] len(date_list) * len(row_list)
		'''
		px = self.my_xlsx
		row_list = []
		date_list = []
		value_list = []
		each_date_values = []

		#当天的有效行数index列表和值列表
		start_row, end_row = px.all_areas_row_range[sheet][area_name]
		row_list = list(range(start_row, end_row+1))
		#获取当天日期的列坐标
		col_index, date_row_index = px.get_item_col(sheet, self.date)
		if col_index == None:
			print("DEBUG error, col_index not found!")
			return None, None, None
		today_rows, today_values = px.get_avail_rows_values(sheet, row_list, col_index)
		if len(today_rows) == 0:
			print("{},{},{}当天无有效数据".format(area_name,sheet,self.date))
			return None, None, None
		#找到一共邻近7天的有效数据，一旦某一天的某一行有none值，略过该天
		#如果不够7天的数据，直到找到不为日期那一天为止
		row_list = today_rows
		date_list.append(self.date)
		value_list.append(list(map(float,today_values)))

		already_number = 1
		col_index
		ignore_number = 0
		while 1:
			col_index -= 1
			v = px.get_value(sheet, date_row_index, col_index)
			#如果不是日期型，说明过了最早的开头了，退出循环
			if not 'datetime' in str(type(v)):
				break
			old_rows, old_values = px.get_avail_rows_values(sheet, today_rows, col_index)
			if old_rows == today_rows:
				#找到一列有效值
				date_list.append(px.get_value(sheet, date_row_index, col_index))
				value_list.append(list(map(float,old_values)))
				already_number += 1
				if already_number == needed_num:
					break
			else:
				#这一天的有none值，略过
				#如果有none值的情况隔了5天，就不在找了
				ignore_number += 1
				if ignore_number >= 5:
					break
				continue

		return row_list, date_list, value_list
	##############find_avail_rows_dates_values()#############################################


	def draw_settlement_table(self, sheet, row_list, date_list, value_list,\
		 init_values, old_acc_values, cell_row):
		'''
		画沉降监测表格
		'''
		d = self.docx
		px = self.my_xlsx

		t = d.add_table(rows=13, cols=10, style='Table Grid')
		t.cell(0,0).merge(t.cell(0,9))
		s1 = '仪器型号: '
		s2 = '               仪器出厂编号：'
		s3 = '               检定日期：'
		t.cell(0,0).text = s1+s2+s3
		t.cell(1,0).merge(t.cell(2,0))
		t.cell(1,1).merge(t.cell(1,3))
		t.cell(1,0).text = '监测点号'
		t.cell(1,1).text = '沉降变化量(mm)'
		t.cell(1,4).merge(t.cell(2,4))
		t.cell(1,4).text = '备注'
		t.cell(1,5).merge(t.cell(2,5))
		t.cell(1,5).text = '监测点号'
		t.cell(1,6).merge(t.cell(1,8))
		t.cell(1,6).text = '沉降变化量(mm)'
		t.cell(1,9).merge(t.cell(2,9))
		t.cell(1,9).text = '备注'

		t.cell(2,1).text = '上次变量'
		t.cell(2,2).text = '本次变量'
		t.cell(2,3).text = '累计变量'
		t.cell(2,6).text = '上次变量'
		t.cell(2,7).text = '本次变量'
		t.cell(2,8).text = '累计变量'

		#填入数值
		last_diffs = []
		this_diffs = []
		this_acc_diffs = []

		ln_row = len(row_list)
		ln_date = len(date_list)
		#value_list = [[date7_v1, date7_v2,...], [date6_v1, date6_v2,...],...]
		#value_list should be ln_date*ln_row
		#init_values should be ln_row*1

		if ln_date > 2:
			for i in range(ln_row):
				this_diffs.append(round((value_list[0][i] - value_list[1][i])*1000,2))
				last_diffs.append(round((value_list[1][i] - value_list[2][i])*1000,2))
				this_acc_diffs.append(round((value_list[0][i] - init_values[i])*1000+ \
					old_acc_values[i],2))

		elif ln_date ==2:
			for i in range(ln_row):
				this_diffs.append(round((value_list[0][i]-value_list[1][i])*1000,2))
				this_acc_diffs.append(round((value_list[0][i] - init_values[i])*1000+ \
					old_acc_values[i],2))
			last_diffs = [0 for j in range(ln_row)]

		elif ln_date == 1:
			for i in range(ln_row):
				this_acc_diffs.append(round((value_list[0][i] - init_values[i])*1000+ \
					old_acc_values[i],2))
			this_diffs = [0 for j in range(ln_row)]
			last_diffs = [0 for j in range(ln_row)]

		else:
			print("Error, date_list None")
			this_diffs = [0 for j in range(ln_row)]
			last_diffs = this_diffs
			this_acc_diffs = this_diffs

		#表格变化值填写
		base_index = 3
		for i in range(cell_row):
			#如果观测点数小于cell行数，则当填写完观测点即退出
			if ln_row < cell_row and i == ln_row:
				break
			#监测点号, 注意表格格式，直接从第二列获取
			t.cell(base_index+i,0).text = px.get_value(sheet,row_list[i],2)
			#上次变量
			t.cell(base_index+i,1).text = str(last_diffs[i])
			#本次变量
			t.cell(base_index+i,2).text = str(this_diffs[i])
			#累计变量
			t.cell(base_index+i,3).text = str(this_acc_diffs[i])
			#另外一侧的表格
			j = i+cell_row
			if ln_row > j:
				t.cell(base_index+i,5).text = px.get_value(sheet, row_list[j],2)
				t.cell(base_index+i,6).text = str(last_diffs[j])
				t.cell(base_index+i,7).text = str(this_diffs[j])
				t.cell(base_index+i,8).text = str(this_acc_diffs[j])

		#求七天的累计变化列表
		#array type 矩阵
		print("DEBUG array(value_list)", array(value_list))
		print("shape array = ", array(value_list).shape)
		print("DEBUG init_values", init_values)
		print("DEBUG init_vlaues le2=",len(init_values))
		all_acc_diffs = []
		all_acc_diffs = (array(value_list) - init_values)*1000 + old_acc_values
		print("DEBUG all_acc_diffs=",all_acc_diffs)
		print("DEBUG all_acc_diffs.shape=",all_acc_diffs.shape)

		#画图
		idx_list = []
		for row_idx in row_list:
			idx_list.append(px.get_value(sheet,row_idx,2))
		fig_path = self.my_plot.draw_date_plot(list(map(d_s,date_list)), \
			all_acc_diffs.transpose(), idx_list)
		if not os.path.exists(fig_path):
			print("Debug, ERROR, fig_path not exists!")
			fig_path = r'C:\Users\tarzonz\Desktop\oreport\demo.jpg'

		t.cell(11,0).text = '累计变化量曲线图'
		t.cell(11,1).merge(t.cell(11,9))

		#插入曲线图
		p = t.cell(11,1).paragraphs[0]
		run = p.add_run()
		run.add_picture(fig_path, width=Cm(13), height=Cm(5))
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

		t.cell(12,0).text = '备注'
		t.cell(12,1).merge(t.cell(12,9))
		t.cell(12,1).text = '1、“-”为下降、“+”为上升；2、监测点布设图见附图'

	##################draw_settlement_table()###################################


	def one_settlement_table(self, area_name):
		'''
		一个沉降区间的变化监测表
		步骤：
		找到初始值列，和邻近7天的有效值列
		直到找到不为日期格式的列位置，有多少列有效值就添加多少列
		如果只有一列，即当天的，那么上次变化值为0
		根据当天有效值的行数确定矩阵行数即为坐标的观测点行数范围，
		如果该行数范围内前一天有None值，则略过改天。最终要求所有
		有效值列都是有值的。如果当天的值都为None，那么跳过该sheet.
		'''
		print("Start 'one_settlement_table' for area_name:",area_name,\
			self.str_date)

		px = self.my_xlsx
		d = self.docx

		#找到这个area的所有观测项目
		related_sheets = []
		for sheet in px.sheets:
			if area_name in px.all_areas_row_range[sheet].keys():
					#related_sheets.append = [sheet1,sheet2,...]
					related_sheets.append(sheet)

		print("DEBUG {}涵盖这些观测项目:{}".format(area_name,\
			related_sheets))

		#遍历这个站所有有关的测量数据,绘制表格	
		for sheet in related_sheets:
			#略过这几个观测sheet，excel表格有疑问
			if sheet == '建筑物倾斜' or sheet == '安薛区间混撑' or\
				 sheet == '支撑轴力':
				print("由于excel表格格式疑惑，暂时略过 {}".format(sheet))
				continue
			print("DEBUG, '{}, {}' 沉降变化监测报表".format(area_name,\
				sheet))
			table_cap = area_name + sheet + '报表'

			#找到7天的有效数据值,包括行坐标，日期纵坐标和测量数据值矩阵!
			row_list = []
			date_list = []
			value_list = []
			row_list,date_list,value_list = \
			self.find_avail_rows_dates_values(sheet,area_name,7)
			print("DEBUGGGGGGGG date_list=",date_list)
			if not row_list:
				print("没有有效值")
				continue
			#从第三列获取到相应行的初始值和旧累计
			#表格格式注意，第三列和第四列为初值，旧累计
			_,initial_values = px.get_avail_rows_values(sheet, row_list, 3)
			_,old_acc_values = px.get_avail_rows_values(sheet, row_list, 4, True)

			#计算每个表能填多少个观测点
			'''
			监测点号一边8个，共两边，按照总监测点是16的x倍数，
			则以总点数/x 来分
			x/16 > x//16:
			y = x//16+1
			or
			y = x//16
			'''
			ln = len(row_list)
			split_num = 0
			total_row = 16
			if ln/total_row > ln//total_row:
				split_num = ln//total_row + 1
			else:
				split_num = ln//total_row

			start = 0
			end = 0
			for i in range(1, split_num+1):
				#最后一个就是剩下的所有的
				if i == split_num:
					end = ln
				else:
					end = (ln//split_num)* i
				sub_row_list = row_list[start:end]
				#value_list = [[],[],[],...,[]] len(date) * len(rows)
				sub_value_list = [values[start:end] for values in value_list]
				sub_initial_values = initial_values[start:end]
				sub_old_acc_values = old_acc_values[start:end]
				start = end 

				print("------DEBUG, '{}, {}' 沉降变化监测表{}/{}---".format(\
					area_name, sheet,i,split_num))
				###new page###########
				d.add_page_break()
				self.write_settlement_header(area_name)
				p = d.add_paragraph()	
				p.add_run(area_name+sheet+'监测报表'+'%d/%d'%(i,split_num))
				p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
				last_date = ''
				if len(date_list)==1:
					last_date = '初始值'
				else:
					last_date = date_to_str(date_list[1])
				p = d.add_paragraph()	
				p.add_run('上次监测时间: '+last_date)
				p.add_run('              本次监测时间: '+ self.str_date)
	
				#制表
				self.draw_settlement_table(sheet, sub_row_list, date_list,\
				 sub_value_list, sub_initial_values, sub_old_acc_values, total_row//2)
				self.write_settlement_foot()
				print("----finished-----\n")

	#############one_settlement_table()################################


	def write_settlement_header(self, area_name):
		'''
		沉降变化表头项目信息
		'''
		d = self.docx
		p = d.add_paragraph()
		p.add_run(self.proj.name)
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p = d.add_paragraph()
		p.add_run("%s主体"%area_name).underline = True
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

		p = d.add_paragraph()
		p.add_run("施工单位: ")
		p.add_run(self.proj.builder).underline = True
		p.add_run("    编号: ")
		p.add_run(self.proj.code).underline = True

		p = d.add_paragraph()
		p.add_run("监理单位: ")
		p.add_run(self.proj.supervisor).underline = True
		p = d.add_paragraph()
		p.add_run("施工监测单位: ")
		p.add_run(self.proj.builder_observer).underline = True

	################write_settlement_header()########################


	def write_settlement_foot(self):
		'''
		沉降变化表页脚信息
		'''
		d = self.docx
		p = d.add_paragraph()
		s = '现场监测人:              '
		p.add_run(s)
		s = '计算人:              '
		p.add_run(s)
		s = '校核人:              '
		p.add_run(s)

		p = d.add_paragraph()
		s = '检测项目负责人:              '
		p.add_run(s)

		s = '第三方监测单位: '
		p.add_run(s)
		p.add_run(self.proj.third_observer)
	##################write_settlementn_foot()###########################


	def make_settlement_pages(self):
		'''
		沉降变化监测表
		'''
		print("Start 'make_settlement_pages'")

		result = False
		d = self.docx
		areas = self.my_xlsx.areas
		proj = self.proj

		for area_name in areas:
			print("###开始生成 {} 沉降监测报表###".format(area_name))
			#test debug only one area
			if '衡山路站' in area_name:
				self.one_settlement_table(area_name)
			#Test open to all	
			else:
				pass
				#self.one_settlement_table(area_name)

		result = True
		return result

	################make_settlement_pages()##########################

	def make_layout_pages(self):
		'''
		把self.xlsx_path下的图片文件追加的docx中
		'''
		print("DEBUG start make_layout_pages")


		d = self.docx


		#获取文件夹下的所有文件地址:
		file_list = get_file_list(self.xlsx_path, [])

		for item in file_list:
			print(item)
			sufx = os.path.basename(item)
			if '.xlsx' in sufx or '.docx' in sufx or '.dr' in sufx or '.txt' in sufx:
				continue
			try:
				d.add_picture(item, width=Cm(23), height=Cm(18))
				print("DEBUG success inserted!")
			except:
				print("not a picture file ".format(item))

		print("Insert pictures done")
		return True
	#####################concatenate_new_docx()#######################


if __name__ == '__main__':

	#测试
	print("Start Test")
	xlsx_path = r'C:\Users\tarzonz\Desktop\演示工程A\一二工区计算表2018.1.1.xlsx' 
	docx_path = r'C:\Users\tarzonz\Desktop\演示工程A\demo1.docx'
	date_v = '2018/1/1'
	date_v = datetime.strptime(date_v, '%Y/%m/%d')
	project_info = ["青岛市地铁1号线工程", "一、二工区", "DSFJC02-RB-594", \
	"M1-ZX-2016-222", "中国中铁隧道局、十局集团有限公司",\
	 "北京铁城建设监理有限责任公司", "中国铁路设计集团有限公司",\
	 '中铁隧道勘察设计研究院有限公司', xlsx_path, date_v]


	my_xlsx = read_xlsx.MyXlsx(xlsx_path)
	my_docx = MyDocx(docx_path, project_info, my_xlsx)

	res = my_docx.gen_docx()	
	if res:
		print("'{}' has been created".format(docx_path))
		print("Done")
