#!/usr/bin/env python3

'''
写docx
'''

from docx import Document
from collections import namedtuple
import os

from docx.enum.text import WD_ALIGN_PARAGRAPH
import read_xlsx

ProInfo = namedtuple("ProInfo", ['name', 'area', 'code', 'contract', 'builder',\
		'supervisor', 'observer', 'xlsx_path', 'date'])



#日报信息头页，总体监测分析表， 现场巡查表， 沉降监测表(地表，建筑物，管线),
#测斜监测表，爆破振动监测表，平面布点图
PAGE_CATEGORY = ['header', 'overview', 'environment', 'settlement_ground',\
	'settlement_buidling', 'settlement_pipeline', 'inclinometer', 'blasting',\
	'floor_layout']


class MyDocx(object):
	def __init__(self, docx_path, proj_info, my_xlsx):
		print("__init__ MyDocx")

		self.proj = ProInfo(*proj_info)
		self.docx = None
		self.path = docx_path
		self.date = proj_info[-1]
		#xlsx实例
		self.my_xlsx = my_xlsx
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

		#创建数据分析页
		if not self.make_overview_pages():
			print("DEBUG make_overview_pages error")
		else:
			pass

		#创建现场安全巡视页
		if not self.make_security_pages():
			print("DEBUG make_security_pages error")
		else:
			pass

		print("Saving...")
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
		p.add_run("%s" % self.proj.observer).underline = True
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
		p.add_run("%s" % self.proj.date).underline = True

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
		p = d.add_paragraph("%s" %self.proj.date)

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
		print("Start 'one_overview_table' for area_name:",area_name,self.date)

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
			col_index = px.get_item_col(sheet, self.date)
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
		print("Start 'one_security_table' for area_name:",area_name,self.date)

		d = self.docx
		proj = self.proj
		ds = proj.date.strftime("%Y/%m/%d")
		dss = ds.split('/')[0] + '年' + ds.split('/')[1] + '月' + \
		ds.split('/')[2] + '日'

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
		t.cell(1,3).text = proj.observer

		t.cell(2,0).text = '施工部位'
		t.cell(2,1).text = proj.name + '主体'
		t.cell(2,2).text = '天气'
		t.cell(2,3).text = ''
		t.cell(2,4).text = '施工方监测单位'
		t.cell(2,5).text = proj.builder

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
		t.cell(8,1).text = '          '+ dss
		t.cell(8,3).text = '项目技术负责人'
		t.cell(8,4).merge(t.cell(8,5))
		t.cell(8,4).text = '          '+ dss

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
			#test debug only one area
			if '衡山路站' in area_name:
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



if __name__ == '__main__':

	from datetime import datetime
	#测试
	date_v = '2018/1/1'
	date_v = datetime.strptime(date_v, '%Y/%m/%d')
	xlsx_path = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx' 
	project_info = ["青岛市地铁1号线工程", "一、二工区", "DSFJC02-RB-594", \
	"M1-ZX-2016-222", "中国中铁隧道局、十局集团有限公司",\
	 "北京铁城建设监理有限责任公司", "中国铁路设计集团有限公司",\
	  xlsx_path, date_v]

	docx_path = r'C:\Users\tarzonz\Desktop\demo1.docx'
	#with open(docx_path, 'wb') as fobj:
	#	pass

	data_source = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx'
	my_xlsx = read_xlsx.MyXlsx(xlsx_path)
	my_docx = MyDocx(docx_path, project_info, my_xlsx)
	res = my_docx.gen_docx()	

	if res:
		print("'{}' has been created".format(docx_path))
		print("Done")