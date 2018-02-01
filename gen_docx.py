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
		'supervisor', 'observor', 'xlsx_path', 'date'])



#日报信息头页，总体监测分析表， 现场巡查表， 沉降监测表(地表，建筑物，管线),
#测斜监测表，爆破振动监测表，平面布点图
PAGE_CATEGORY = ['header', 'overview', 'environment', 'settlement_ground',\
	'settlement_buidling', 'settlement_pipeline', 'inclinometer', 'blasting',\
	'floor_layout']


class MyDocx(object):
	def __init__(self, docx_path, proj_info, my_xlsx):

		self.proj = ProInfo(*proj_info)
		self.docx = None
		self.path = docx_path
		self.date = proj_info[-1]
		#xlsx实例
		self.my_xlsx = my_xlsx

	def gen_docx(self):
		'''
		生成docx文件
		'''

		if not self.path or not os.path.exists(self.path):
			print("error, not an available path")
			return
		
		self.docx = Document()

		#创建首页
		print("start making header pages")
		if not self.make_header_pages():
			print("DEBUG make_head_page error")
		else:
			pass

		#创建数据分析页
		print("start making overview pages")
		if not self.make_overview_pages():
			print("DEBUG make_overview_pages error")
		else:
			pass


		print("Saving...")
		self.docx.save(self.path)
		return True
	#######gen_docx()########################

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
		p.add_run("%s" % self.proj.observor).underline = True


	def make_header_pages(self):
		'''
		首页
		'''
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

		###page###########
		d.add_page_break()
		###page###########

		self.write_header()

		d.add_paragraph("第三方检测审核单")
		t = d.add_table(rows=1, cols=1, style = 'Table Grid')
		t.cell(0, 0).text = "审核意见:\n\n\n\n\n" + " "*80 +"监理工程师:" + " "*30 + "日期:" 


		result = True
		return result
	##################make_header_pages()################	


	def one_overview_table(self, area_name):
		'''
		一个区间的监测数据分析表
		'''
		d = self.docx
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
		all_range = self.my_xlsx.all_areas_row_range
		for sheet in self.my_xlsx.sheets:
			if area_name in all_range[sheet].keys():
				#还要考虑这一天的测量点有值:
				pass
				#related_sheets.append = [{sheet1:(1,10)},{sheet2:(3,15)}...]
				related_sheets.append({sheet:all_range[sheet][area_name]})

		print("DEBUG area: {}, related_sheets: {}".format(area_name,related_sheets))
		for sheet in related_sheet:
			pass


		
	##########one_overview_table()############

	def make_overview_pages(self):
		'''
		监测数据分析表
		'''
		result = False
		d = self.docx
		areas = self.my_xlsx.areas

		###page###########
		d.add_page_break()
		###page###########

		p = d.add_paragraph("检测分析报告")
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p = d.add_paragraph()
		p.add_run("一、施工概况")
		p = d.add_paragraph()

		###page###########
		d.add_page_break()
		###page###########
		p.add_run("二、数据分析")

		#表标题
		table_cap = "监测数据分析表"
		i = 0
		for area_name in areas:
			#test debug only one area
			if '衡山路' in area_name:
				i += 1
				ss = '表' + '%d'%i + area_name + table_cap
				d.add_paragraph(ss).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
				self.one_overview_table(area_name)

		result = True
		return result

	##################make_head_page()################	

if __name__ == '__main__':

	#测试
	xlsx_path = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx' 
	project_info = ["青岛市地铁1号线工程", "一、二工区", "DSFJC02-RB-594", "M1-ZX-2016-222", \
	"中国中铁隧道局、十局集团有限公司", "北京铁城建设监理有限责任公司",\
	"中国铁路设计集团有限公司", xlsx_path, "2018/1/1"]

	docx_path = r'C:\Users\tarzonz\Desktop\demo1.docx'

	with open(docx_path, 'wb') as fobj:
		pass

	data_source = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx'
	my_xlsx = read_xlsx.MyXlsx(xlsx_path)
	my_docx = MyDocx(docx_path, project_info, my_xlsx)
	res = my_docx.gen_docx()	

	if res:
		print("'{}' has been created".format(docx_path))
		print("Done")