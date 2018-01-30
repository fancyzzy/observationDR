#!/usr/bin/env python3

'''
写docx
'''

from docx import Document
from collections import namedtuple
import os

from docx.enum.text import WD_ALIGN_PARAGRAPH

ProInfo = namedtuple("ProInfo", ['name', 'area', 'code', 'contract', 'builder',\
		'supervisor', 'observor', 'xlsx_path', 'date'])



#日报信息头页，总体监测分析表， 现场巡查表， 沉降监测表(地表，建筑物，管线),
#测斜监测表，爆破振动监测表，平面布点图
PAGE_CATEGORY = ['header', 'overview', 'environment', 'settlement_ground',\
	'settlement_buidling', 'settlement_pipeline', 'inclinometer', 'blasting',\
	'floor_layout']


class MyDocx(object):
	def __init__(self, docx_path, proj_info, data_source):

		self.proj = ProInfo(*proj_info)
		self.docx = None
		self.path = docx_path

	def gen_docx(self):

		if not self.path or not os.path.exists(self.path):
			print("error, not an available path")
			return
		
		self.docx = Document()
		if not self.make_header_pages():
			print("DEBUG make_head_page error")

		self.docx.save(self.path)

		return True

	def project_header(self, d):
		'''
		项目信息
		'''
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

		self.project_header(d)
		d.add_paragraph("第三方检测审核单")
		t = d.add_table(rows=1, cols=1, style = 'Table Grid')
		t.cell(0, 0).text = "审核意见:\n\n\n\n\n" + " "*80 +"监理工程师:" + " "*30 + "日期:" 

		###page###########
		d.add_page_break()
		###page###########

		p = d.add_paragraph("检测分析报告")
		p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p = d.add_paragraph()
		p.add_run("一、施工概况")

		result = True
		return result
	##################make_header_pages()################	


	def make_overview_pages(self):
		'''
		监测数据分析表
		'''
		result = False
		d = self.docx
		p = d.add_paragraph()
		p.add_run("一、数据分析")

		pass



	##################make_head_page()################	

if __name__ == '__main__':


	xlsx_path = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx' 
	project_info = ["青岛市地铁1号线工程", "一、二工区", "DSFJC02-RB-594", "M1-ZX-2016-222", \
	"中国中铁隧道局、十局集团有限公司", "北京铁城建设监理有限责任公司",\
	"中国铁路设计集团有限公司", xlsx_path, "2018/1/1"]

	docx_path = r'C:\Users\tarzonz\Desktop\demo1.docx'

	with open(docx_path, 'wb') as fobj:
		pass

	data_source = r'C:\Users\tarzonz\Desktop\oreport\一二工区计算表2018.1.1.xlsx'

	my_docx = MyDocx(docx_path, project_info, data_source)
	res = my_docx.gen_docx()	
	if res:
		print("'{}' has been created".format(docx_path))
		print("Done")