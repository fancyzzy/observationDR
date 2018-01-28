#!/usr/bin/env python3

'''
日报信息汇总
'''
from collections import namedtuple
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import askyesnocancel


#工程项目名, 编号， 施工单位， 监理单位， 监测单位
PRO_L = []
PRO_INFO_TUP = namedtuple("Project_Info","name code contract builder\
 supervisor observor")
PRO_INFO = PRO_INFO_TUP("xxx工程","xx编号","xx合同","xx施工单位",\
	"xx监理单位","xx监测单位")
PRO_L.append(PRO_INFO)


IS_UPDATED = False
def is_project_updated():
	return IS_UPDATED


class MyPro(object):
	def __init__(self, parent_top):
		self.parent_top = parent_top
		self.pro_top = tk.Toplevel(parent_top)
		self.pro_top.title("工程信息")
		self.pro_top.geometry('680x320+400+280')
	
		ttk.Label(self.pro_top, text='').pack()

		#工程项目名称
		fm_name = tk.Frame(self.pro_top)
		ttk.Label(fm_name, text='工程项目: ').pack(side=tk.LEFT)
		self.v_name = tk.StringVar()
		ttk.Entry(fm_name, width=45, textvariable=self.v_name).pack()
		fm_name.pack()

		ttk.Label(self.pro_top, text='').pack()

		#工程信息
		fm_info = tk.Frame(self.pro_top)
		#单位
		fm_company = tk.Frame(fm_info)
		ttk.Label(fm_company, text='施工单位: ').grid(row=0, column=0)
		self.v_builder = tk.StringVar()
		ttk.Entry(fm_company, width=35, textvariable=self.v_builder).grid(row=0, column=1)

		ttk.Label(fm_company, text='监理单位: ').grid(row=1, column=0)
		self.v_supervisor = tk.StringVar()
		ttk.Entry(fm_company, width=35, textvariable=self.v_supervisor).grid(row=1, column=1)

		ttk.Label(fm_company, text='监测单位: ').grid(row=2, column=0)
		self.v_observor = tk.StringVar()
		ttk.Entry(fm_company, width=35, textvariable=self.v_observor).grid(row=2, column=1)
		fm_company.pack(side=tk.LEFT)

		ttk.Label(fm_info, width=2, text='').pack(side=tk.LEFT)

		#合同，编号
		fm_con = tk.Frame(fm_info)
		ttk.Label(fm_con, text='合同号: ').grid(row=0, column=0)
		self.v_contract = tk.StringVar()
		ttk.Entry(fm_con, width=35, textvariable=self.v_contract).grid(row=0, column=1)

		ttk.Label(fm_con, text='编号: ').grid(row=1, column=0)
		self.v_code = tk.StringVar()
		ttk.Entry(fm_con, width=35, textvariable=self.v_code).grid(row=1, column=1)

		ttk.Label(fm_con, text='').grid(row=2, column=0)
		ttk.Label(fm_con, text='').grid(row=2, column=1)
		fm_con.pack()
		fm_info.pack()

		ttk.Label(self.pro_top, text='').pack()

		#确认，退出按钮
		fm_button = tk.Frame(self.pro_top)
		ttk.Button(fm_button, text="确认", width=15, command=self.save_project).grid(\
			row=0, column=0)
		ttk.Label(fm_button, width=2, text='').grid(row=0, column=1)
		ttk.Button(fm_button, text="退出", width=15, command=self.discard_project).grid(\
			row=0, column=2)
		fm_button.pack()


	def save_project(self):
		self.update_project_info()
		self.pro_top.destroy()


	def discard_project(self):
		global IS_UPDATED
		IS_UPDATED = False
		self.pro_top.destroy()


	def update_project_info(self):
		global PRO_INFO
		global IS_UPDATED
		global PRO_L
		PRO_INFO = PRO_INFO._replace(name=self.v_name.get(), code=self.v_code.get(),\
			contract=self.v_contract.get(), builder=self.v_builder.get(),\
			supervisor=self.v_supervisor.get(), observor=self.v_observor.get())
		PRO_L[0] = PRO_INFO
		IS_UPDATED = True


def check_project_info():
	print(PRO_INFO)

if __name__ == '__main__':

	'''
	PRO_INFO = PRO_INFO_TUP("青岛市地铁1号线工程", "DSFJC02-RB", "M1-ZX-2016-222", \
	"中国中铁隧道局、十局集团有限公司", "北京铁城建设监理有限责任公司",\
	"中国铁路设计集团有限公司")
	'''
	print(PRO_INFO)

	top = tk.Tk()
	ttk.Button(top, text="Check", command=check_project_info).pack()
	my_pro = MyPro(top)
	top.mainloop()	


