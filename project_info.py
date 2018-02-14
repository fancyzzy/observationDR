#!/usr/bin/env python3

'''
工程项目信息汇总
新建, 打开, 保存
'''
import tkinter as tk
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror
from tkinter.messagebox import showwarning
import os
from tkinter.filedialog import askopenfilename


#工程项目名, 编号， 施工单位， 监理单位， 监测单位, 区间
D = {"name":0,"area":1,"code":2,"contract":3,"builder":4,"supervisor":5,\
 "third_observer":6,"builder_observer":7, "xlsx_path":8,"date":9}
PRO_INFO = ["xxx工程","xx区间","xx编号","xx合同","xx施工单位","xx监理单位",\
"xx第三方监测单位","xx施工方监测单位","数据源文件地址","x年x月x日"]

IS_UPDATED = False
def is_project_updated():
	return IS_UPDATED

#工程文件目录
PRO_PATH = []

class MyPro(object):
	def __init__(self, parent_top, file_path=None):
		print("__init__ MyPro")

		self.parent_top = parent_top
		self.pro_top = tk.Toplevel(parent_top)
		self.pro_top.title("工程信息")
		self.pro_top.geometry('680x320+400+280')
		#Always get focused
		self.pro_top.grab_set()

		self.project_path = file_path

		#工程项目名称
		tk.Label(self.pro_top, text='').pack()
		fm_name = tk.Frame(self.pro_top)
		tk.Label(fm_name, text='项目工程: ').pack(side=tk.LEFT)
		self.v_name = tk.StringVar()
		tk.Entry(fm_name, width=45, textvariable=self.v_name).pack()
		fm_name.pack()

		#工程区间
		fm_area = tk.Frame(self.pro_top)
		tk.Label(fm_area, text='项目区间: ').pack(side=tk.LEFT)
		self.v_area = tk.StringVar()
		tk.Entry(fm_area, width=35, textvariable=self.v_area).pack()
		fm_area.pack()

		tk.Label(self.pro_top, text='').pack()

		#其他工程信息
		fm_info = tk.Frame(self.pro_top)
		#单位
		fm_company = tk.Frame(fm_info)
		tk.Label(fm_company, text='施工单位: ').grid(row=0, column=0)
		self.v_builder = tk.StringVar()
		tk.Entry(fm_company, width=35, textvariable=self.v_builder)\
		.grid(row=0, column=1)

		tk.Label(fm_company, text='监理单位: ').grid(row=1, column=0)
		self.v_supervisor = tk.StringVar()
		tk.Entry(fm_company, width=35, textvariable=self.v_supervisor)\
		.grid(row=1, column=1)

		tk.Label(fm_company, text='第三方监测单位: ').grid(row=2, column=0)
		self.v_third_observer = tk.StringVar()
		tk.Entry(fm_company, width=35, textvariable=self.v_third_observer)\
		.grid(row=2, column=1)

		tk.Label(fm_company, text='施工方监测单位: ').grid(row=3, column=0)
		self.v_builder_observer = tk.StringVar()
		tk.Entry(fm_company, width=35, textvariable=self.v_builder_observer)\
		.grid(row=3, column=1)

		fm_company.pack(side=tk.LEFT)

		tk.Label(fm_info, width=2, text='').pack(side=tk.LEFT)

		#合同，编号
		fm_con = tk.Frame(fm_info)
		tk.Label(fm_con, text='合同号: ').grid(row=0, column=0)
		self.v_contract = tk.StringVar()
		tk.Entry(fm_con, width=35, textvariable=self.v_contract)\
		.grid(row=0, column=1)

		tk.Label(fm_con, text='编号: ').grid(row=1, column=0)
		self.v_code = tk.StringVar()
		tk.Entry(fm_con, width=35, textvariable=self.v_code)\
		.grid(row=1, column=1)

		tk.Label(fm_con, text='').grid(row=2, column=0)
		tk.Label(fm_con, text='').grid(row=2, column=1)
		fm_con.pack()
		fm_info.pack()

		tk.Label(self.pro_top, text='').pack()


		#xlsx数据源
		fm_xlsx = tk.Frame(self.pro_top)
		tk.Label(fm_xlsx, text='excel数据源: ').pack(side=tk.LEFT)
		self.v_xlsx_path = tk.StringVar()
		tk.Entry(fm_xlsx, width=65, textvariable=self.v_xlsx_path)\
		.pack(side=tk.LEFT)
		tk.Button(fm_xlsx, text="...", width=5, command=self.select_xlsx)\
		.pack(side=tk.LEFT)
		fm_xlsx.pack()
		tk.Label(self.pro_top, text='注:把平面布点图片文件放到excel数据源同目录下').\
		pack()

		tk.Label(self.pro_top, text='').pack()
		tk.Label(self.pro_top, text='').pack()

		#确认，退出按钮
		fm_button = tk.Frame(self.pro_top)
		tk.Button(fm_button, text="确认", width=15, command=self.confirm_project)\
		.grid(row=0, column=0)
		tk.Label(fm_button, width=2, text='').grid(row=0, column=1)
		tk.Button(fm_button, text="取消", width=15, command=self.discard_project)\
		.grid(row=0, column=2)
		fm_button.pack()

		#当加载的项目文件非空，更新页面项目信息为文件中的信息
		if self.project_path and os.path.exists(self.project_path):
			if self.load_project():
				self.retrieve_project_info()
			else:
				self.discard_project()
	#############__init__()#####################


	def select_xlsx(self):
		'''
		选择数据源文件
		'''
		print("select xlsx file")
		xlsx_path = askopenfilename(filetypes=[("excel数据源文件","xlsx")])
		if xlsx_path and os.path.exists(xlsx_path):
			xlsx_path = os.path.normpath(xlsx_path)
			self.v_xlsx_path.set(xlsx_path)
		else:
			pass


	def confirm_project(self):
		'''
		保存确认按钮函数
		'''
		global PRO_PATH
		#如果有文件路径，说明是经过打开菜单进来的
		#认直接保存原来的这个文件
		if self.project_path:
			pass
		else:
			if not self.v_name.get():
				return
			#如果文件路径是None,说明是新建菜单进来的
			#保存时，打开文件保存对话框，选择保存的文件
			project_name = self.v_name.get() + ".dr"
			#新建一个文件，用于监测项目工程文件
			self.project_path = asksaveasfilename(initialfile= project_name,\
				filetypes=[("监测日报项目文件","dr")])

			if self.project_path:
				#创建空文件
				with open(self.project_path, 'wb') as foj:
					pass
			else:
				return

		PRO_PATH.append(self.project_path)

		self.update_project_info()
		self.save_project()
		self.pro_top.destroy()


	def discard_project(self):
		'''
		退出按钮函数
		'''
		global IS_UPDATED
		IS_UPDATED = False
		self.pro_top.destroy()


	def update_project_info(self):
		'''
		保存页面显示值到全局变量
		'''
		global PRO_INFO
		global IS_UPDATED
		PRO_INFO[:] = [self.v_name.get(), self.v_area.get(), self.v_code.get(),\
		 self.v_contract.get(), self.v_builder.get(), self.v_supervisor.get(), \
		 self.v_third_observer.get(), self.v_builder_observer.get(),\
		  self.v_xlsx_path.get(), 'x年x月x日']
		IS_UPDATED = True


	def retrieve_project_info(self):
		'''
		刷新页面显示值
		'''
		self.v_name.set(PRO_INFO[D['name']])
		self.v_area.set(PRO_INFO[D['area']])
		self.v_code.set(PRO_INFO[D['code']])
		self.v_contract.set(PRO_INFO[D['contract']])
		self.v_builder.set(PRO_INFO[D['builder']])
		self.v_supervisor.set(PRO_INFO[D['supervisor']])
		self.v_third_observer.set(PRO_INFO[D['third_observer']])
		self.v_builder_observer.set(PRO_INFO[D['builder_observer']])
		self.v_xlsx_path.set(PRO_INFO[D['xlsx_path']])


	def save_project(self):
		'''
		保存项目信息到本地硬盘文件
		'''
		global PRO_INFO
		with open(self.project_path, "wb") as fobj:
			for item in PRO_INFO:
				item = item + os.linesep
				item = item.encode('utf-8')
				fobj.write(item)
		print("save success")


	def load_project(self):
		'''
		从硬盘读取项目文件
		'''
		global PRO_INFO
		print("DEBUG self.project_path =", self.project_path)
		with open(self.project_path, 'rb') as fobj:
			lines = fobj.readlines()
			ln = len(lines)
			max_ln = len(PRO_INFO)
			if max_ln != ln:
				print("Warnning, mismatch")
				print("{}, ln= {}, max_ln={}".format(self.project_path, ln, max_ln))
				showerror(title="项目文件错误", message="文件{}, 行数={}, 应该={}\
					".format(self.project_path, ln, max_ln))
				return False
			if ln == max_ln:
				for i in range(max_ln):
					ss = lines[i].decode('utf-8')
					PRO_INFO[i] = ss.strip().strip(os.linesep)

				print("load success, PRO_INFO=",PRO_INFO)
				#if '' in PRO_INFO:
				#	showwarning(title="警告", message="项目信息有缺失")
				return True


def check_project_info():
	print(PRO_INFO)

if __name__ == '__main__':

	PRO_INFO = ["青岛市地铁1号线工程", "一二工区", "DSFJC02-RB", "M1-ZX-2016-222", \
	"中国中铁隧道局、十局集团有限公司", "北京铁城建设监理有限责任公司",\
	"中国铁路设计集团有限公司"]
	print(PRO_INFO)

	top = tk.Tk()
	tk.Button(top, text="Check", command=check_project_info).pack()
	my_pro = MyPro(top)
	top.mainloop()	


