#!/usr/bin/env python3

'''
工程项目信息汇总
新建, 打开, 保存
'''
import tkinter as tk
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror
from tkinter.messagebox import showinfo
from tkinter.messagebox import showwarning
import os
from tkinter.filedialog import askopenfilename
#from tkinter.filedialog import askdirectory

#拷贝文件
import shutil


#工程项目名, 编号， 施工单位， 监理单位， 监测单位, 区间
D = {"name":0,"area":1,"code":2,"contract":3,"builder":4,"supervisor":5,\
 "third_observer":6,"builder_observer":7, "xlsx_path":8,"date":9}
#注意main_gui会用到PRO_INFO，但是只有开始import，所以PRO_INFO一定不能赋值操作
#要保证PRO_INFO的id不变！
PRO_INFO = ["xxx工程","xx区间","xx编号","xx合同","xx施工单位","xx监理单位",\
"xx第三方监测单位","xx施工方监测单位","数据源文件地址","x年x月x日"]
PRO_INFO_BEFORE = PRO_INFO[:]

IS_UPDATED = False
def is_project_updated():
	print("DEBUG IS_UPDATED= ",IS_UPDATED)
	return IS_UPDATED

def set_project_updated_false():
	global IS_UPDATED
	IS_UPDATED = False
	
#工程文件目录
PRO_PATH = []

#工程项目备份
PRO_BAK_PATH = os.path.join(os.getcwd(), 'project_backup')
PRO_BAK_TXT = os.path.join(PRO_BAK_PATH, 'all_projects.txt')


class MyPro(object):
	def __init__(self, parent_top, file_path=None):
		print("__init__ MyPro")

		self.parent_top = parent_top
		self.pro_top = tk.Toplevel(parent_top)
		self.pro_top.title("工程信息")
		self.pro_top.geometry('680x320+400+280')
		#Always get focused
		#self.pro_top.grab_set()


		self.project_path = file_path
		self.project_bak_dir = None
		#保存.dr文件的默认文件夹
		#从new_proj文件时获取
		self.initial_dir = None

		#工程项目名称
		tk.Label(self.pro_top, text='').pack()
		fm_name = tk.Frame(self.pro_top)
		tk.Label(fm_name, text='* 项目工程: ').pack(side=tk.LEFT)
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
		tk.Label(fm_xlsx, text='* excel数据源: ').pack(side=tk.LEFT)
		self.v_xlsx_path = tk.StringVar()
		self.v_xlsx_path.set('汇总数据源.xlsx')
		tk.Entry(fm_xlsx, width=65, textvariable=self.v_xlsx_path)\
		.pack(side=tk.LEFT)
		tk.Button(fm_xlsx, text="...", width=5, command=self.select_xlsx)\
		.pack(side=tk.LEFT)
		fm_xlsx.pack()
		tk.Label(self.pro_top, text="注: '平面布点图'和'签名'文件夹需要和'数据源文件.xlsx'同一目录").\
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
		initial_path = None
		xlsx_path = self.v_xlsx_path.get()
		if os.path.exists(xlsx_path):
			initial_path = os.path.dirname(xlsx_path)
		else:
			initial_path = os.path.join(os.path.expanduser('~'), 'Desktop')
		xlsx_path = askopenfilename(filetypes=[("excel数据源文件","xlsx")],title="选择数据源",\
			initialdir=initial_path)
		#xlsx_path = askdirectory(title="选择数据源文件夹")
		print("DEBUG xlsx_path=",xlsx_path)
		if xlsx_path and os.path.exists(xlsx_path):
			xlsx_path = os.path.normpath(xlsx_path)
			self.v_xlsx_path.set(xlsx_path)
		else:
			pass

	#########select_xlsx()######################


	def confirm_project(self):
		'''
		保存确认按钮函数
		'''
		global PRO_PATH
		global PRO_BAK_PATH
		#如果有文件路径，说明是经过打开菜单进来的
		#认直接保存原来的这个文件
		if self.project_path:
			pass
		else:
			if not self.v_name.get():
				s = "项目名称不能为空!"
				self.popup_window(s)
				return
			if not self.v_xlsx_path.get():
				s = "excel数据源文件不能为空!"
				self.popup_window(s)
				return
			#如果文件路径是None,说明是新建菜单进来的
			#保存时，打开文件保存对话框，选择保存的文件
			project_name = self.v_name.get() + ".dr"
			desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')

			if self.initial_dir == None:
				self.initial_dir = desktop_dir
			self.project_path = asksaveasfilename(initialfile= project_name,\
				filetypes=[("监测日报项目文件","dr")], title="保存工程文件",\
				 initialdir=self.initial_dir)

			#asksavesasfilename Cancel:
			if not self.project_path:
				print("DEBUG, 没有保存项目文件")
				return
			print("DEBUG 项目文件保存为self.project_path:",self.project_path)

		PRO_PATH.append(self.project_path)

		self.update_project_info()
		self.save_project()

		#创建备份文件夹，以项目名称命名
		p_name = self.v_name.get()
		self.project_bak_dir = os.path.join(PRO_BAK_PATH, p_name)
		print("备份文件夹：",self.project_bak_dir)
		if not os.path.isdir(self.project_bak_dir):
			os.mkdir(self.project_bak_dir)

		self.bak_file(self.project_path)

		self.pro_top.destroy()

	#########confirm_project()#####################


	def discard_project(self):
		'''
		退出按钮函数
		'''
		global PRO_INFO
		global PRO_INFO_BEFORE
		global IS_UPDATED
		print("DEBUG discard_project点击了取消")
		IS_UPDATED = False

		#还原回去
		for i in range(len(PRO_INFO)):
			PRO_INFO[i] = PRO_INFO_BEFORE[i]
		print("还原PRO_INFO:",PRO_INFO)
		self.pro_top.destroy()

	##########discard_project()#####################


	def update_project_info(self):
		'''
		保存页面显示值到全局变量
		'''
		global PRO_INFO
		global IS_UPDATED
		global PRO_INFO_BEFORE
		#注意这里[:]用来防止id变更
		PRO_INFO[:] = [self.v_name.get(), self.v_area.get(), self.v_code.get(),\
		 self.v_contract.get(), self.v_builder.get(), self.v_supervisor.get(), \
		 self.v_third_observer.get(), self.v_builder_observer.get(),\
		  self.v_xlsx_path.get(), 'x年x月x日']
		IS_UPDATED = True
		#更新备份的PRO_INFO
		for i in range(len(PRO_INFO)):
			PRO_INFO_BEFORE[i] = PRO_INFO[i]
		print("DEBUG 同步PRO_INFO_BEFORE:",PRO_INFO_BEFORE)

	###########update_project_info()#################


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
	#############retrieve_project_info()#############


	def save_project(self):
		'''
		保存项目信息到本地硬盘文件
		'''
		global PRO_INFO
		global PRO_BAK_TXT

		with open(self.project_path, "wb") as fobj:
			for item in PRO_INFO:
				item = item + os.linesep
				item = item.encode('utf-8')
				fobj.write(item)
		print("DEBUG save success")

		#保存项目列表文件
		#读取项目文件
		with open(PRO_BAK_TXT, "rb") as fobj:
			all_projects_list = []
			while True:
				buff = fobj.readline().decode('utf-8').strip(os.linesep)
				if buff == '':
					break
				#这个怎么产生的?
				if buff == '\ufeff':
					continue
				else:
					#print("DEBUG buff=",buff)
					all_projects_list.append(buff)
		#print("DEBUG after read, all_projects_list:",all_projects_list)

		#将最新的项目文件移到最前面
		f_path = os.path.normpath(self.project_path)
		if f_path in all_projects_list:
			all_projects_list.remove(f_path)
		all_projects_list.insert(0,f_path)

		#print("DEBUG after insert, all_projects_list:",all_projects_list)

		with open(PRO_BAK_TXT, "wb") as fobj:
			#再写回去
			for item in all_projects_list:
				item = item + os.linesep
				item = item.encode('utf-8')
				fobj.write(item)
		print("更新项目列表文件成功")

	################save_project()##################


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
	###############load_project()#####################


	def bak_file(self,file_path):
		'''
		备份份文件到
		self.project_bak_dir里
		'''
		if os.path.isfile(file_path):
			if self.project_bak_dir == None:
				print("Error, self.project_bak_dir:NONE!")
				return False
			else:
				try:
					shutil.copy(file_path, self.project_bak_dir)
				except Exception as e:
					print("备份Error:",e)
					return False

		else:
			print("Error, file_path:{}不存在!".format(file_path))
			return Fasle

		return True
	################bak_file()########################	


	def popup_window(self, s, error= False):
		'''
		弹出信息通知窗口
		'''
		if not error:
			showinfo(message = s)
		else:
			showerror(message = s)

############class MyPro(object):



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


