#!/usr/bin/env python3

'''
工程监测日报生成系统
用于生成固定格式的每日日报
python3.6.1
author: Felix
email:fancyzzy@163.com
'''
import tkinter as tk

from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo

import os
from project_info import *
import gen_docx
import read_xlsx
from datetime import datetime

my_color_office_blue ='#%02x%02x%02x' % (43,87,154)
my_color_orange ='#%02x%02x%02x' % (192,121,57)
my_color_light_orange = '#%02x%02x%02x' % (243,183,95)

logo_name = 'pic\pen.png'
logo_path = os.path.join(os.getcwd(),logo_name)

class MyTop(object):
	def __init__(self):
		self.top = tk.Tk()
		self.top.title("监测日报")
		self.top.geometry('750x520+400+280')

		#每次窗口获得焦点，更新标题
		self.top.bind("<FocusIn>", self.enter_top)

		#工程文件路径
		self.f_path = None
		#xlsx类实例
		self.my_xlsx = None

		#菜单
		self.menu_bar = tk.Menu(self.top)
		self.top.config(menu=self.menu_bar)
		file_menu = tk.Menu(self.menu_bar, tearoff=0)
		proj_menu = tk.Menu(self.menu_bar, tearoff=0)
		self.menu_bar.add_cascade(label="文件", menu=file_menu)
		self.menu_bar.add_cascade(label="工程", menu=proj_menu)

		file_menu.add_command(label="新建工程", command=self.new_project)
		file_menu.add_command(label="打开工程", command=self.open_project)
		#file_menu.add_separator()
		proj_menu.add_command(label="更改工程信息", command=\
			self.display_update_project)
		self.menu_bar.entryconfig("工程", state="disable")

		for i in range(3):
			tk.Label(self.top, text='').pack()

		#初始标题
		self.fm_init = tk.Frame(self.top)
		#插图logo
		tl = tk.Label(self.fm_init, compound='top')
		print("DEBUG os.getcwd()",os.getcwd())
		mg = tk.PhotoImage(file=logo_path)
		tl.lenna_image_png = mg
		tl['image'] = mg
		tl.pack()

		for i in range(1):
			tk.Label(self.fm_init, text='').pack()

		label_init = tk.Label(self.fm_init, text='监测日报助手1.0', \
			font = ('楷体', 32, 'bold'), fg= my_color_office_blue)
		label_init.pack()



		self.fm_init.pack()

		#新工程
		self.fm_pro = tk.Frame(self.top)
		#工程项目名称, 区间
		fm_title = tk.Frame(self.fm_pro)
		self.label_title = tk.Label(fm_title, text='XX工程监测日报',\
			font = ('楷体', 28, 'bold'), fg= my_color_orange)
		self.label_title.pack()
		self.label_area = tk.Label(fm_title, text='XX区间',\
			font = ('楷体', 28, 'bold'), fg= my_color_orange)
		self.label_area.pack()
		fm_title.pack()

		for i in range(3):
			tk.Label(self.fm_pro, text='').pack()

		#No 编号
		fm_no = tk.Frame(self.fm_pro, width=400)
		tk.Label(fm_no, text='期号：  第 ', font=('楷体', 18, 'bold')).\
		pack(side=tk.LEFT)
		self.v_no = tk.StringVar()
		large_font = ('楷体', 24, 'normal')
		self.entry_no = tk.Entry(fm_no, width=5, font=large_font, \
			relief='flat',textvariable=self.v_no)
		self.entry_no.pack(side=tk.LEFT)
		tk.Label(fm_no, text=' 期', font=('楷体', 18, 'bold')).\
		pack(side=tk.LEFT)
		tk.Label(fm_no, text=' '*10).pack(side=tk.LEFT)
		fm_no.pack()

		tk.Label(self.fm_pro, text='').pack()

		#date 日期
		fm_date = tk.Frame(self.fm_pro, width=400)
		tk.Label(fm_date, text=' '*9+'测量日期： ',font = ('楷体', 18, 'bold')).\
		pack(side=tk.LEFT)
		self.v_date = tk.StringVar()
		self.entry_date = tk.Entry(fm_date, width=10, font=large_font,\
			relief='flat', textvariable=self.v_date)
		self.entry_date.pack(side=tk.LEFT)
		tk.Label(fm_date, text='(年/月/日)',font=('楷体', 14)).\
		pack(side=tk.LEFT)
		tk.Label(fm_date, text=' ', font=('楷体', 12)).\
		pack(side=tk.LEFT)
		

		fm_date.pack()

		for i in range(2):
			tk.Label(self.fm_pro, text='').pack()

		#生成日报按钮
		tk.Label(self.fm_pro, text='').pack()
		fm_button = tk.Frame(self.fm_pro)
		#self.button_gen =tk.Button(fm_button, text="生成日报", \
		#	command=self.gen_report)
		self.button_gen =tk.Button(fm_button, text="生成日报", font=('楷体', 24, 'bold'),\
			width=10, height=1, bg=my_color_light_orange, command=self.gen_report)
		self.button_gen.pack()
		fm_button.pack()
		#初始化不显示工程标题
		#self.fm_pro.pack()
	########__init__()#################################################


	def enter_top(self,event):  
		'''
		当焦点在主界面时，根据工程是否存在，刷新主界面的显示内容
		'''
		global PRO_PATH
		if event.widget == self.top:
			print("Main GUI get the focus")
			self.update_title()
			if len(PRO_PATH) > 0:
				self.f_path = PRO_PATH[-1]
	#########enter_top()###############################################


	def new_project(self):
		'''
		新建空的工程文件
		'''
		print("new project")
		#None 表示新建文件工程
		my_pro = MyPro(self.top, None)
	#########new_project()###############################################


	def open_project(self):
		'''
		选择工程dr文件，打开并且显示工程信息
		'''
		print("Opened project")
		self.f_path = askopenfilename(filetypes=[("监测日报项目文件","dr")])
		print("DEBUG self.f_path = ",self.f_path)
		if self.f_path and os.path.exists(self.f_path):
			self.f_path = os.path.normpath(self.f_path)
			my_pro = MyPro(self.top, self.f_path)
		else:
			pass
	##################open_project()#####################################


	def display_update_project(self):
		'''
		更改工程信息
		'''
		if self.f_path and os.path.exists(self.f_path):
			my_pro = MyPro(self.top, self.f_path)
		else:
			pass
	#############display_update_project()#################################


	def update_title(self):
		global PRO_INFO
		global IS_UPDATED
		if is_project_updated():
			self.label_title.config(text=PRO_INFO[D['name']])
			self.label_area.config(text=PRO_INFO[D['area']])
			self.fm_init.pack_forget()
			self.fm_pro.pack()
			self.menu_bar.entryconfig("工程", state="normal")
	############update_title()####################################		


	def load_xlsx(self):
		'''
		读取解析xlsx数据库
		'''
		global PRO_INFO
		global D
		print("start to load xlsx database")
		self.my_xlsx = read_xlsx.MyXlsx(PRO_INFO[D['xlsx_path']])

		print("load finished")
		return True
	#######load_xlsx()##############################################


	def gen_report(self):
		'''
		生成日报
		'''
		global D
		global PRO_INFO
		print("Generate report start")

		#更新code = code+期号
		project_info = PRO_INFO[:]
		if self.v_no.get():
			project_info[D['code']] += '-%s'%(self.v_no.get())
		
		#更新日期
		s = self.v_date.get()
		s.replace(r'\\', '/')
		s.replace(r'.', '/')
		datetime_value = datetime.strptime(s, '%Y/%m/%d')
		project_info[D['date']] = datetime_value
		#检查日期是否合法	
		#pass

		#日报文件名
		docx_name = project_info[D['name']] + '监测日报' +\
		 	s.replace('/','.') + '.docx'

		#获取xlsx数据源
		if not self.my_xlsx:
			if not self.load_xlsx():
				print("load xlsx failed")
				return False
		else:
			pass

		#创建日报docx文件, 默认存放日报文件地址和项目文件一个文件夹
		docx_path = os.path.join(os.path.dirname(self.f_path), docx_name)
		with open(docx_path, 'wb') as fobj:
			pass

		#生成日报
		my_docx = gen_docx.MyDocx(docx_path, project_info, self.my_xlsx)
		if my_docx.gen_docx():
			s = "成功生成日报文件!\n %s" %docx_path
			print(s)
			self.popup_window(s)
		else:
			s = "日报文件生成失败!"
			print(s)
			self.popup_window(s)

		return True
	########gen_report#########################################################


	def popup_window(self, s):
		'''
		弹出信息通知窗口
		'''
		showinfo(message = s)


#class MyTop(object) end


if __name__ == '__main__':

	my_top = MyTop()
	my_top.top.mainloop()




