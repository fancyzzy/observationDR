#!/usr/bin/env python3

'''
工程监测日报生成系统
用于生成固定格式的每日日报
python3.6.1
author: Felix
email:fancyzzy@163.com
'''
print("this is main")
import tkinter as tk
from tkinter.ttk import Progressbar,Style

from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
from tkinter.messagebox import showerror

import os
from project_info import *

from datetime import datetime
import threading
import queue
from my_log import printl
from my_log import QUE
from my_log import SENTINEL
from my_log import LOG_PATH

L_THREADS = []

my_color_office_blue ='#%02x%02x%02x' % (43,87,154)
my_color_orange ='#%02x%02x%02x' % (192,121,57)
my_color_light_orange = '#%02x%02x%02x' % (243,183,95)
sunken_grey = '#%02x%02x%02x' % (240,240,240)

logo_path = os.path.join(os.getcwd(),'pic\pen.png')
icon_path = os.path.join(os.getcwd(),'pic\pen.ico')

class MyTop(object):
	def __init__(self):
		self.top = tk.Tk()
		self.top.title("监测日报")
		self.top.geometry('750x520+400+280')
		self.top.iconbitmap(icon_path)
		#每次窗口获得焦点，更新标题
		self.top.bind("<FocusIn>", self.enter_top)

		#工程文件路径
		self.f_path = None
		#xlsx类实例
		self.my_xlsx = None
		self.is_generating = False

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
		self.fm_pro = tk.Frame(self.top, width=750, height=520)
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
		self.entry_no.focus_set()
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
		self.entry_date.bind('<Return>',self.gen_report)
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
		self.button_gen =tk.Button(fm_button, text="生成日报", font=('楷体', 24, 'bold'),\
			width=10, height=1, bg=my_color_light_orange, command=self.gen_report)
		self.button_gen.bind('<Return>',self.gen_report)
		self.button_gen.pack()
		fm_button.pack()


		#进度条窗口
		self.prog = ProgBar(self.top)

		'''
		#status bar
		self.fm_status = tk.Frame(self.top)
		for i in range(1):
			tk.Label(self.fm_status, text='').grid(row=i,column=0)

		self.v_status = tk.StringVar()
		self.v_status.set(''.join(list(map(str,[i for i in range(60)]))))
		self.label_status = tk.Label(self.fm_status,textvariable=self.v_status, bd=1,\
		 relief='sunken',justify='left')
		#self.label_status.pack(fill=tk.X)
		self.label_status.grid(row=2,column=0)
		#self.fm_status.pack(side=tk.LEFT)
		'''

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
		print("new project done")

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
			#self.fm_status.pack(side=tk.LEFT)
			#self.prog.show_status(False)
			if not self.is_generating:
				self.menu_bar.entryconfig("工程", state="normal")

			#重置excel数据源
			self.my_xlsx = None
	############update_title()####################################		


	def load_xlsx(self):
		'''
		读取解析xlsx数据库
		'''
		import read_xlsx
		global PRO_INFO
		global D
		print("start to load xlsx database")
		try:
			self.my_xlsx = read_xlsx.MyXlsx(PRO_INFO[D['xlsx_path']])
		except Exception as e:
			print("Error! 加载excel数据源错误:{}".format(e))
			self.popup_window(e,error=True)
			return False

		print("load finished")
		return True
	#######load_xlsx()##############################################


	def gen_report(self, event=None):
		'''
		生成日报
		'''

		global D
		global PRO_INFO
		global L_THREADS
		global QUE
		global LOG_PATH
		outqueue = QUE

		print("Generate report start")
		#更新code = code+期号
		project_info = PRO_INFO[:]
		if self.v_no.get():
			project_info[D['code']] += '-%s'%(self.v_no.get())
		#更新日期
		s = self.v_date.get()
		s.replace(r'\\', '/')
		s.replace(r'.', '/')
		#创建日报docx文件, 默认存放日报文件地址和项目文件一个文件夹
		docx_name = s.replace('/','.') + '监测日报' + '.docx'
		docx_path = os.path.join(os.path.dirname(self.f_path), docx_name)
		LOG_PATH[0] = os.path.join(os.path.dirname(self.f_path), 'my_log.txt')
		print("DEBUG LOG_PATH=",LOG_PATH)

		try:
			#检查日期是否合法	
			datetime_value = datetime.strptime(s, '%Y/%m/%d')
			project_info[D['date']] = datetime_value
		except Exception as e:
			err_s = "请输入合法日期，比如:2018/1/1"
			printl(err_s,False)
			self.popup_window(err_s)
			return False

		#清空log 消息队列
		outqueue.queue.clear()
		self.prog.p_bar["value"] = 0
		#显示进度条
		self.prog.show_status(True)

		#启动生成日报线程，防止主界面freeze
		t = threading.Thread(target=self.run_gen_report,args=(docx_path,\
			project_info))
		L_THREADS.append(t)
		t.start()	

		#更新主GUI
		self.top.after(250, self.update)

		return True
	########gen_report#########################################################


	def run_gen_report(self, docx_path, project_info):
		'''
		线程回调函数
		使用线程防止主界面freeze
		'''
		print("run_gen_report start")
		global QUE
		global SENTINEL
		outqueue = QUE

		print("生成日报ing...")
		self.button_gen.config(bg=sunken_grey,relief='sunken',state='disabled')
		self.menu_bar.entryconfig("文件", state="disable")
		self.menu_bar.entryconfig("工程", state="disable")
		self.is_generating = True

		#获取xlsx数据源
		#12% percent
		if not self.my_xlsx:
			outqueue.put('loading xlsx...')
			if not self.load_xlsx():
				print("12@load xlsx failed")
				self.button_gen.config(bg=my_color_light_orange,relief='raised',\
					state='normal')
				self.menu_bar.entryconfig("文件", state="normal")
				self.menu_bar.entryconfig("工程", state="normal")
				self.is_generating = False
				return False
			else:
				outqueue.put('loading finished')
				outqueue.put('12@')
		else:
			outqueue.put('12@load xlsx finished')
			pass

		#生成日报
		result = None
		try:
			#延迟加载
			import gen_docx
			my_docx = gen_docx.MyDocx(docx_path, project_info, self.my_xlsx)
			result = my_docx.gen_docx()
		except Exception as e:
			s = "Error, 生成日报错误:{}".format(e)
			printl(s,False)
			self.popup_window(s, True)

		#debug percentage not 100%
		self.prog.p_bar["value"]=100.

		if result:
			s = "生成日报文件成功!\n %s" %docx_path
			print(s)
			self.popup_window(s)
		else:
			s = "日报文件生成失败!"
			print(s)
			self.popup_window(s)

		self.button_gen.config(bg=my_color_light_orange,relief='raised',\
					state='normal')
		self.menu_bar.entryconfig("文件", state="normal")
		self.menu_bar.entryconfig("工程", state="normal")
		self.is_generating = False

		if result:
			printl("日报文件存储于: %s\n"%(docx_path))
		else:
			printl("日报生成遇到问题\n")
		#send the finish flag
		outqueue.put(SENTINEL)
		print("日报线程结束")
	##########fun_gen_report()################################################

	def update(self):
		global QUE
		global SENTINEL
		outqueue = QUE
		try:
			msg = outqueue.get_nowait()
			if msg is not SENTINEL:
				#处理progress log
				v = 0
				s = ''
				if '@' in msg:
					v,_= msg.split('@')
					#更新进度条
					if self.prog.p_bar["value"] < 100:
						self.prog.p_bar["value"] += float(v)

				else:
					s = msg
					self.prog.update_log(s)
				self.top.after(250, self.update)

			else:
				s = "收到sentinel"
				print(s)
				#self.prog.update_log(s)
				print("self.prog.p_bar=",self.prog.p_bar["value"])
				
		except queue.Empty:
			self.top.after(250, self.update)

		self.prog.style_bar.configure("LabeledProgressbar",\
					 text="{}%      ".format(round(self.prog.p_bar["value"]),1))
	#############update()####################################################


	def popup_window(self, s, error= False):
		'''
		弹出信息通知窗口
		'''
		if not error:
			showinfo(message = s)
		else:
			showerror(message = s)

#class MyTop(object) end


class ProgBar(object):
	def __init__(self,top):
		print("this is Progress bar status")
		#进度条窗口
		#status bar
		self.fm_status = tk.Frame(top)
		for i in range(1):
			tk.Label(self.fm_status, text='').grid(row=i,column=0)

		self.v_status = tk.StringVar()
		self.v_status.set('Start')
		self.label_status = tk.Label(self.fm_status,textvariable=self.v_status, bd=1,\
		 justify='left')
		self.label_status.grid(row=3,column=1,sticky=tk.S)

		#进度条
		self.style_bar = Style(top)
		self.style_bar.layout("LabeledProgressbar",
         [('LabeledProgressbar.trough',
           {'children': [('LabeledProgressbar.pbar',
                          {'side': 'left', 'sticky': 'ns'}),
                         ("LabeledProgressbar.label",
                          {"sticky": ""})],
           'sticky': 'nswe'})])

		self.p_bar = Progressbar(self.fm_status, orient=tk.HORIZONTAL,\
		 length=100, mode='determinate',style="LabeledProgressbar")
		self.style_bar.configure("LabeledProgressbar", text="0 %      ")
		self.p_bar.grid(row=3,column=0)
		self.p_bar["maximum"]=100.
		self.p_bar["value"] = 0

	def update_log(self,s):
		self.v_status.set(s)

	def show_status(self, flag=True):
		if flag:
			self.fm_status.pack(side=tk.LEFT)
		else:
			self.fm_status.pack_forget()


##############class ProgWind##################################################


if __name__ == '__main__':

	print("Main start")
	my_top = MyTop()
	my_top.top.mainloop()
	print("Main end")




