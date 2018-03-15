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
from time import clock
import my_bak

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
		#工程信息实例
		self.my_proj = None

		#菜单
		self.menu_bar = tk.Menu(self.top)
		self.top.config(menu=self.menu_bar)
		self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
		self.proj_menu = tk.Menu(self.menu_bar, tearoff=0)
		self.menu_bar.add_cascade(label="文件", menu=self.file_menu)
		self.menu_bar.add_cascade(label="工程列表", menu=self.proj_menu)

		self.file_menu.add_command(label="新建工程", command=self.new_project)
		self.file_menu.add_command(label="打开工程", command=self.open_project)
		#self.file_menu.add_separator()
		self.file_menu.add_command(label="更改工程", command=self.update_project,\
			state='disable')
		#更新级联菜单项状态
		#self.file_menu.entryconfig("更改工程",state="normal")

		#打开工程列表文件
		global PRO_BAK_TXT
		p_list=[]
		n = 0
		with open(PRO_BAK_TXT, "rb") as fobj:
			while True:
				n += 1
				buff = fobj.readline().decode('utf-8').strip(os.linesep)
				if buff == '':
					break
				else:
					p_list.append(buff)
				if n >= 15:
					break
		#添加工程列表内容
		self.p_name=tk.StringVar()
		for item in p_list:
			self.proj_menu.add_radiobutton(label=item, variable=self.p_name,\
			 command=self.display_project)

		for i in range(3):
			tk.Label(self.top, text='').pack()

		#初始标题
		self.fm_init = tk.Frame(self.top)
		#插图logo
		tl = tk.Label(self.fm_init, compound='top')
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
		self.my_proj = MyPro(self.top, None)
		print("new project done")

	#########new_project()###############################################


	def open_project(self):
		'''
		选择工程dr文件，打开并且显示工程信息
		'''
		print("Opened project")
		self.f_path = askopenfilename(filetypes=[("监测日报项目文件","dr")])
		if self.f_path and os.path.exists(self.f_path):
			self.f_path = os.path.normpath(self.f_path)
			self.my_proj = MyPro(self.top, self.f_path)
		else:
			pass
	##################open_project()#####################################


	def update_project(self):
		'''
		更改工程信息
		'''
		if self.f_path and os.path.exists(self.f_path):
			self.my_proj = MyPro(self.top, self.f_path)
		else:
			pass
	#############update_project()#################################


	def display_project(self):
		'''
		根据选择的工程文件，显示工程
		'''

		project_path = self.p_name.get()
		print("DEBUG display_project:",project_path)
		if project_path and os.path.exists(project_path):
			self.my_proj = MyPro(self.top, project_path)
		else:
			s = ("没有找到项目文件:{}\n".format(project_path))
			showinfo(message = s)
			#从备份项目列表文件中删除

	###############display_project()##############################


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
				#更新菜单项可用
				self.file_menu.entryconfig("更改工程",state="normal")

			#重置excel数据源
			self.my_xlsx = None
	############update_title()####################################		


	def bak_dir_files(self):
		'''
		保存所有工程文件
		签名，布点图等
		'''

		if self.my_proj == None:
			return

		if self.f_path == None:
			return 

		if self.my_proj.project_bak_dir == None:
			return

		bak_path = self.my_proj.project_bak_dir
		dst_path = os.path.dirname(self.f_path)

		layout_bak_path = os.path.join(bak_path,'平面布点图')			
		sig_bak_path = os.path.join(bak_path,'签名')

		layout_path = os.path.join(dst_path, '平面布点图')
		sig_path = os.path.join(dst_path, '签名')

		my_bak.bak_directory(layout_path, layout_bak_path)
		my_bak.bak_directory(sig_path, sig_bak_path)
		return


	###########bak_all_files()###################################

	def load_xlsx(self):
		'''
		读取解析xlsx数据库
		'''
		import read_xlsx
		global PRO_INFO
		global D
		print("start to load xlsx database")
		xlsx_data_path = PRO_INFO[D['xlsx_path']]
		try:
			self.my_xlsx = read_xlsx.MyXlsx(xlsx_data_path)
		except Exception as e:
			print("Error! 加载excel数据源错误:{}".format(e))
			self.popup_window(e,error=True)
			return False

		#备份
		if self.my_proj.bak_file(xlsx_data_path):
			printl("备份数据源成功")
		else:
			printl("备份数据源失败")

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
		#feature, 增加时间后缀，避免命名重复
		s_now = datetime.now().strftime("%Y%m%d%H%M%S")
		docx_name = s.replace('/','.') + '监测日报' + ("_%s.docx"%s_now)

		docx_path = os.path.join(os.path.dirname(self.f_path), docx_name)
		#my_log.txt写在备份区
		#LOG_PATH[0] = os.path.join(os.path.dirname(self.f_path), 'my_log.txt')
		LOG_PATH[0] = os.path.join(self.my_proj.project_bak_dir, 'my_log.txt')
		print("DEBUG LOG_PATH=",LOG_PATH)
		#fron now on, log can be recorded at the right bakup place
		printl("\n")
		printl("开始生成日报，时间:{}".format(str(datetime.now())))
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

		start = clock()

		print("生成日报ing...")
		self.button_gen.config(bg=sunken_grey,relief='sunken',state='disabled')
		self.menu_bar.entryconfig("文件", state="disable")
		self.menu_bar.entryconfig("工程列表", state="disable")
		self.is_generating = True

		#获取xlsx数据源
		#12% percent
		if not self.my_xlsx:
			outqueue.put('加载数据源...')
			if not self.load_xlsx():
				print("12@加载数据源失败!")
				self.button_gen.config(bg=my_color_light_orange,relief='raised',\
					state='normal')
				self.menu_bar.entryconfig("文件", state="normal")
				self.menu_bar.entryconfig("工程列表", state="normal")
				self.is_generating = False
				return False
			else:
				outqueue.put('加载成功')
				outqueue.put('12@')
		else:
			outqueue.put('12@数据源已加载')
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
		table_num = my_docx.get_table_num()

		end = clock()
		interval = end - start
		s_interval = ''
		s_min = interval/60
		s_sec = interval%60.00
		s_interval = "%d分%.2f秒"%(s_min,s_sec)

		if result:
			s = "生成日报文件成功!\n表格: %d个\n用时: %s\n %s"%(table_num,s_interval,docx_path)
			print(s)
			self.popup_window(s)
		else:
			s = "日报文件生成失败!"
			print(s)
			self.popup_window(s)

		self.button_gen.config(bg=my_color_light_orange,relief='raised',\
					state='normal')
		self.menu_bar.entryconfig("文件", state="normal")
		self.menu_bar.entryconfig("工程列表", state="normal")
		self.is_generating = False

		if result:
			#备份
			if self.my_proj.bak_file(docx_path):
				printl("备份日志文件成功")
			else:
				printl("备份日志文件失败")

			#备份签名和平面布点图文件夹:	
			print("DEbug,开始备份文件夹\n")
			self.bak_dir_files()
			printl("日报文件存储于: %s\n"%(docx_path))
		else:
			printl("日报生成遇到问题\n")
		#send the finish flag
		outqueue.put(SENTINEL)
		print("日报线程结束")
	##########run_gen_report()################################################

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




