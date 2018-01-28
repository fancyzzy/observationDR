#!/usr/bin/env python3

'''
工程监测日报生成系统
用于生成固定格式的每日日报
python3.6.1
author: Felix
email:fancyzzy@163.com
'''
import tkinter as tk
from tkinter import ttk
from project_info import *

class MyTop(object):
	def __init__(self):
		self.top = tk.Tk()
		self.top.title("监测日报")
		self.top.geometry('680x520+400+280')

		#每次窗口获得焦点，更新标题
		self.top.bind("<FocusIn>", self.enter_top)

		#菜单
		self.menu_bar = tk.Menu(self.top)
		self.top.config(menu=self.menu_bar)
		project_menu = tk.Menu(self.menu_bar, tearoff=0)
		project_menu.add_command(label="新建", command=self.new_project)
		project_menu.add_separator()
		project_menu.add_command(label="打开", command=self.open_project)
		self.menu_bar.add_cascade(label="工程", menu=project_menu)

		ttk.Label(self.top, text='').pack()

		#初始标题
		#frame.bind("<Button-1>", callback)
		self.fm_init = tk.Frame(self.top)
		label_init = ttk.Label(self.fm_init, text='工程监测日报自动生成系统')
		label_init.pack()
		self.fm_init.pack()

		#新工程
		self.fm_pro = tk.Frame(self.top)
		#标题
		fm_title = tk.Frame(self.fm_pro)
		self.label_title = ttk.Label(fm_title, text='XX工程监测日报')
		self.label_title.pack()
		fm_title.pack()

		#No 编号
		ttk.Label(self.fm_pro, text='').pack()
		fm_no = tk.Frame(self.fm_pro)
		ttk.Label(fm_no, text='期号: 第').grid(row=0, column=0)
		self.v_no = tk.StringVar()
		self.entry_no = ttk.Entry(fm_no, width=12, textvariable=self.v_no)
		self.entry_no.grid(row=0, column=1)
		ttk.Label(fm_no, text='期').grid(row=0,column=2)
		fm_no.pack()

		#date 日期
		fm_date = tk.Frame(self.fm_pro)
		ttk.Label(fm_date, text='测量日期: ').grid(row=1,column=0)
		self.v_date = tk.StringVar()
		self.entry_date = ttk.Entry(fm_date, width=20, textvariable=self.v_date)
		self.entry_date.grid(row=1,column=1)
		ttk.Label(fm_date, text='年/月/日').grid(row=1,column=2)
		fm_date.pack()

		#生成日报按钮
		ttk.Label(self.fm_pro, text='').pack()
		fm_button = tk.Frame(self.fm_pro)
		self.button_gen =ttk.Button(fm_button, text="生成日报", command=self.gen_report)
		self.button_gen.pack()
		fm_button.pack()
		#初始化不显示工程标题
		#self.fm_pro.pack()

	########__init__()################

	def enter_top(self,event):  
		if event.widget == self.top:
			print("I have gained the focus")
			self.update_title()


	def new_project(self):
		my_pro = MyPro(self.top)
		print("new project")

	######new_pro()#########

	def open_project(self):
		print("Opened project")
	######open_pro()#########

	def update_title(self):
		global PRO_L
		if is_project_updated():
			self.label_title.config(text=PRO_L[0].name)
			self.fm_init.pack_forget()
			self.fm_pro.pack()
	#######show_title##########

	def gen_report(self):
		global PRO_INFO
		global PRO_L
		print("Generate Daily Report")
		print(PRO_L)
	########gen_report########




if __name__ == '__main__':

	my_top = MyTop()
	my_top.top.mainloop()




