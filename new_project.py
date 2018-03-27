#!/usr/bin/env python3

'''
新建项目工程文件夹，
并且把样版数据源文件一并拷贝进去
'''
import tkinter as tk
from tkinter.messagebox import showerror
from tkinter.messagebox import showinfo
#from tkinter.filedialog import asksaveasfilename
#from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory

#拷贝文件
import my_bak
import project_info

#获取桌面
import os
def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), 'Desktop')


class NewProj(object):
	def __init__(self, parent_top):
		#print("__init__ NewProj")
		self.parent_top = parent_top
		self.pro_top = tk.Toplevel(parent_top)
		self.pro_top.title("新建工程")
		self.pro_top.geometry('580x200+400+280')
		#Always get focused
		self.pro_top.grab_set()

		#以项目名称命名的文件夹路径
		self.proj_dir_path = None

		#project info instance
		#file_path=None 表示项目工程.dr文件为none，意为新建文件工程
		self.my_proj = project_info.MyPro(parent_top, file_path=None)
		#隐藏
		self.my_proj.pro_top.withdraw()

		#工程项目名称
		tk.Label(self.pro_top, text='').pack()
		tk.Label(self.pro_top, text='').pack()
		fm_name = tk.Frame(self.pro_top)
		tk.Label(fm_name, text='* 工程名称: ').pack(side=tk.LEFT)
		self.v_proj_name = tk.StringVar()
		proj_name_entry = tk.Entry(fm_name, width=62, textvariable=self.v_proj_name)
		proj_name_entry.focus_set()
		proj_name_entry.pack()
		fm_name.pack()

		tk.Label(self.pro_top, text='').pack()

		#项目文件夹地址
		fm_folder = tk.Frame(self.pro_top)
		tk.Label(fm_folder, text='         * 位置: ').pack(side=tk.LEFT)
		self.v_proj_dir = tk.StringVar()
		#初始化位置就是桌面
		desktop_path = get_desktop_path()
		self.v_proj_dir.set(desktop_path + os.path.sep)
		tk.Entry(fm_folder, width=55, textvariable=self.v_proj_dir)\
		.pack(side=tk.LEFT)
		tk.Button(fm_folder, text="...", width=5, command=self.select_folder)\
		.pack(side=tk.LEFT)
		fm_folder.pack()
		tk.Label(self.pro_top, text="注: 会在该位置创建以工程名称命名的文件夹").pack()

		tk.Label(self.pro_top, text='').pack()

		#确认，退出按钮
		fm_button = tk.Frame(self.pro_top)
		confirm_button = tk.Button(fm_button, text="下一步", width=15, command=self.confirm_project)
		confirm_button.grid(row=0, column=0)
		confirm_button.bind("<Return>", self.confirm_project)
		tk.Label(fm_button, width=2, text='').grid(row=0, column=1)
		tk.Button(fm_button, text="取消", width=15, command=self.discard_project)\
		.grid(row=0, column=2)
		fm_button.pack()

	#############__init__()#####################

	def select_folder(self):
		'''
		选择文件夹
		'''
		print("select folder")
		#folder = askopenfilename(filetypes=[("excel数据源文件","xlsx")],title="选择数据源")
		folder = askdirectory(title="请选择保存位置",initialdir=get_desktop_path())
		print("DEBUG folder=",folder)
		if folder:
			self.v_proj_dir.set(os.path.normpath(folder)+os.path.sep)
		else:
			pass
	#########select_xlsx()######################


	def confirm_project(self,event=None):
		'''
		下一步按钮函数
		'''
		self.proj_dir_path = self.v_proj_dir.get() + self.v_proj_name.get()
		print("DEBUG self.proj_dir_path=",self.proj_dir_path)
		if not self.v_proj_name.get():
			s = "项目名称不能为空!"
			showerror(message=s)
			return
		if not self.proj_dir_path:
			s = "项目文件夹位置不能为空!"
			showerror(message=s)
			return

		#创建文件夹，并且拷贝样版数据源
		if not os.path.exists(self.proj_dir_path):

			print("DEBUG 项目文件地址保存为:",self.proj_dir_path)
			os.mkdir(self.proj_dir_path)
			if self.copy_data_template():
				showinfo(message ="创建文件夹成功!\n{}".format(self.proj_dir_path))
				self.pro_top.destroy()

				#显示project_info实例,并且配置工程项目名称，和数据源
				self.my_proj.v_name.set(self.v_proj_name.get())
				data_source = os.path.join(self.proj_dir_path,'数据源\汇总数据源.xlsx')
				self.my_proj.v_xlsx_path.set(data_source)
				self.my_proj.initial_dir = self.proj_dir_path
				self.my_proj.pro_top.deiconify()
				#Always get focused
				self.my_proj.pro_top.grab_set()
				return
		else:
			print("DEBUG here")
			s = "创建失败，同名文件夹已经存在!"
			showerror(message=s)
			return

		self.pro_top.destroy()
		return
	#########confirm_project()#####################


	def discard_project(self):
		'''
		退出按钮函数
		'''
		self.proj_dir_path = None
		self.pro_top.destroy()
	##########discard_project()#####################


	def copy_data_template(self):
		'''
		拷贝样版数据源
		'''
		target_dir = os.path.join(self.proj_dir_path,"数据源")
		src_dir = os.path.join(os.getcwd(),"project_template"+os.sep+"数据源")
		result = my_bak.bak_directory(src_dir, target_dir)
		return result
	##########copy_data_template()##################


###############class NewProj(object):#####################


def check_new_project():
	print("this is new project")

if __name__ == '__main__':
	print("main start")

	top = tk.Tk()
	tk.Button(top, text="Check", command=check_new_project).pack()
	new_proj = NewProj(top)
	top.mainloop()	

	print("main done")