#!/usr/bin/env python3

import win32com.client
import os

'''
通过python调用excel的VBA函数

'''

class MyVBA(object):
	def __init__(self, xlsm_path):
		import pythoncom
		pythoncom.CoInitialize()
		self.xl = win32com.client.Dispatch("Excel.Application")
		self.xlsm = xlsm_path

	def call_vba(self):
		print("使用{}进行合并数据源".format(self.xlsm))
		
		self.xl.Workbooks.Open(Filename=self.xlsm, ReadOnly=1)
		self.xl.Application.Run("merge.xlsm!Merge_Sheets.Auto_Merge")
		self.xl.Application.Quit()
		del self.xl
		print("合并结束")
		return True
		'''	
		except Exception as e:
			print("合并出问题:",e)
			return False
		'''



def start_vba(xlsm_path):
	import pythoncom
	pythoncom.CoInitialize()
	xl = win32com.client.Dispatch("Excel.Application")
	xlsm = os.path.normpath(xlsm_path)
	print("使用{}进行合并数据源".format(xlsm))
	xl.Workbooks.Open(Filename=xlsm, ReadOnly=1)
	xl.Application.Run("merge.xlsm!Merge_Sheets.Auto_Merge")
	xl.Application.Quit()
	del xl
	print("合并结束")
######class MyVBA##############

if __name__ == '__main__':
	print("Call VBA main:")

	xlsm_path = r"C:\Users\tarzonz\Desktop\python_call_vba\数据源\merge.xlsm"
	my_vba = MyVBA(xlsm_path)
	my_vba.call_vba()
	#start_vba(xlsm_path)


	print("Main done!")
