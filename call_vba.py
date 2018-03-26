#!/usr/bin/env python3

import win32com.client

'''
通过python调用excel的VBA函数

'''

class MyVBA(object):
	def __init__(self, xlsm_path):
		self.xl = win32com.client.Dispatch("Excel.Application")
		self.xlsm = xlsm_path

	def call_vba(self):
		print("使用{}进行合并数据源".format(self.xlsm))
		self.xl.Workbooks.Open(Filename=self.xlsm, ReadOnly=1)
		self.xl.Application.Run("merge.xlsm!Merge_Sheets.Auto_Merge")
		self.xl.Application.Quit()
		del self.xl
		print("合并结束")


######class MyVBA##############

if __name__ == '__main__':
	print("Call VBA main:")

	xlsm_path = r"C:\Users\tarzonz\Desktop\python_call_vba\数据源\merge.xlsm"
	my_vba = MyVBA(xlsm_path)
	my_vba.call_vba()


	print("Main done!")
