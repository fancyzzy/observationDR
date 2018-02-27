#!/usr/bin/env python3

'''
日志，等全局变量
'''

import queue
import os

QUE = queue.Queue()
SENTINEL = object()


LOG_PATH = ['my_log.txt']

def write_log(s,file_path):
	with open(file_path, "ab+") as fobj:
			s = s + os.linesep
			s = s.encode('utf-8')
			fobj.write(s)

def printl(s,que_output=True):
	'''
	输出日志文件
	'''
	global QUE
	global LOG_PATH

	if '@' in s:
		pass
	else:
		print(s)
		write_log(s, LOG_PATH[0])

	if que_output:
		QUE.put(s.strip('\n'))
############printl()##############

if __name__ == '__main__':
	pass