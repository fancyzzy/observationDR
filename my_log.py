#!/usr/bin/env python3

'''
日志，等全局变量
'''

import queue

QUE = queue.Queue()
SENTINEL = object()

def printl(s):
	'''
	输出日志文件
	'''
	global QUE

	print(s)
	QUE.put(s.strip('\n'))
############printl()##############

if __name__ == '__main__':
	pass