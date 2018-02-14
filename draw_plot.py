#!/usr/bin/env python3

'''
画曲线图
'''
import matplotlib.pyplot as plt
import matplotlib.dates as mdate
from matplotlib.ticker import MultipleLocator
import os
from math import ceil

#显示中文
MARKERS_ARRAY=['.', ',', 'o', 'v', '^', '<', '>', '1', '2', '3', '4', '8', 's', 'p', \
'*', 'h', 'H', '+', 'x', 'D', 'd', '|', '_', 'P', 'X']
len_marker = len(MARKERS_ARRAY)
class MyPlot(object):
	def __init__(self):
		self.plt = plt
		self.plt.rcParams['figure.figsize'] = (11,5)
		self.plt.rcParams['font.sans-serif'] = ['SimHei']
		self.plt.rcParams['axes.unicode_minus'] = False
		self.plt.rcParams['xtick.direction'] = 'in' 
		self.plt.rcParams['ytick.direction'] = 'in' 
		self.plt.rcParams['xtick.labelsize'] = 'x-large'
		self.plt.rcParams['ytick.labelsize'] = 15

		#new
		#fig = self.plt.figure()
		#self.ax = self.plt.subplot(111)

	def draw_settlement_fig(self, date_list, value_arrays, sample_list,\
	 save_flag = True):
		'''
		画沉降监测图，横轴日期，纵轴是各个观测点的当天沉降数据
		共sample_list个观测点的曲线
		'''
		ln = len(sample_list)
		#self.ax.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d %H:%M:%S'))
		#self.plt.gca().xaxis.set_major_formatter(mdate.DateFormatter('%m-%d'))
		#self.plt.gca().xaxis.set_major_locator(mdate.DayLocator())
		fig = self.plt.figure()
		ax = self.plt.subplot(111)

		for i in range(ln):
			ax.plot(date_list, value_arrays[i],linewidth='1.0', linestyle='-',\
			 marker=MARKERS_ARRAY[-(i%len_marker)], markersize=12,label= sample_list[i])
			#self.ax.plot(date_list, value_arrays[i],linewidth='0.8', linestyle='-',\
			# marker=MARKERS_ARRAY[-(i%len_marker)], label= sample_list[i])

		#self.plt.legend(loc='upper center',bbox_to_anchor=(0.5,1.08),ncol=4,\
		#	fancybox=True,shadow=False)
		ax.legend(loc='center right',bbox_to_anchor=(1.,0.5),ncol=1,\
			fancybox=True,shadow=True,markerscale=1.1,borderpad = 0.5,\
			labelspacing=0.02,handlelength=1.3,columnspacing=0.02, fontsize=20)
		ax.grid(linewidth='0.8',linestyle='-.')

		#旋转横轴刻度文字
		#self.plt.gcf().autofmt_xdate()

		#self.plt.xlabel('日期', fontsize='x-large')
		self.plt.ylabel('累计变化量(mm)', fontsize=22)
		#多增加x轴的日期显示，空出来用于放置lengend
		self.plt.xlim(0,len(date_list)+1)

		self.plt.tight_layout()
		if save_flag:
			aim_path = 'temped_fig.png'
			self.plt.savefig(aim_path, format='png',dpi=200,bbox_inches='tight')
			#设置成close，负责会影响测斜图的长宽比
			#Issue 这里会造成main gui的错误退出
			#self.plt.close(fig)
			self.plt.close('all')
			return aim_path
		else:
			self.plt.show()
	#################draw_date_plot()##########################


	def draw_inclinometer_fig(self, deep_values, diff_values, acc_values, save_flag=True):
		'''
		画测斜变化图
		纵轴是深度值
		横轴是本次变化和累计变化的值
		共两条曲线
		'''

		self.plt = plt
		self.plt.rcParams['figure.figsize'] = (4.0,deep_values[-1]*1.12+4)
		self.plt.rcParams['font.sans-serif'] = ['SimHei']
		self.plt.rcParams['axes.unicode_minus'] = False
		self.plt.rcParams['xtick.direction'] = 'in' 
		self.plt.rcParams['ytick.direction'] = 'in' 
		self.plt.rcParams['xtick.labelsize'] = 8
		self.plt.rcParams['ytick.labelsize'] = 12
		self.plt.rcParams['xtick.major.size'] = 10
		self.plt.rcParams['xtick.major.width'] = 4

		fig = self.plt.figure()
		aax = self.plt.subplot(111)	

		aax.plot(diff_values, deep_values,linewidth='1.8', linestyle='-',\
			 marker='o', color='#2B303B', markersize=13,label= '本次变化(mm)')

		aax.plot(acc_values, deep_values,linewidth='1.8', linestyle='-',\
			 marker='s', markersize=10,color='#000000',label= '累计变化(mm)')


		#plt.gca().invert_yaxis()
		ax = self.plt.gca()
		ax.spines['right'].set_color('none')
		ax.spines['top'].set_color('none')
		#ax.yaxis.set_ticks_position('none')
		#ax.xaxis.set_ticks_position('top')
		#使y轴到x轴下面
		#bottom控制x轴
		ax.spines['bottom'].set_position(('data',0))
		ax.spines['bottom'].set_linewidth(4.5)
		#使y轴到中心位置
		#left控制y轴
		ax.spines['left'].set_position(('data',0))
		ax.spines['left'].set_linewidth(4)
		#使y轴多显示，空白出来用来放legend
		#self.plt.ylim(0,len(acc_values))
		ax.invert_yaxis()
		#设置x轴标签的位置，靠外
		ax.xaxis.labelpad = -150

		#y坐标离轴多远
		ax.tick_params(axis='y', which='major', pad=20)
		ax.tick_params(axis='x', which='major', pad=-44)

		self.plt.legend(loc='lower center',ncol=1,bbox_to_anchor=(0.5,-0.180),\
			markerscale=1.5,borderpad=0.5,labelspacing=0.52,frameon=False,\
			handlelength=2.0,columnspacing=0.12, fontsize=28,fancybox=False,\
			shadow=False)

		y = self.plt.ylabel('深度(m)', fontsize=40,labelpad=20)
		#y.set_rotation(0)
		self.plt.xlabel('变化量(mm)', fontsize=40)
		#使y轴显示整数深度的刻度
		self.plt.yticks(list(map(ceil,deep_values)))
		#x轴刻度间隔倍数
		max_value = ceil(max(list(map(abs,acc_values))))
		xmajorLocator   = MultipleLocator(int(max_value/2))
		ax.xaxis.set_major_locator(xmajorLocator)
		#刻度数值大小
		self.plt.tick_params(axis='both', labelsize=32)

		#self.plt.tight_layout()
		if save_flag:
			aim_path = 'inclinometer_fig.PNG'
			self.plt.savefig(aim_path, format='png',dpi=300,bbox_inches='tight')
			#self.plt.clf()
			#self.plt.close(fig)
			self.plt.close('all')
			return aim_path
		else:
			self.plt.show()

		demo_fig = r'C:\Users\tarzonz\Desktop\observationDR\inclinometer_fig.PNG'
		return demo_fig
	###########draw_inclinometer_fig()########################




if __name__ == '__main__':
	
	my_plot = MyPlot()	

	x = ['day1','day2','day3','day4','day5','day6','day7']
	x = ['2018年01月01日', '2017年12月31日', '2017年12月30日', '2017年12月29日',\
	 '2017年12月28日', '2017年12月27日', '2017年12月26日']
	x = ['2018/01/01', '2017/12/31', '2017/12/30', '2017/12/29',\
	 '2017/12/28', '2017/12/27', '2017/12/26']

	import datetime
	'''
	x = [datetime.datetime(2018, 1, 1, 0, 0), datetime.datetime(2017, 12, 31, 0, 0), datetime.datetime(2017, 12, 30, 0, 0), datetime.datetime(2017, 12, 29, 0, 0), datetime.datetime(2017, 12, 28, 0, 0), datetime.datetime(2017, 12, 27, 0, 0), datetime.datetime(2017, 12, 26, 0, 0)]
	'''
	value_arrays = [[-0.05,-0.2,-0.17,-0.8,-0.2,0.1,0.2],
					[1,2,2,-0.2,-0.12,0.0,0.0],
					[3,2,1,0,0.12,1.2,3],
					[3,1,1,0,1.12,1.2,2],
					[3,2,1,0,0.12,1.2,1.3],
					[-3,-2,-1,0,0.15,1.21,0.3],
					[-2.3,-1.2,1,0.01,0.12,-1.12,-2.3],
					[3,2,1,0,0.12,-1.2,3],
					[3,2,1,0,0.12,-1.2,3],
					[3,3.2,2.1,0,-0.12,1.2,-1.3],
					[-1.3,-2,1,0,0.32,-1.6,3.2],]
	sample_list = ['one','two','three','four','five','six','seven','eight','nine'\
	,'ten','eleven']

	'''
	#测试沉降图
	'''
	my_plot.draw_settlement_fig(x,value_arrays,sample_list,False)
	#my_plot.draw_settlement_fig(x,value_arrays,sample_list,True)

	#测试测斜图
	from numpy import array

	deep_values = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0,
				   5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 
				   10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0, 15.5]
	
	diff_values = [-0.03, -0.1, -0.13, 0.03, -0.02, -0.02, -0.09, 0.07, -0.09, -0.07, 
				   0.22, -0.13, 0.07, -0.08, -0.23, 0.09, 0.21, 0.19, -0.17, 0.21, 
				   0.15, -0.0, -0.29, 0.09, -0.09, 0.07, -0.06, 0.02, 0.05, 0.04, -0.11]

	#acc_values = [-14.43, -12.0, -10.44, -7.21, -5.48, -3.91, -0.45, 3.53, 3.18, 6.44, 
	acc_values = [-14.43, -12.0, -10.44, -7.21, -5.48, -3.91, -0.45, 3.53, 3.18, None, 
				   7.9, 7.93, 8.11, 10.35, 9.07, 7.95, 11.19, 11.85, 10.88, 9.04,
				   11.07, 13.43, 6.93, 8.34, 7.08, 6.68, 7.23, 6.74, 6.38, 3.6, 0.33]

	acc_values = array(acc_values, dtype=float)
	my_plot.draw_inclinometer_fig(deep_values,diff_values,acc_values,False)
