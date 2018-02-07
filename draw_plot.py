#!/usr/bin/env python3

'''
画曲线图
'''
import matplotlib.pyplot as plt
import matplotlib.dates as mdate
import os
#显示中文
MARKERS_ARRAY=['.', ',', 'o', 'v', '^', '<', '>', '1', '2', '3', '4', '8', 's', 'p', \
'*', 'h', 'H', '+', 'x', 'D', 'd', '|', '_', 'P', 'X']
len_marker = len(MARKERS_ARRAY)
class MyPlot(object):
	def __init__(self):
		self.plt = plt
		self.plt.rcParams['font.sans-serif'] = ['SimHei']
		self.plt.rcParams['axes.unicode_minus'] = False
		self.plt.rcParams['xtick.direction'] = 'in' 
		self.plt.rcParams['ytick.direction'] = 'in' 

		#new
		#fig = self.plt.figure()
		#self.ax = self.plt.subplot(111)

	def draw_date_plot(self, date_list, value_arrays, sample_list,\
	 save_flag = True):
		ln = len(sample_list)

		print("DEBUG draw_date_plot")
		print("DEBUG date_list=",date_list)
		print("DEBUG sample_list=",sample_list)

		#self.ax.xaxis.set_major_formatter(mdate.DateFormatter('%Y-%m-%d %H:%M:%S'))
		#self.plt.gca().xaxis.set_major_formatter(mdate.DateFormatter('%m-%d'))
		#self.plt.gca().xaxis.set_major_locator(mdate.DayLocator())

		for i in range(ln):
			self.plt.plot(date_list, value_arrays[i],linewidth='0.8', linestyle='-',\
			 marker=MARKERS_ARRAY[-(i%len_marker)], label= sample_list[i])
			#self.ax.plot(date_list, value_arrays[i],linewidth='0.8', linestyle='-',\
			# marker=MARKERS_ARRAY[-(i%len_marker)], label= sample_list[i])

		#self.plt.legend(loc='upper center',bbox_to_anchor=(0.5,1.08),ncol=4,\
		#	fancybox=True,shadow=False)
		self.plt.legend(loc='center right',bbox_to_anchor=(1.1,0.5),ncol=2,\
			fancybox=True,shadow=True,markerscale=0.8,borderpad = 0.1,labelspacing=0.5,handlelength=1.2,\
			columnspacing=0.2)
		self.plt.grid(color='#9B9B9B',linewidth='0.5',linestyle='-.')

		#self.plt.gcf().autofmt_xdate()


		self.plt.xlabel('日期')
		self.plt.ylabel('累计变化量(mm)')
		self.plt.xlim(0,len(date_list)+1)

		'''
		self.plt.grid(color='#9B9B9B',linewidth='0.5',linestyle='-.')
		#self.plt.xlim(0,len(date_list))

		#self.plt.legend(loc = 'upper center')
		self.plt.legend()
		self.plt.xlabel('日期')
		self.plt.ylabel('累计变化量(mm)')
		'''

		if save_flag:
			aim_path = 'temped_fig.png'
			self.plt.savefig(aim_path, format='png',dpi=300)
			print("DEBUG png file: aim_path has been saved",os.path.join(os.getcwd(),aim_path))
			self.plt.clf()
			return aim_path
		else:
			self.plt.show()
	#################draw_date_plot()##########################



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

	my_plot.draw_date_plot(x,value_arrays,sample_list,False)
	#my_plot.draw_date_plot(x,value_arrays,sample_list,True)
