#!/usr/bin/env python3


'''
shutil.copytree() study
'''
import shutil
import os


def bak_directory(s_dir, bak_dir):
	is_existed = False
	if os.path.exists(bak_dir) and os.path.isdir(bak_dir):
		print('文件夹已经存在, 进行增量备份')
		file_list = os.listdir(s_dir)
		existed_list = os.listdir(bak_dir)
		new_file_list = list(set(file_list)-set(existed_list))
		if len(new_file_list) > 0: 
			print("有新文件:",new_file_list)
			for file in new_file_list:
				full_path = os.path.join(s_dir, file)
				if os.path.isfile(full_paht):
					try:
						shutil.copy(full_path, bak_dir)
					except Exception as e:
						print("Error!, e:{}".format(e))
				elif os.path.isdir(full_path):
					bak_directory(full_path, os.path.join(bak_dir,file))
		else:
			print("没有文件要备份")
	else:
		try:
			shutil.copytree(s_dir, bak_dir)
		except Exception as e:
			print("ERROR, e: {}".format(e))
			return False
	print("bak done:{} -> {}".format(s_dir,bak_dir))
##################bak_dir()###############################



DIR_PATH = r'C:\Users\tarzonz\Desktop\演示工程A\平面布点图'
DIR_PATH2 = r'C:\Users\tarzonz\Desktop\struct_test'
BAK_PATH = r'C:\Users\tarzonz\Desktop\BAK'

if __name__ == '__main__':
	print("DEBUG start testing mybak.py, main")
	try:
		bak_directory(DIR_PATH, BAK_PATH)
	except Exception as e:
		print("文件已经存在？e:",e)
		print("执行增量备份")

		file_list = os.listdir(DIR_PATH)
		existed_list = os.listdir(BAK_PATH)
		new_file_list = list(set(file_list)-set(existed_list))
		print("DEBUG new_file_lisit=",new_file_list)

		for file in new_file_list:
			shutil.copy(os.path.join(DIR_PATH,file), BAK_PATH)

	#bak_dir(DIR_PATH2, BAK_PATH)

	print("Done")
