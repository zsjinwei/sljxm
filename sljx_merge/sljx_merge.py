#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd
import xlwt
import os, sys
import datetime
import re
#import uniout
from collections import OrderedDict

class order:
	def __init__(self):
		self.id = 0
		self.name = ''
		self.order_num = 0
		self.match_name = ''
		self.dajian = ''
		self.changdu = ''
		self.xiaojian = ''
		self.liangzhu = ''
		self.longmeng = ''
		self.sizhu = ''
		self.special = ''

	def __str__(self):
		return u"id=%d,name=%s,order_num=%d" %(self.id, self.name,
		self.order_num)

	def is_match(self, ml):
		if (self.dajian==ml.dajian) and\
		(self.changdu==ml.changdu) and\
		(self.xiaojian==ml.xiaojian) and\
		(self.liangzhu==ml.liangzhu) and\
		(self.longmeng==ml.longmeng) and\
		(self.sizhu==ml.sizhu):
			return True
		else:
			return False

class order_list:
	def __init__(self, order_file_path):
		self.order_list = []
		self.order_srow = 1
		self.order_file_path = order_file_path
		self.sheet_map = OrderedDict({'sheet1':[]}) # sheet1 contains class model, others contain order
		self.load_data()

	def load_data(self):
		order_workbook = xlrd.open_workbook(self.order_file_path)
		order_sheets = order_workbook.sheet_names()
		print('order_file worksheets is %s' %order_sheets)
		order_sheet = order_workbook.sheet_by_name(order_sheets[0])
		num_rows = order_sheet.nrows
		order_id = 1
		for curr_row in range(self.order_srow, num_rows):
			m_row = order_sheet.row_values(curr_row)
			#if len(m_row) < 4 and m_row[0]=='' and m_row[1]=='' and m_row[2]=='' and m_row[3]=='':
			#	break #default null row is finishing to read
			#print('[%s]  %s' %(curr_row, m_row))
			if type(m_row[3])==int or type(m_row[3])==float:
				m_row[3] = str(int(m_row[3])) # avoid to read in a int num
			m = re.findall(r'(\w*[0-9]+)\w*',m_row[3])
			if m_row[1]=='' or len(m) <= 0:
				print("name is null or can't convert %s to num." %m_row[3])
				continue
			od = order()
			od.id = order_id
			if type(m_row[1])==int or type(m_row[1])==float:
				m_row[1] = str(int(m_row[1])) # avoid to read in a int num
			o_name = re.sub(r'\s+', ' ', m_row[1]) # 合并连续多个空格为一个
			od.name = o_name
			od.order_num = int(m[0])
			od.dajian = m_row[6]
			od.changdu = m_row[7]
			od.xiaojian = m_row[8]
			od.liangzhu = m_row[9]
			od.longmeng = m_row[10]
			od.sizhu = m_row[11]
			od.special = m_row[12]
			order_id = order_id + 1
			self.order_list.append(od)
		for o in self.order_list:
			print("%s %s %s %s" %(o.id,o.name,o.order_num,o.special))

	def merge(self, model_list):
		# first to build model name sheets
		for ml in model_list.model_list:
			self.sheet_map[ml.name] = []
		self.sheet_map['unknow'] = []
		for od in self.order_list:
			found_flag = False
			for ml in model_list.model_list:
				if od.special==u'是':
					break
				if od.is_match(ml):
					self.sheet_map[ml.name].append(od)
					print("Found %s in model list(%s),num=%d" %(od.name, ml.name, od.order_num))
					found_flag = True
					break
			if not found_flag:
				print("%s is not found" %(od.name))
				self.sheet_map['unknow'].append(od)
		# merge other sheets into sheet1
		for s_key,s_val in self.sheet_map.items():
			if (not s_key=='sheet1') and (not s_key=='unknow'):
				order_count = 0;
				for od in s_val:
					order_count = order_count + od.order_num
					print("%s order count add %d" %(s_key,od.order_num))
				print("%s order total = %d" %(s_key,order_count))
				sheet1_row = model()
				sheet1_row.name = s_key
				sheet1_row.merge_num = order_count
				self.sheet_map['sheet1'].append(sheet1_row)
			elif s_key=='unknow':
				for od in s_val:
					ml_unknow = model()
					ml_unknow.cp_from_od(od)
					self.sheet_map['sheet1'].append(ml_unknow)
		#print(self.sheet_map['sheet1'])

	def export_data(self):
		row0 = [u'编号',u'名称型号',u'数量']
		if not os.path.exists('./output/'):
			os.mkdir('./output')
		output_path = './output/' + datetime.datetime.now().strftime(u"%Y%m%d_%H%M%S_model/")
		os.mkdir(output_path)
		o_workbook = xlwt.Workbook()
		# export sheet1 first
		sheet1 = o_workbook.add_sheet('sheet1',cell_overwrite_ok=True)
		for i in range(0,len(row0)):
			sheet1.write(0,i,row0[i])
		it_id = 1
		#写入在model表里面的结果
		for it in self.sheet_map['sheet1']:
			sheet1.write(it_id,0,it_id)
			sheet1.write(it_id,1,it.name)
			sheet1.write(it_id,2,it.merge_num)
			it_id = it_id + 1
		#export other sheets
		for s_key,s_val in self.sheet_map.items():
			if len(s_val)==0 or s_key=='sheet1':
				continue
			print("export %s: %d" %(s_key, len(s_val)))
			sheet1 = o_workbook.add_sheet(s_key,cell_overwrite_ok=True)
			for i in range(0,len(row0)):
				sheet1.write(0,i,row0[i])
			it_id = 1
		#写入在model表里面的结果
			for it in s_val:
				sheet1.write(it_id,0,it_id)
				sheet1.write(it_id,1,it.name)
				sheet1.write(it_id,2,it.order_num)
				it_id = it_id + 1
		o_workbook.save(output_path+u'型号归并表'+'.xls')

class model:
	def __init__(self):
		self.id = 0
		self.name = ''
		self.merge_num = 0
		self.dajian = ''
		self.changdu = ''
		self.xiaojian = ''
		self.liangzhu = ''
		self.longmeng = ''
		self.sizhu = ''

	def cp_from_od(self, od):
		self.id = od.id
		self.name = od.name
		self.dajian = od.dajian
		self.changdu = od.changdu
		self.xiaojian = od.xiaojian
		self.liangzhu = od.liangzhu
		self.longmeng = od.longmeng
		self.sizhu = od.sizhu
		self.merge_num = od.order_num

class model_list:
	def __init__(self, model_file_path):
		self.model_list = []
		self.model_srow = 1
		self.model_scol = 1
		self.model_file_path = model_file_path
		self.load_data()

	def load_data(self):
		model_workbook = xlrd.open_workbook(self.model_file_path)
		model_sheets = model_workbook.sheet_names()
		print('order_file worksheets is %s' %model_sheets)
		model_sheet = model_workbook.sheet_by_name(model_sheets[0])
		num_rows = model_sheet.nrows
		model_id = 1
		for curr_row in range(self.model_srow, num_rows):
			m_row = model_sheet.row_values(curr_row)
			m = re.findall(r'(\w*[0-9]+)\w*',m_row[3])

			ml = model()
			ml.id = model_id
			if type(m_row[0])==int or type(m_row[0])==float:
				m_row[0] = str(int(m_row[0])) # avoid to read in a int num
			print(m_row[0])
			o_name = re.sub(r'\s+', ' ', m_row[0]) # 合并连续多个空格为一个
			ml.name = o_name
			ml.dajian = m_row[4]
			ml.changdu = m_row[5]
			ml.xiaojian = m_row[6]
			ml.liangzhu = m_row[7]
			ml.longmeng = m_row[8]
			ml.sizhu = m_row[9]
			model_id = model_id + 1
			self.model_list.append(ml)

if __name__=="__main__":
	order_file_path = u'./orders.xls'
	model_file_path = './models.xlsx'
	print(u"**********************************")
	print(u"欢迎使用顺力机械有限公司订单归并工具")
	print(u"版本: V1.0")
	print(u"制作: 黄锦伟")
	print(u"2016-12 @ 珠海")
	print(u"**********************************")
	print(u"检测备货单列表("+order_file_path+u")文件...")
	if not os.path.exists(order_file_path):
		print(u"没有发现"+order_file_path+u", 请将其放在当前目录.")
		i=raw_input("Press enter key to continue...")
		exit()
	else:
		print(u"发现"+order_file_path)
	print(u"检测型号列表("+model_file_path+u")文件...")
	if not os.path.exists(model_file_path):
		print(u"没有发现"+model_file_path+u", 请将其放在当前目录.")
		i=raw_input("Press enter key to continue...")
		exit()
	else:
		print(u"发现"+model_file_path)
	i=raw_input("Press enter key to continue...")
	print(u"读取备货单...")
	ol = order_list(order_file_path)
	print(u"读取型号列表...")
	ml = model_list(model_file_path)
	print(u"合并数据...")
	ol.merge(ml)
	print(u"导出数据...")
	result = ol.export_data()
	i=raw_input("Press enter key to continue...")

