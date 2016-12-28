#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd
import xlwt
import os, sys
import datetime
import re
import uniout

class order:
	def __init__(self):
		self.id = 0
		self.name = ''
		self.order_num = 0

	def __str__(self):
		return u"id=%d,name=%s,order_num=%d" %(self.id, self.name,
		self.order_num)

class order_list:
	def __init__(self, order_file_path, model_list):
		self.order_list = []
		self.order_srow = 2
		self.order_file_path = order_file_path
		self.model_reco_list = [] #可以在model表找到的order
		self.model_unre_list = [] #没有在model表找到的order
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
			#print('[%s]  %s' %(curr_row, m_row))
			m = re.findall(r'(\w*[0-9]+)\w*',m_row[3])
			if m_row[1]=='' or len(m) <= 0:
				print("name is null or can't convert %s to num." %m_row[3])
				continue
			od = order()
			od.id = order_id
			order_id = order_id + 1
			o_name = re.sub(r'\s+', ' ', m_row[1]) # 合并连续多个空格为一个
			od.name = o_name
			od.order_num = int(m[0])
			self.order_list.append(od)
		for o in self.order_list:
			print("%s %s %s" %(o.id,o.name,o.order_num))

	def merge(self, model_list):
		

	def exprt_data(self):
		row0 = [u'编号',u'名称型号',u'数量']
		if not os.path.exists('./output/'):
			os.mkdir('./output')
		output_path = './output/' + datetime.datetime.now().strftime("%Y%m%d_%H%M%S/")
		os.mkdir(output_path)
		o_workbook = xlwt.Workbook()
		sheet1 = o_workbook.add_sheet('sheet1',cell_overwrite_ok=True)
		for i in range(0,len(row0)):
			sheet1.write(0,i,row0[i])

		o_id = 1
		#写入在model表里面的结果
		for i in range(0,len(self.model_reco_list)):
			sheet1.write(o_id,0,o_id)
			sheet1.write(o_id,1,self.model_reco_list[i].name)
			sheet1.write(o_id,2,self.model_reco_list[i].order_num)
			o_id = o_id + 1

		for i in range(0,len(self.model_unre_list)):
			sheet1.write(o_id,0,o_id)
			sheet1.write(o_id,1,self.model_unre_list[i].name)
			sheet1.write(o_id,2,self.model_unre_list[i].order_num)
			o_id = o_id + 1

		o_workbook.save(output_path+u'型号归并表'+'.xlsx')
		return [len(model_reco_list) len(model_unre_list)]

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
		print('model_file worksheets is %s' %model_sheets)
		model_sheet = model_workbook.sheet_by_name(model_sheets[0])
		num_rows = model_sheet.nrows
		m_cols = model_sheet.col_values(self.model_scol)
		self.model_list = m_cols[model_srow:]
		for i in range(len(self.model_list)): # 合并连续多个空格为一个
			self.model_list[i] = re.sub(r'\s+', ' ', self.model_list[i]) 
	
if __name__=="__main__":
	order_file_path = './order.xls'
	model_file_path = './model.xlsx'
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
	result = ol.exprt_data()
	print(u"导出%d个可识别行, %d个不可识别行, 完成！" %(result[0],result[1]))
	i=raw_input("Press enter key to continue...")

