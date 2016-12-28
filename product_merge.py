#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd
import xlwt
import os, sys
import datetime
import re
#import uniout

class order:
	def __init__(self):
		self.id = 0
		self.name = ''
		self.order_num = 0

	def __str__(self):
		return u"id=%d,name=%s,order_num=%d" %(self.id, self.name,
		self.order_num)

class order_list:
	def __init__(self, order_file_path):
		self.order_list = []
		self.unrecon_list = []
		self.order_srow = 2
		self.order_file_path = order_file_path
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
			od.name = m_row[1]
			od.order_num = int(m[0])
			self.order_list.append(od)
		for o in self.order_list:
			print("%s %s %s" %(o.id,o.name,o.order_num))

if __name__=="__main__":
	order_file_path = './order.xls'
	ol = order_list(order_file_path)
