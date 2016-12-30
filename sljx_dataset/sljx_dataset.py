#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd
import xlwt
import os, sys
import datetime

'''
产品序号	产品名称	产品型号	产品库存	产品生产计划
1	35DL	35DLM	0	10
2	36DL	36DLM	0	0
3	37DL	37DLM	0	12
'''
class product: # pruducts for menufacture
	def __init__(self, name, p4m_file_path, e4p_file_path, i4e_file_path):
		self.id = 0
		self.name = name
		self.model = ''
		self.reserve = 0
		self.order = 0
		self.e4p = e4p(self.name,
			e4p_file_path,
			i4e_file_path) # slement list for this product

	def __str__(self):
		return "id=%d,name=%s,model=%s,"\
		"reserve=%s,order=%s" %(self.id,
		self.name,self.model, self.reserve,
		self.order)
'''
零件编号	零件名称	零件型号	零件数量	零件库存
1	36DL-A	36DL-AM	1	3
2	36DL-B	36DL-BM	2	0
3	36DL-C	36DL-CM	1	2
4	36DL-D	36DL-DM	2	1
'''
class element: # elements for a product
	def __init__(self, name, e4p_file_path, i4e_file_path):
		self.id = 0
		self.name = name
		self.model = ''
		self.per_num = 0 # how many this element for a product
		self.reserve = 0
		self.product = ''
		self.i4e = i4e(self.name, i4e_file_path)

	def __str__(self):
		return "\tid=%d,name=%s,model=%s,"\
		"per_num=%d,reserve=%s,product=%s" %(self.id,
		self.name,self.model, self.per_num,
		self.reserve, self.product)
'''
配件序号	配件名称	配件型号	配件数量	配件规格	配件所属产品	配件所属零件	A部	B部	C部	D部	E部	F部	G部	H部	I部	J部
1	37DL-D-A	37DL-D-AM	2	12u	37DL	37DL-D	y				y		y			y
2	37DL-D-B	37DL-D-BM	2	5u	37DL	37DL-D		y		y					y
'''
class ingredient:
	def __init__(self, name):
		self.id = 0
		self.name = name
		self.model = ''
		self.per_num = 0
		self.spec = ''
		self.product = ''
		self.element = ''
		self.departments = []
		self.reserve = 0
		self.total_num = 0

	def __str__(self):
		return "\t\tid=%d,name=%s,model=%s,"\
		"per_num=%d,spec=%s,product=%s,"\
		"element=%s" %(self.id, self.name,
		self.model, self.per_num, self.spec,
		self.product, self.element)

	def __getitem__(self, item):
		return self.spec

'''
部门序号	部门名称	是否按规格排序
1	A部	N
2	B部	Y
3	C部	N
4	D部	N
5	E部	N
6	F部	N
7	G部	Y
8	H部	N
9	I部	N
10	J部	N
'''
class department: # ingredients for a element
	def __init__(self):
		self.id = 0
		self.name = ''
		self.need_sort = 0

class p4m:
	def __init__(self, p4m_file_path, e4p_file_path, i4e_file_path):
		self.m_srow = 1 # skip header
		#
		self.p4m_file_path = p4m_file_path
		self.e4p_file_path = e4p_file_path
		self.i4e_file_path = i4e_file_path
		self.product_list = []
		self.d_map = {}
		self.load_data()

	def load_data(self):
		p4m_workbook = xlrd.open_workbook(self.p4m_file_path)
		p4m_sheets = p4m_workbook.sheet_names()
		print('p4m_file worksheets is %s' %p4m_sheets)
		p4m_sheet = p4m_workbook.sheet_by_name(p4m_sheets[0])
		num_rows = p4m_sheet.nrows
		for curr_row in range(self.m_srow, num_rows):
			m_row = p4m_sheet.row_values(curr_row)
			print('p4m_file row%s is %s' %(curr_row, m_row))
			# construct class product
			p = product(m_row[1],
				self.p4m_file_path,
				self.e4p_file_path,
				self.i4e_file_path)
			# init members except name
			p.id = int(m_row[0])
			p.model = m_row[2]
			p.reserve = int(m_row[3])
			p.order = int(m_row[4])
			print(p)
			# append to product_list
			self.product_list.append(p)

	def get_department_data(self):
		for cur_p in self.product_list:
			cur_p.e4p.get_department_data(self.d_map, cur_p.order)

	def export_data(self):
		if not os.path.exists('./output/'):
			os.mkdir('./output')
		output_path = './output/' + datetime.datetime.now().strftime("%Y%m%d_%H%M%S/")
		os.mkdir(output_path)
		for d_name, i_list in self.d_map.items():
			d_workbook = xlwt.Workbook()
			sheet1 = d_workbook.add_sheet('sheet1',cell_overwrite_ok=True)
			row0 = [u'编号',u'分配部门',u'所属型号',u'所属零件',u'规格', u'数量']
			#生成第一行
			for i in range(0,len(row0)):
				sheet1.write(0,i,row0[i])

			for j in range(0, len(i_list)):
				sheet1.write(j+1,0,j+1)
				sheet1.write(j+1,1,d_name)
				sheet1.write(j+1,2,i_list[j].product)
				sheet1.write(j+1,3,i_list[j].element)
				sheet1.write(j+1,4,i_list[j].spec)
				sheet1.write(j+1,5,i_list[j].total_num)
				#保存该excel文件,有同名文件时直接覆盖
				d_workbook.save(output_path+d_name+'.xlsx')

	def department_sort(self, d4s_file_path):
		if not os.path.exists(d4s_file_path):
			print("Can't not find "+d4s_file_path+" skip sort.")
			return
		d4s_workbook = xlrd.open_workbook(d4s_file_path)
		d4s_sheets = d4s_workbook.sheet_names()
		print('\td4s_file worksheets is %s' %d4s_sheets)
		d4s_sheet = d4s_workbook.sheet_by_name(d4s_sheets[0])
		num_rows = d4s_sheet.nrows
		ds_map = {}
		for curr_row in range(1, num_rows):
			p_row = d4s_sheet.row_values(curr_row)
			ds_map[p_row[1]] = p_row[2]
		print(ds_map)
		for ds_key,ds_val in ds_map.items():
			if self.d_map.has_key(ds_key) and ds_val == 'Y':
				self.d_map[ds_key].sort(lambda x,y:cmp(x[4],y[4]))

class e4p:
	def __init__(self, p_name, e4p_file_path, i4e_file_path):
		self.p_srow = 2 # start row of elements list(begin with 0)
		#
		self.e4p_file_path = e4p_file_path
		self.i4e_file_path = i4e_file_path
		self.p_name = p_name
		self.element_list = []
		self.load_data()

	def load_data(self):
		e4p_workbook = xlrd.open_workbook(self.e4p_file_path)
		e4p_sheets = e4p_workbook.sheet_names()
		print('\te4p_file worksheets is %s' %e4p_sheets)
		if self.p_name in e4p_sheets:
			e4p_sheet = e4p_workbook.sheet_by_name(self.p_name)
			print("\tproduct %s found in e4p_file." %self.p_name)
		else:
			print("Product %s is not found in %s." %(p_name, e4p_file_path))
			return
		num_rows = e4p_sheet.nrows
		for curr_row in range(self.p_srow, num_rows):
			p_row = e4p_sheet.row_values(curr_row)
			print('\te4p_file row%s is %s' %(curr_row,p_row))
			# construct class product
			e = element(p_row[1],
				self.e4p_file_path,
				self.i4e_file_path) # init members except name
			e.id = int(p_row[0])
			e.model = p_row[2]
			e.per_num = int(p_row[3])
			e.reserve = int(p_row[4])
			e.product = self.p_name
			print(e)
			# append to product_list
			self.element_list.append(e)

	def get_department_data(self, d_map, p_num):
		for cur_e in self.element_list:
			e_num = cur_e.per_num * p_num - cur_e.reserve
			#if e_num < 0:
			#	e_num = 0
			cur_e.i4e.get_department_data(d_map, e_num)

class i4e: # ingredients for an element
	def __init__(self, e_name, i4e_file_path):
		self.depart_srow = 1 # department name(table header) start row(begin with 0)
		self.depart_scol = 7
		self.i_srow = 2 # start row number for content(begin with 0)
		#
		self.i4e_file_path = i4e_file_path
		self.ingredient_list = []
		self.e_name = e_name
		self.load_data()

	def load_data(self):
		i4e_workbook = xlrd.open_workbook(self.i4e_file_path)
		i4e_sheets = i4e_workbook.sheet_names()
		print('\t\ti4e_file worksheets is %s' %i4e_sheets)
		if self.e_name in i4e_sheets:
			i4e_sheet = i4e_workbook.sheet_by_name(self.e_name)
			print("\t\telement %s found in i4e_file." %self.e_name)
		else:
			print("Element %s is not found in %s." %(e_name, i4e_file_path))
			return
		num_rows = i4e_sheet.nrows
		department_row = i4e_sheet.row_values(self.depart_srow)
		department_name = department_row[self.depart_scol:]
		#print('%s' %(department_name))
		for curr_row in range(self.i_srow, num_rows):
			i_row = i4e_sheet.row_values(curr_row)
			print('\t\ti4e_file row%s is %s' %(curr_row,i_row))
			# construct class product
			i = ingredient(i_row[1]) # init members except name
			i.id = int(i_row[0])
			#i.name = i_row[1]
			i.model = i_row[2]
			i.per_num = int(i_row[3])
			i.spec = i_row[4]
			i.product = i_row[5]
			i.element = i_row[6]
			i.reserve = 0
			for d_i in range(len(department_name)):
				if i_row[d_i+self.depart_scol]!='':
					i.departments.append(department_name[d_i])
			#print(i)
			print("\t\tdepartments: %s" %i.departments)
			# append to product_list
			self.ingredient_list.append(i)

	def get_department_data(self, d_map, e_num):
		for cur_i in self.ingredient_list:
			#print("\t\t\t%s" %cur_i)
			cur_i.total_num = cur_i.per_num * e_num
			if cur_i.total_num <= 0:
				continue
			for cur_dm_name in cur_i.departments:
				if d_map.has_key(cur_dm_name):
					d_map[cur_dm_name].append(cur_i)
				else:
					d_map[cur_dm_name]= [cur_i]

if __name__=="__main__":
	p4m_file_path = './p4m.xlsx'
	e4p_file_path = './e4p.xlsx'
	i4e_file_path = './i4e.xlsx'
	d4s_file_path = './d4s.xlsx'
	'''
	p4m = p4m(p4m_file_path, e4p_file_path, i4e_file_path)
	depart_data = p4m.get_depart_data()'''
	# test for class i4e
	#i4e = i4e('37DL-D', i4e_file_path)
	# test for class e4p
	#e4p = e4p('37DL', e4p_file_path, i4e_file_path)
	# test for class p4m
	print(u"**********************************")
	print(u"欢迎使用顺力机械有限公司生产派单工具")
	print(u"版本: V1.0")
	print(u"制作: 黄锦伟")
	print(u"2016-12 @ 珠海")
	print(u"**********************************")
	print(u"检测计划生产型号列表("+p4m_file_path+u")文件...")
	if not os.path.exists(p4m_file_path):
		print(u"没有发现"+p4m_file_path+u", 请将其放在当前目录.")
		i=raw_input("Press enter key to continue...")
		exit()
	else:
		print(u"发现"+p4m_file_path)
	print(u"检测型号所需零件列表("+e4p_file_path+u")文件...")
	if not os.path.exists(e4p_file_path):
		print(u"没有发现"+e4p_file_path+u", 请将其放在当前目录.")
		i=raw_input("Press enter key to continue...")
		exit()
	else:
		print(u"发现"+e4p_file_path)
	print(u"检测零件所需配料列表("+i4e_file_path+u")文件...")
	if not os.path.exists(i4e_file_path):
		print(u"没有发现"+i4e_file_path+u", 请将其放在当前目录.")
		i=raw_input("Press enter key to continue...")
		exit()
	else:
		print(u"发现"+i4e_file_path)
	i=raw_input("Press enter key to continue...")
	print(u"读取Excel数据文件")
	p4m = p4m(p4m_file_path, e4p_file_path, i4e_file_path)
	print(u"读取成功,部门派单数据开始生成")
	p4m.get_department_data()
	print(u"派单数据生成成功,开始依规格排序")
	p4m.department_sort(d4s_file_path)
	print(u"排序完成,输出数据到Excel表")
	p4m.export_data()
	print(u"输出成功,请在当前目录output文件夹查看")
	i=raw_input("Press enter key to continue...")
