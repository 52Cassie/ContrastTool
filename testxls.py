#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xlwt
from xlutils.copy import copy
import math
import string 
import sys
import re

def read_xls_file(file):
	data = xlrd.open_workbook(file)
	table = data.sheets()[0]
	return table

def find_cols_rows(data):
	nrows = data.nrows
	ncols = data.ncols
	flag = 0
	r,net,x,y = 0,0,0,0
	flagname = [0]*8
	for i in xrange(0,nrows):
		rowValues = data.row_values(i)
		for j in xrange(0,ncols):
			if type(rowValues[j])== unicode :
				if str.upper(str(rowValues[j].encode('utf-8')))=="LOCATION" :
					flagname[0] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="NAIL" :
					flagname[1] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="X" :
					flagname[2] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="Y" :
					flagname[3] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="NET" :
					r = i+1
					flagname[4] = j
					flag = flag+2
				elif str.upper(str(rowValues[j].encode('utf-8')))=="T/B" :
					flagname[5] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="VIRTUAL" :
					flagname[6] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="PIN/VIA" :
					flagname[7] = j
					flag = flag+1
		# if flag == 4 :
		# 	break
	return r,nrows,flagname

def create_xls_file(result,title):
	workbook = xlwt.Workbook(encoding = 'utf-8')
	worksheet = workbook.add_sheet('My Workbook')
	row0 = [title,u'变化','Nail']
	row = ['Location','X','Y', 'Net','Virtual','Pin/Via', ' ','Nail','X','Y','Net','T/B','Virtual','Pin/Via']

	style = xlwt.XFStyle() 

	font = xlwt.Font() 
	font.name = 'Heiti SC Light'
	font.bold = True 
	style.font = font 

	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER 
	alignment.vert = xlwt.Alignment.VERT_CENTER 
	style.alignment = alignment

	worksheet.write_merge(0,0,0,5,row0[0],style) #(x1,x2,y1,y2)
	worksheet.write(0,6,row0[1],style)
	worksheet.write_merge(0,0,7,13,row0[2],style)

	for i in xrange(len(row)):
		worksheet.write(1,i,row[i],style)
		worksheet.col(i).width = 256*20

	tall_style = xlwt.easyxf('font:height 360;')
	worksheet.row(0).set_style(tall_style)
	workbook.save(result)

def write_xls_file(result,row):
	readbook = xlrd.open_workbook(result)
	data = readbook.sheets()[0]
	nrows = data.nrows
	workbook = copy(readbook)
	worksheet = workbook.get_sheet(0)
	for i in xrange(len(row)):
		worksheet.write(nrows,i,row[i])
	workbook.save(result)

def create_txt_file(result,title):
	f = open(result,'w')
	row0 = ["title",'变化','Nail']
	row = ['Location','X','Y', 'Net','Virtual','Pin/Via', ' ','Nail','X','Y','Net','T/B','Virtual','Pin/Via']
	write_txt_file(result,row0)
	write_txt_file(result,row)
	f.close()

def write_txt_file(result,row):
	f = open(result,'a+')
	s = ""
	for i in xrange(len(row)-1):
		s = s + row[i] + ","
	s = s + row[len(row)-1] + "\n"
	f.write(s)
	f.close()

def merge_cell(sheet):
    rt = {}
    if sheet.merged_cells:
        # exists merged cell
        for item in sheet.merged_cells:
            for row in range(item[0], item[1]):
                for col in range(item[2], item[3]):
                    rt.update({(row, col): (item[0], item[2])})
    return rt

def get_merged(filename):
    book = xlrd.open_workbook(filename)
    sheet = book.sheets()[0]    
    # 获取合并的单元格
    merged = merge_cell(sheet)
    # 获取sheet的行数（默认每一行就是一条用例）
    rows = sheet.nrows
    # 如果sheet为空，那么rows是0
    if rows:
        for row in range(rows):
            data = sheet.row_values(row)   # 单行数据
            for index, content in enumerate(data):
                if merged.get((row, index)):
                    # 这是合并后的单元格，需要重新取一次数据
                    data[index] = sheet.cell_value(*merged.get((row, index)))
    return data

def comparison(file_new,file_old,result):
	data_old = read_xls_file(file_old)
	data_new = read_xls_file(file_new)
	title = get_merged(file_new)
	print title[0]

	matchObj = re.match( r'.*\.(.*)', result, re.I)
	# print matchObj.group(1)
	if matchObj.group(1) == 'xls' :
		create_xls_file(result,title[0])
	elif matchObj.group(1) == 'txt' :
		create_txt_file(result,title[0])

	r_new,nrows_new,name_new = find_cols_rows(data_new)
	r_old,nrows_old,name_old = find_cols_rows(data_old)
	#name_new = [location,nail,x,y,net,t/b,virtual,pin/via]
	minn = 99999
	flag_list = [0]*nrows_old 
	temp = 0
	for i in xrange(r_new,nrows_new):
		rowValues_new = data_new.row_values(i)
		for j in xrange(r_old,nrows_old):
			rowValues_old = data_old.row_values(j)
			if flag_list[j]==0 and rowValues_new[name_new[4]]==rowValues_old[name_old[4]]:
				xd = rowValues_new[name_new[2]]-rowValues_old[name_old[2]]
				yd = rowValues_new[name_new[3]]-rowValues_old[name_old[3]]
				d = math.sqrt(xd**2+yd**2)
				if d < minn:
					minn = d
					temp = j
		if temp != 0:
			rowValues_old = data_old.row_values(temp)
			flag_list[temp] = 1 
			if rowValues_new[name_new[2]]!=rowValues_old[name_old[2]] and rowValues_new[name_new[3]]!=rowValues_old[name_old[3]]:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'x,y change',str(rowValues_old[name_old[1]]),str(rowValues_old[name_old[2]]),str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			elif rowValues_new[name_new[2]]!=rowValues_old[name_old[2]]:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'x change',str(rowValues_old[name_old[1]]),str(rowValues_old[name_old[2]]),str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			elif rowValues_new[name_new[3]]!=rowValues_old[name_old[3]]:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'y change',str(rowValues_old[name_old[1]]),str(rowValues_old[name_old[2]]),str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			else :
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),' ',str(rowValues_old[name_old[1]]),str(rowValues_old[name_old[2]]),str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
		else :
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),' ']
		#print s
		if matchObj.group(1) == 'xls' :
			write_xls_file(result,s)
		elif matchObj.group(1) == 'txt' :
			write_txt_file(result,s)
		
		minn = 99999
		temp = 0

if __name__ == '__main__':
	comparison(sys.argv[1],sys.argv[2],sys.argv[3])

# file_old = "/Users/mac/Desktop/test/S-1_nail.xlsx"
# file_new = "/Users/mac/Desktop/test/s-1.xlsx"
# result = "/Users/mac/Desktop/test/jjj.txt"
# comparison(file_new,file_old,result)

# create_xls_file(result,"IN680-F(820-01365-01-07)P2 TPs_20180122(US Dry run)")
#write_xls_file(result,['gg'])

