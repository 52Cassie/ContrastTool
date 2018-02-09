#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xlwt
import string 
import sys
import re
import os

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
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="NAIL" :
					flagname[1] = j
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="X" :
					flagname[2] = j
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="Y" :
					flagname[3] = j
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="NET" :
					r = i+1
					flagname[4] = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="T/B" :
					flagname[5] = j
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="VIRTUAL" :
					flagname[6] = j
					# flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="PIN/VIA" :
					flagname[7] = j
					# flag = flag+1
		if flag == 1 :
			break
	return r,nrows,flagname

def txt_to_xls_file(txtpath):
	xlspath = os.path.join(os.path.dirname(txtpath),   
        	os.path.splitext(os.path.basename(txtpath))[0] + '.xls')
	workbook = xlwt.Workbook(encoding = 'utf-8')
	worksheet = workbook.add_sheet('My Workbook')
	
	style = xlwt.XFStyle() 
	font = xlwt.Font() 
	font.name = 'Heiti SC Light'
	font.bold = True 
	style.font = font 

	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER 
	alignment.vert = xlwt.Alignment.VERT_CENTER 
	style.alignment = alignment

	BUFSIZE = 1024
	EXCEL_ROWS = 65535
	EXCEL_COLS = 256
	FIELD_SEPARATOR = ','
	title = 0
	with open(txtpath,'r') as f:
		nrows = 0
		lines = f.readlines(BUFSIZE)
		while lines:
			for line in lines:
				values = line.split(FIELD_SEPARATOR)
				cols_num = EXCEL_COLS if len(values) > EXCEL_COLS else len(values)
				if title == 0:
					worksheet.write_merge(0,0,0,5,values[0],style) #(x1,x2,y1,y2)
					worksheet.write(0,6,values[1],style)
					worksheet.write_merge(0,0,7,13,values[2],style)
					title = 1
				else :
					for ncol in xrange(cols_num):
						worksheet.write(nrows,ncol,values[ncol],style)
				nrows = nrows + 1
			lines = f.readlines(BUFSIZE)
	
	tall_style = xlwt.easyxf('font:height 360;')
	worksheet.row(0).set_style(tall_style)
	workbook.save(xlspath)

def create_txt_file(result,title):
	f = open(result,'w')
	row0 = [title,'变化','Nail']
	row = ['Location','X','Y', 'Net','Virtual','Pin/Via', ' ','Nail','X','Y','Net','T/B','Virtual','Pin/Via']
	write_txt_file(result,row0)
	write_txt_file(result,row)
	f.close()

def write_txt_file(result,row):
	f = open(result,'a+')
	s = ""
	for i in xrange(len(row)-1):
		s = s + str(row[i]) + ","
	s = s + str(row[len(row)-1]) + "\n"
	f.write(s)
	f.close()

def comparison(file_new,file_old,result):
	data_old = read_xls_file(file_old)
	data_new = read_xls_file(file_new)
	merge = []
	for (rlow,rhigh,clow,chigh) in data_new.merged_cells:
		merge.append([rlow,clow])
	title = data_new.cell_value(merge[0][0],merge[0][1])

	matchObj = re.match( r'.*\.(.*)', result, re.I)
	# print matchObj.group(1)
	if matchObj.group(1) == 'xls' :
		result = os.path.join(os.path.dirname(result),   
        	os.path.splitext(os.path.basename(result))[0] + '.txt')
	# print result
	create_txt_file(result,title)

	r_new,nrows_new,name_new = find_cols_rows(data_new)
	r_old,nrows_old,name_old = find_cols_rows(data_old)
	#name_new = [location,nail,x,y,net,t/b,virtual,pin/via]
	flag_list = [0]*nrows_old 
	for i in xrange(r_new,nrows_new):
		rowValues_new = data_new.row_values(i)
		flag = 4
		temp = 0
		for j in xrange(r_old,nrows_old):
			rowValues_old = data_old.row_values(j)
			if flag_list[j]==0 and rowValues_new[name_new[4]]==rowValues_old[name_old[4]]:
				if rowValues_new[name_new[2]] == rowValues_old[name_old[2]] :
					if rowValues_new[name_new[3]] == rowValues_old[name_old[3]] :
						flag = 0 # no change
						temp = j
						break
					else :
						if flag > 1:
							flag = 1 # y change
							temp = j
				else :
					if rowValues_new[name_new[3]] == rowValues_old[name_old[3]] :
						if flag > 2:
							flag = 2 # x change
							temp = j
					else :
						if flag > 3:
							flag = 3 # x,y change
							temp = j
		if temp != 0:
			rowValues_old = data_old.row_values(temp)
			if flag == 3:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'x,y change',' ',' ',' ',str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			elif flag == 2:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'x change',' ',' ',str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			elif flag == 1:
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),'y change',' ',str(rowValues_old[name_old[2]]),' ',str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
			elif flag == 0 :
				flag_list[temp] = 1 
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),' ',int(rowValues_old[name_old[1]]),str(rowValues_old[name_old[2]]),str(rowValues_old[name_old[3]]),str(rowValues_old[name_old[4]]),str(rowValues_old[name_old[5]]),str(rowValues_old[name_old[6]]),str(rowValues_old[name_old[7]])]
		else :
				s = [str(rowValues_new[name_new[0]]),str(rowValues_new[name_new[2]]),str(rowValues_new[name_new[3]]),str(rowValues_new[name_new[4]]),str(rowValues_new[name_new[6]]),str(rowValues_new[name_new[7]]),' ']
		#print s
		write_txt_file(result,s)
	print result
	if matchObj.group(1) == 'xls' :
		txt_to_xls_file(result)
		

if __name__ == '__main__':
	comparison(sys.argv[1],sys.argv[2],sys.argv[3])

# file_old = "/Users/mac/Desktop/test/Nail.xlsx"
# file_new = "/Users/mac/Desktop/test/data_new.xlsx"
# result = "/Users/mac/Desktop/test/aa.xls"
# comparison(file_new,file_old,result)

# create_xls_file(result,"IN680-F(820-01365-01-07)P2 TPs_20180122(US Dry run)")
#write_xls_file(result,['gg'])

