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
	for i in xrange(0,nrows):
		rowValues = data.row_values(i)
		for j in xrange(0,ncols):
			if type(rowValues[j])== unicode :
				if str.upper(str(rowValues[j].encode('utf-8')))=="NET" :
					r = i+1
					net = j
					flag = flag+2
				elif str.upper(str(rowValues[j].encode('utf-8')))=="X" :
					x = j
					flag = flag+1
				elif str.upper(str(rowValues[j].encode('utf-8')))=="Y" :
					y = j
					flag = flag+1
		if flag == 4 :
			break
	return r,net,x,y,nrows

def create_xls_file(result):
	row = [u'针点表','x','y', u'总表','x','y', u'是否发生变化']
	workbook = xlwt.Workbook(encoding = 'utf-8')
	worksheet = workbook.add_sheet('My Workbook')
	style = xlwt.XFStyle() 
	font = xlwt.Font() 
	font.name = 'Heiti SC Light' 
	font.bold = True 
	style.font = font 
	for i in xrange(len(row)):
		worksheet.write(1,i,row[i],style)
	workbook.save(result)

def create_txt_file(result):
	f = open(result,'w')
	row = "针点表,x,y,总表,x,y,是否发生变化\n"
	f.write(row)
	f.close()

def write_txt_file(result,row):
	f = open(result,'a+')
	s = ""
	for i in xrange(len(row)-2):
		s = s + row[i] + ","
	s = s + row[len(row)-1] + "\n"
	f.write(s)
	f.close()

def write_xls_file(result,row):
	readbook = xlrd.open_workbook(result)
	data = readbook.sheets()[0]
	nrows = data.nrows
	workbook = copy(readbook)
	worksheet = workbook.get_sheet(0)
	for i in xrange(len(row)):
		worksheet.write(nrows,i,row[i])
	workbook.save(result)

def comparison(file_new,file_old,result):
	data_old = read_xls_file(file_old)
	data_new = read_xls_file(file_new)

	matchObj = re.match( r'.*\.(.*)', result, re.I)
	print matchObj.group(1)
	if matchObj.group(1) == 'xls' :
		create_xls_file(result)
	elif matchObj.group(1) == 'txt' :
		create_txt_file(result)

	r_new,net_new,x_new,y_new,nrows_new = find_cols_rows(data_new)
	r_old,net_old,x_old,y_old,nrows_old = find_cols_rows(data_old)
	#print r_new,net_new,x_new,y_new,nrows_new
	minn = 99999
	flag_list = [0]*nrows_old 
	temp = 0
	for i in xrange(r_new,nrows_new):
		rowValues_new = data_new.row_values(i)
		for j in xrange(r_old,nrows_old):
			rowValues_old = data_old.row_values(j)
			if flag_list[j]==0 and rowValues_new[net_new]==rowValues_old[net_old]:
				xd = rowValues_new[x_new]-rowValues_old[x_old]
				yd = rowValues_new[y_new]-rowValues_old[y_old]
				d = math.sqrt(xd**2+yd**2)
				if d < minn:
					minn = d
					temp = j
		if temp != 0:
			rowValues_old = data_old.row_values(temp)
			flag_list[temp] = 1 
			if rowValues_new[1]!=rowValues_old[1] and rowValues_new[2]!=rowValues_old[2]:
				s = [rowValues_new[3],str(rowValues_new[1]),str(rowValues_new[2]),rowValues_old[7],str(rowValues_old[1]),str(rowValues_old[2]),'x,y change']
			elif rowValues_new[1]!=rowValues_old[1]:
				s = [rowValues_new[3],str(rowValues_new[1]),str(rowValues_new[2]),rowValues_old[7],str(rowValues_old[1]),str(rowValues_old[2]),'x change']
			elif rowValues_new[2]!=rowValues_old[2]:
				s = [rowValues_new[3],str(rowValues_new[1]),str(rowValues_new[2]),rowValues_old[7],str(rowValues_old[1]),str(rowValues_old[2]),'y change']
			else :
				s = [rowValues_new[3],str(rowValues_new[1]),str(rowValues_new[2]),rowValues_old[7],str(rowValues_old[1]),str(rowValues_old[2]),'no change']
		else :
				s = [rowValues_new[3],str(rowValues_new[1]),str(rowValues_new[2]),'null']
		#print s
		if matchObj.group(1) == 'xls' :
			write_xls_file(result,s)
		elif matchObj.group(1) == 'txt' :
			write_txt_file(result,s)
		
		minn = 99999
		temp = 0

if __name__ == '__main__':
	comparison(sys.argv[1],sys.argv[2],sys.argv[3])

#file_old = "C:\Users\Cassie\Desktop\S-1_nail.xlsx"
#file_new = "C:\Users\Cassie\Desktop\s-1.xlsx"
#result = "C:\Users\Cassie\Desktop\jjj.xls"
#comparison(file_new,file_old,result)

# create_xls_file(result)
#write_xls_file(result,['gg'])

