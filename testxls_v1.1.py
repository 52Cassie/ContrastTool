#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd 
import math
import string 
import sys

def read_xls_file(file):
	data = xlrd.open_workbook(file)
	table = data.sheets()[0]
	return table

def find_cols_rows(data):
	nrows = data.nrows
	ncols = data.ncols
	flag = 0
	#r,net,x,y,nrows
	r,net,x,y = 0,0,0,0
	for i in xrange(0,nrows):
		rowValues = data.row_values(i)
		for j in xrange(0,ncols):
			if type(rowValues[j])== unicode :
				if str.upper(str(rowValues[j].encode('utf-8')))=="NET" :
					r = i+1
					print r
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
	print 'find_cols_rows',r,net,x,y,nrows
	return r,net,x,y,nrows

def comparison(file_new,file_old,result):
	data_old = read_xls_file(file_old)
	data_new = read_xls_file(file_new)
	f = open(result,'w')
	r_new,net_new,x_new,y_new,nrows_new = find_cols_rows(data_new)
	r_old,net_old,x_old,y_old,nrows_old = find_cols_rows(data_old)
	print "comparison ",r_new,net_new,x_new,y_new,nrows_new

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
				s = rowValues_new[3]+", x,"+str(rowValues_new[1])+", y,"+str(rowValues_new[2])+" , "+rowValues_old[7]+", x,"+str(rowValues_old[1])+", y, "+str(rowValues_old[2])+", x,y exchange\n"
			elif rowValues_new[1]!=rowValues_old[1]:
				s = rowValues_new[3]+", x,"+str(rowValues_new[1])+", y,"+str(rowValues_new[2])+" , "+rowValues_old[7]+", x,"+str(rowValues_old[1])+", y, "+str(rowValues_old[2])+", x exchange\n"
			elif rowValues_new[2]!=rowValues_old[2]:
				s = rowValues_new[3]+", x,"+str(rowValues_new[1])+", y,"+str(rowValues_new[2])+" , "+rowValues_old[7]+", x,"+str(rowValues_old[1])+", y,"+str(rowValues_old[2])+", y exchange\n"
			else :
				s = rowValues_new[3]+", x,"+str(rowValues_new[1])+", y,"+str(rowValues_new[2])+" , "+rowValues_old[7]+", x,"+str(rowValues_old[1])+", y,"+str(rowValues_old[2])+", yes\n"
		else :
				s = rowValues_new[3]+", x,"+str(rowValues_new[1])+", y,"+str(rowValues_new[2])+", null\n"
		f.write(s)
		minn = 99999
		temp = 0
	f.close()

if __name__ == '__main__':
	comparison(sys.argv[1],sys.argv[2],sys.argv[3])

# file_old = "/Users/mac/Desktop/test/Nail.xlsx"
# file_new = "/Users/mac/Desktop/test/data_new.xlsx"
# result = "/Users/mac/Desktop/test/jjj.txt"
# comparison(file_new,file_old,result)