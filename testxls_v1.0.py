import xlrd 
import math
import string 
import sys
 
def read_xls_file_old(file_old):
	data_old = xlrd.open_workbook(file_old)
	table_old = data_old.sheets()[0]
	return table_old		

def read_xls_file_new(file_new):
	data_new = xlrd.open_workbook(file_new)
	table_new = data_new.sheets()[0]
	return table_new

def comparison(file_new,file_old,result):
	data_old = read_xls_file_old(file_old)
	data_new = read_xls_file_new(file_new)
	f = open(result,'w')
	nrows_new = data_new.nrows
	nrows_old = data_old.nrows
	ncols_new = data_new.ncols
	ncols_old = data_old.ncols

	#find file_new net/x/y cols
	flag = 0
	for i in xrange(0,nrows_new):
		rowValues_new = data_new.row_values(i)
		for j in xrange(0,ncols_new):
			if str.upper(str(rowValues_new[j]))=="NET" :
				r_new = i+1
				net_new = j
				flag = flag+2
			elif str.upper(str(rowValues_new[j]))=="X" :
				x_new = j
				flag = flag+1
			elif str.upper(str(rowValues_new[j]))=="Y" :
				y_new = j
				flag = flag+1
		if flag == 4 :
			break

	#find file_old net/x/y cols
	flag = 0
	for i in xrange(0,nrows_old):
		rowValues_old = data_old.row_values(i)
		for j in xrange(0,ncols_old):
			if str.upper(str(rowValues_old[j]))=="NET" :
				r_old = i+1
				net_old = j
				flag = flag+2
			elif str.upper(str(rowValues_old[j]))=="X" :
				x_old = j
				flag = flag+1
			elif str.upper(str(rowValues_old[j]))=="Y" :
				y_old = j
				flag = flag+1
		if flag == 4 :
			break


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
# file_new = "/Users/mac/Desktop/test/s-1.xlsx"
# result = "/Users/mac/Desktop/test"
# comparison(file_new,file_old,result)