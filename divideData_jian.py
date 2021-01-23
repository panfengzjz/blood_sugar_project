#coding: utf-8
import time
import xlrd

from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils  
from xlrd import open_workbook #http://pypi.python.org/pypi/xlrd  
from xlwt import easyxf #http://pypi.python.org/pypi/xlwt  
from dayCount_print import print_excel

#a = [100,90,80,70]
#print_excel(a)

def divide_main_jian(filename, sheetname):
    fd_excel = xlrd.open_workbook(filename) #打开文件
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet
    n1=0
    n2=0
    n3=0
    n4=0
    n5=0
    n6=0
    n7=0
    
   
    row_len = table.col_values(0)
    #print_excel("分时段每种pd个数",[len(row_len)-1],2)
    for i in row_len[1:]:
        if i ==u"":
            break
        if i > 8:
            n1+=1
            if i > 10:
                n2+=1
                if i > 12:
                    n3+=1
                    if i > 14:
                        n4+=1
                        if i > 16:
                            n5+=1
                            if i > 18:
                                n6+=1
                                if i > 20:
                                    n7+=1
                                
    a=[n1,n2,n3,n4,n5,n6,n7]
    print( a)
    print_excel(">8-20",a,2)
    n1=0
    n2=0
    n3=0
    n4=0
    n5=0
    n6=0
    n7=0        

if __name__ == "__main__":
    divide_main_jian("patientday.xls", "Sheet1")
