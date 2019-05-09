#coding: utf-8
import time
import xlrd

from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils  
from xlrd import open_workbook #http://pypi.python.org/pypi/xlrd  
from xlwt import easyxf #http://pypi.python.org/pypi/xlwt  


def print_excel(test_name, log_list, fileNumber=1):
    START_ROW=1 # 0 based (subtract 1 from excel row number)  
    if (fileNumber == 1):
        file_path = "shulie.xls"
    elif (fileNumber == 2):
        file_path = "bianliang.xls"
    
    
    rb=xlrd.open_workbook(file_path)  
    r_sheet=rb.sheet_by_index(0)# read only copy to introspect the file  
    wb=copy(rb)# a writable copy (I can't read values out of this, only write to it)  
    w_sheet=wb.get_sheet(0)# the sheet to write to within the writable copy
    
    w_sheet.write(0, r_sheet.ncols, test_name)
    for row_index in range(0, len(log_list)):
        w_sheet.write(row_index+1, r_sheet.ncols, log_list[row_index])
    
    wb.save(file_path)
    #wb.save(file_path+'.out'+ os.path.splitext(file_path)[-1])
    print( "END")

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
