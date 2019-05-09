#!/usr/bin/python
# -*- coding: utf-8 -*-
# 为了可以显示中文
#就是它！！计算全天平均＞ x 随住院天数变化  的非ICU

import xlrd
import time

from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils  
from xlrd import open_workbook #http://pypi.python.org/pypi/xlrd  
from xlwt import easyxf #http://pypi.python.org/pypi/xlwt

table = ""


def print_excel(test_name, log_list, fileNumber=1):
    START_ROW=1 # 0 based (subtract 1 from excel row number)  
    if (fileNumber == 1):
        file_path = "shulie.xls"
    elif (fileNumber == 2):
        file_path = "bianliang.xls"
    elif (fileNumber == 3):
        file_path = "patientday.xls"
    
    
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


class patient():
    def __init__(self, info):
        self.name = info[0]
        #self.age = info[1]
        #self.gender = info[2]
        #self.hos_num = info[3]
        #self.hos_area= info[4]
        #self.hos_bed = info[5]
        #self.pat_num = info[6]
        self.bld_glu = info[7]
        self.opr_dat = info[11]
        #self.opr_mtd = info[12]
        self.opr_tim = info[13]

def calculate_current_day(first_day, day_pass):
    temp_day = list(first_day)
    #print "temp_day is ", temp_day
    temp_day[2] = temp_day[2] + day_pass
    if(temp_day[1] == 4 or temp_day[1] == 6 or temp_day[1] == 9 or temp_day[1] == 11):
        if(temp_day[2] > 30):
            temp_day[2] -= 30
            temp_day[1] += 1
    elif(temp_day[1] == 1 or temp_day[1] == 3 or temp_day[1] == 5 or temp_day[1] == 7 or temp_day[1] == 8 or temp_day[1] == 10):
        if(temp_day[2] > 31):
            temp_day[2] -= 31
            temp_day[1] += 1
    
    elif(temp_day[1]==2):
        if(temp_day[0]%4==0):
            if(temp_day[2]>29):
                temp_day[2] -= 29
                temp_day[1] += 1  
        else:
            if(temp_day[2]>28):
                temp_day[2] -= 28
                temp_day[1] += 1     
    elif(temp_day[1]==12):
        if(temp_day[2] > 31):
            temp_day[2] -= 31
            temp_day[1] = 1        
            temp_day[0] += 1
    
    current_day = tuple(temp_day)
    return current_day

def find_pname_info_avg(p, current_day, j, max_row, denominator, numerator):
    sum = 0.0
    global table
    i = 0
    while(j < max_row):
        p_shadow = patient(table.row_values(j))
        if(p.name != p_shadow.name):
            break
        p_day = xlrd.xldate_as_tuple(p_shadow.opr_dat, 0)[0:3]
        if(p_day == current_day):
            sum += p_shadow.bld_glu
            i += 1
            j += 1  #需要遍历p病人current_day这天的全部记录，在这里不需要return
            continue
            
        elif(p_day < current_day):
            j += 1
            continue
        else:
            if(i != 0):
                denominator += 1
                if(sum/i >= 12):
                    numerator += 1
            break
    return (j, denominator, numerator)  #当病人改变，或已经大于current_day，就return
        
def jian_main(filename, sheetname):
    cur_time = time.time()
    fd_excel = xlrd.open_workbook(filename) #打开文件
    print( "Load excel time used is: ", time.time()-cur_time)
    global table
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet0
    max_row = table.nrows    
    array=[]
    arrayn=[]
    arrayd=[]
    #prev = patient(table.row_values(0))  #听从建议，把prev变量删除
    for i in range (0,6):  #最多采集7天的数据
        denominator = 0     #分母为总样本数
        numerator = 0       #分子为不达标样本数
        j = 0
        patient_number=0
        p = patient(table.row_values(0))
        while (j < max_row-1):
            j += 1
            former = p  #避免一次次重复加载，将前一次的p赋给former
            #print "j equals ", j
            p = patient(table.row_values(j))
            #former = patient(table.row_values(j-1))

            if(p.name != former.name):
                patient_number+=1   #寻找新的病人
                first_day = xlrd.xldate_as_tuple(p.opr_dat, 0)[0:3] #只允许加载一次第一天
                current_day = calculate_current_day(first_day, i)
                #print "The", i, "loop: first day is ", first_day
                #print current_day
            else:
                continue    #如果上下两条名字相同，则跳过
            
            (j, denominator, numerator) = find_pname_info_avg(p, current_day, j, max_row, denominator, numerator)
            j -= 1          #考虑边界，需要-1
            
        #print "The %d day: denominator is %d, numerator is %d" %(i+1, denominator, numerator) 
        #print float(numerator), float(denominator)
        if (denominator!=0):
            array.append(float(numerator)/float(denominator))
            arrayn.append(numerator)
            arrayd.append(denominator)
        print( "total time used is: ", time.time()-cur_time)
    #print "We have %d patient totally" %patient_number
    print_excel("rate",array,2)
    print_excel("分子",arrayn,2)
    print_excel("分母",arrayd,2)

if __name__ == "__main__":
    jian_main("jian.xlsx", "Sheet")