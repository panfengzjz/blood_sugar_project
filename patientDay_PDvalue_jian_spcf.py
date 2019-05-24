# revised 180130
# coding:utf-8
import xlrd
import xlwt
from xlutils.copy import copy
import time

class patient():
    def __init__(self, info):
        self.name = info[0]
        self.age = info[1]
        self.hos_area= info[4]
        self.bld_glu = info[7]
        self.opr_dat = info[8]
        self.opr_tim = info[13]

#def timepoint(p):
    #if (p.opr_tim.find("16点30") >= 0) or (p.opr_tim.find("10点30") >= 0) or (p.opr_tim.find("6点") >= 0) or (p.opr_tim.find("早餐前") >= 0):
        #timetype=1
    #elif (p.opr_tim.find("8点30") >= 0) or (p.opr_tim.find("13点") >= 0) or (p.opr_tim.find("19点") >= 0):
        #timetype=2
    #else:
        #timetype=3
    #return timetype

def print_excel(test_name, log_list, fileNumber=1):
    START_ROW=1 # 0 based (subtract 1 from excel row number)  
    if (fileNumber == 1):
        file_path = "./shulie.xls"
    elif (fileNumber == 2):
        file_path = "./bianliang.xls"
    elif (fileNumber == 3):
        file_path = "./patientday.xls"
    
    
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


def PDvalue_main_jian_spcf(filename, sheetname, startDate="", endDate="", district=[]):
    start_time = time.time()
    fd_excel = xlrd.open_workbook(filename) #打开文件
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet0
  
    n=0
    i=0  #写到第几行
    n=0  #每个pd里的血糖个数
    nr=0 #每个pd里达标的血糖个数
    nb=0 #每个pd里超标的血糖个数
    average_r=0  #均值达标的pd个数
    all_r=0  #所有值达标的pd个数
    single_nr=0 #至少一个bg不达标的pd个数
    
  
    sum0=0 #每个pd里血糖总和
   
    average0=0 #每个pd里平均血糖
   
    max_row = table.nrows
    pdaverage=[] #pd平均血糖一列
    low_perpd=0 #每个pd中低血糖次数
    pd_withlow=0 #有至少一次低血糖的pd个数
    pdn=[]  #每个pd中的血糖个数合集
   
    p0=patient(table.row_values(1)) 
    n=1    
    sum0=p0.bld_glu   
    if p0.bld_glu <=3.9 :
        low_perpd +=1       
    for r in range(2,max_row):
        p = patient(table.row_values(r))    
        former=patient(table.row_values(r-1))
        former_day=xlrd.xldate_as_tuple(former.opr_dat, 0)[0:3]
        
        p_day = xlrd.xldate_as_tuple(p.opr_dat, 0)[0:3]
        if (startDate != ""):
            start_date = tuple(map(int, startDate.split("-")))
            if (start_date > p_day):
                continue
        if (endDate != ""):
            end_date = tuple(map(int, endDate.split("-")))
            if (end_date < p_day):
                continue
        if (district != []) and (p.hos_area not in district):
            continue
        
        if (former.name==p.name)and (former_day==p_day):
            
                n+=1
                sum0+=p.bld_glu
                if p.bld_glu >=8 and p.bld_glu<=12:   #监护室控制范围
                    nr +=1
                elif p.bld_glu>12:
                    nb +=1                
                if p.bld_glu <=3.9 :
                    low_perpd +=1               
        else:
            if n!=0:
                average0=float(sum0)/float(n)
            if n!=0:
                
                if average0>=8 and average0<=12:
                    average_r +=1
                if n==nr:
                    all_r +=1
                if nb !=0:
                    single_nr +=1            
            if average0!=0:
                pdaverage.append(average0)
            if low_perpd!=0:
                pd_withlow +=1
                low_perpd=0
            pdn.append(n)
            n=1
            sum0=p.bld_glu
            average0=0
            if p.bld_glu <=3.9 :
                low_perpd +=1            
    last=patient(table.row_values(max_row-1))
    last_former=patient(table.row_values(max_row-2))
    last_former_day=xlrd.xldate_as_tuple(last_former.opr_dat, 0)[0:3]
    last_day = xlrd.xldate_as_tuple(last.opr_dat, 0)[0:3]    
    if (last_former.name!=last.name)or (last_former_day!=last_day):
        average0=last.bld_glu
        pdaverage.append(average0)
        
    
    
    print_excel("不分时刻pd总数", [len(pdaverage)],2)        
    print_excel("不分时刻每个pd", pdaverage,3)
    print_excel("不分时刻平均值达标数，每个值达标数，至少一个值不达标的pd数,pd总数", [average_r,all_r,single_nr,len(pdaverage)],2)
    print_excel("不分时刻pd至少一次低血糖的个数,pd总数", [pd_withlow,len(pdaverage)],2)  
    print_excel("每个pd里测量次数", pdn)
    
    print("the project costs: ", time.time()-start_time)
        
if __name__ == "__main__":
    PDvalue_main_jian_spcf("huge.xlsx", "jian", startDate="2018-07-01", endDate="2019-06-30",
                           district=['2病区', '4病区'])
