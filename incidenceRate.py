#coding: utf-8
#coding: utf-8
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils  
from dayCount_print import print_excel

class patient():
    def __init__(self, info):
        self.name = info[0]
        self.age = info[1]
        self.hos_area= info[4]
        self.bld_glu = info[7]
        self.opr_dat = info[8]
        self.opr_tim = info[13]
def fuce(s1,s2):
    h1=s1[0]
    h2=s2[0]
    m1=s1[1]
    m2=s2[1]
    if h1==h2:
        if m2-m1<=20 and  m2-m1>=10:
            return True
        else:
            return False
    elif h2-h1==1:
        if m2+60-m1<=20 and  m2+60-m1>=10:
            return True
        else:
            return False
    else:
        return False

def is_night(excel_time):
    hour = excel_time[0]
    if (hour>=22 or hour<5):
        return True
    else:
        return False

def rate_main(filename, sheetname, startDate="", endDate="", district=[]):
    fd_excel = xlrd.open_workbook(filename) #打开文件
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet0
    max_row = table.nrows  
    before=0
    before_high=0
    before_low=0
    after=0
    after_high=0
    after_low=0
    night=0
    night_high=0
    night_low=0
    latenight=0
    latenight_low=0
    other=0
    other_low=0
    total_low=0
    again=0
    againsuccess=0
    mmp=0  #同一人低血糖多测次数
    for r in range(1,max_row):
        p = patient(table.row_values(r))
       
        if p.bld_glu<=3.9:
            
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
            #if (former.name==p.name)and (former_day==p_day)and(former.opr_tim==p.opr_tim):
                #pass
            #else:
                #print "和前一行不一样",p.name
            if former.bld_glu>3.9:   
                total_low +=1            
                flag=0
                for s in range(r+1,min(max_row,r+6)):
                    p_shadow=patient(table.row_values(s))
                    #print "开始往下数",p_shadow.bld_glu
                    #print "此时P=",p.bld_glu
                    p_day = xlrd.xldate_as_tuple(p.opr_dat, 0)[0:3]
                    p_shadow_day = xlrd.xldate_as_tuple(p_shadow.opr_dat, 0)[0:3]
                    p_min=xlrd.xldate_as_tuple(p.opr_dat, 0)[3:5]
                    p_shadow_min=xlrd.xldate_as_tuple(p_shadow.opr_dat, 0)[3:5]
                    if (p_shadow.name==p.name) and (p_day==p_shadow_day):
                        #print "和p一样shadow",p_shadow.name
                        p_shadow_former_min=xlrd.xldate_as_tuple(patient(table.row_values(s-1)).opr_dat, 0)[3:5]
                        if fuce(p_shadow_former_min, p_shadow_min)==True:
                            flag+=1
                        #print "flag",flag
                        if p_shadow.bld_glu<=3.9:
                            #print "<3.9",p_shadow.bld_glu
                            mmp +=1     
                            #print "mmp",mmp
                        #else:
                if flag >0:
                    again +=1            
                          
                    #else:
                        #if flag>0:
                            #again +=1  
                            ##print "again",again
                            #if patient(table.row_values(s-1)).bld_glu>3.9:
                                #againsuccess +=1    
                               
                
        
                if (p.opr_tim.find("16点30") >= 0) or (p.opr_tim.find("10点30") >= 0) or (p.opr_tim.find("6点") >= 0) or (p.opr_tim.find("早餐前") >= 0):
                    before +=1
                    if p.bld_glu>8:
                        before_high +=1
                    elif p.bld_glu<=3.9:
                        before_low +=1                
                elif (p.opr_tim.find("8点30") >= 0) or (p.opr_tim.find("13点") >= 0) or (p.opr_tim.find("19点") >= 0):
                    after +=1
                    if p.bld_glu>10:
                        after_high +=1 
                    elif p.bld_glu<=3.9:
                        after_low +=1                
                elif (p.opr_tim.find("21点") >= 0): 
                    night +=1
                    if p.bld_glu>10:
                        night_high +=1     
                    elif p.bld_glu<=3.9:
                        night_low +=1      
                elif is_night(xlrd.xldate_as_tuple(p.opr_dat, 0)[3:5])==True:
                    latenight +=1
                    if p.bld_glu<=3.9:
                        latenight_low +=1                         
                else:     
                    other +=1
                    if p.bld_glu<=3.9:
                        other_low +=1                        
     ###           
        
    print_excel("低血糖复测人次", [again],2)
    #print_excel("复测成功次数", [againsuccess],2)
    print_excel("同一人低血糖多测次数",[mmp],2)
    print_excel("低血糖复测率",[float(again)/(total_low)],2)
    #print_excel("single餐前高血糖发生率",[before_high,before,float(before_high)/before],2)
    #print_excel("single餐后高血糖发生率",[after_high,after,float(after_high)/after],2)
    #print_excel("single睡前高血糖发生率",[night_high,night,float(night_high)/night],2)
    print_excel("记录总条数",[max_row-1],2)
    print_excel("低血糖总条数",[total_low],2)
    print_excel("低血糖发生率",[float(total_low/(max_row-1-mmp))],2)
    print_excel("餐前低血糖总条数",[before_low],2)
    print_excel("餐后低血糖总条数",[after_low],2)
    print_excel("睡前低血糖总条数",[night_low],2)
    print_excel("半夜低血糖总条数",[latenight_low],2)
    print_excel("其它低血糖总条数",[other_low],2)
    
    
    #print( "低血糖复测人次",again,"复测成功次数",againsuccess,"同一人低血糖多测次数",mmp)
    #print( "低血糖复测率",float(again)/(total_low-mmp))
   
    #print( "single餐前高血糖发生率", float(before_high)/before,before_high,"/",before)
    #print( "single餐后高血糖发生率", float(after_high)/after,after_high,"/",after)
    #print( "single睡前高血糖发生率", float(night_high)/night,night_high,"/",night)
    
    #print( "记录总条数",max_row-1)
    #print( "低血糖总数%d,发生在餐前%d，发生在餐后%d，发生在睡前%d"%(total_low,before_low,after_low,night_low))
    
if __name__ == "__main__":
    rate_main("wai-sample.xlsx", "Sheet1", startDate="2018-07-01", endDate="2019-06-30", 
              district=['11病区', '17病区'])
    
    
 