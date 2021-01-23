# revised 180130
# coding:utf-8
import xlrd
import xlwt
from xlutils.copy import copy
import time

from dayCount_print import print_excel

class patient():
    def __init__(self, info):
        self.name = info[0]
        self.age = info[1]
        self.hos_area= info[4]
        self.bld_glu = info[7]
        self.opr_dat = info[8]
        self.opr_tim = info[13]

def timepoint(p):
    if (p.opr_tim.find("16点30") >= 0) or (p.opr_tim.find("10点30") >= 0) or (p.opr_tim.find("6点") >= 0) or (p.opr_tim.find("早餐前") >= 0):
        timetype=1
    elif (p.opr_tim.find("8点30") >= 0) or (p.opr_tim.find("13点") >= 0) or (p.opr_tim.find("19点") >= 0):
        timetype=2
    else:
        timetype=3
    return timetype

def PDvalue_main(filename, sheetname, startDate="", endDate="", district=[]):
    start_time = time.time()
    fd_excel = xlrd.open_workbook(filename) #打开文件
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet0
  
    n=0
    i=0  #写到第几行
    n1=0
    n2=0
    n3=0
    n1r=0  #每个pd里的bg数里有几个达标的
    n2r=0
    n3r=0
    n1b=0   #每个pd里的bg数里有几个超过上限的
    n2b=0
    n3b=0
    sum1=0
    sum2=0
    sum3=0
    average1=0
    average2=0
    average3=0
    max_row = table.nrows
    premeal=[] #pd平均血糖一列
    postmeal=[]
    presleep=[]
    mpre=0   #测过餐前的pd个数
    mpost=0
    mall=0  #同时测了餐前和餐后的pd个数
    mbed=0
    rpre=0  #平均餐前达标的pd个数
    rpost=0
    rall=0  
    rapre=0 #每个bg都达标的pd个数
    rapost=0
    raall=0
    nrpre=0   #至少一个bg不达标的pd个数
    nrpost=0
    nrbed=0
    pdn1=[]  #每个pd中的餐前个数合集
    pdn2=[]
    pdn3=[]
    p0=patient(table.row_values(1)) 
    if timepoint(p0)==1:
        n1+=1
        sum1+=p0.bld_glu
    elif timepoint(p0)==2:
        n2+=1
        sum2+=p0.bld_glu
    elif timepoint(p0)==3:
        n3+=1
        sum3+=p0.bld_glu         
        
    
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
        if (district != []) and (p.hos_area in district):
            continue
        
        if (former.name==p.name)and (former_day==p_day):
            if timepoint(p)==1:
                n1+=1
                sum1+=p.bld_glu
                if p.bld_glu >=4.4 and p.bld_glu<=8:
                    n1r +=1
                elif p.bld_glu>8:
                    n1b +=1
               
            elif timepoint(p)==2:
                n2+=1
                sum2+=p.bld_glu
                if p.bld_glu >=6 and p.bld_glu<=10:
                    n2r +=1     
                elif p.bld_glu>10:
                    n2b +=1                
            elif timepoint(p)==3:
                n3+=1
                if p.bld_glu>8:
                    n3b +=1
                sum3+=p.bld_glu
        else:
            if n1!=0:
                average1=float(sum1)/float(n1)
            if n2!=0:
                average2=float(sum2)/float(n2)
            if n3!=0:
                average3=float(sum3)/float(n3)
            
            if n1!=0:
                mpre +=1
                if average1>=4.4 and average1<=8:
                    rpre +=1
                if n1==n1r:
                    rapre +=1
                if n1b !=0:
                    nrpre +=1
            if n2!=0:
                mpost +=1
                if average2>=6 and average2<=10:
                    rpost +=1
                if n2==n2r:
                    rapost +=1    
                if n2b !=0:
                    nrpost +=1
            if n3!=0:
                mbed +=1
                if n3b !=0:
                    nrbed +=1                
    
            if n1!=0 and n2!=0:
                mall +=1
                if average1>=4.4 and average1<=7.8 and average2>=6.1 and average2<=10:
                #if average1>=4.4 and average1<=10 and average2>=6.1 and average2<=13.9:
                    rall +=1
                if n1==n1r and n2==n2r:
                    raall +=1                
            
            if average1!=0:
                premeal.append(average1)
            if average2!=0:
                postmeal.append(average2)
            if average3!=0:
                presleep.append(average3)
                
            
            
            pdn1.append(n1)
            pdn2.append(n2)
            pdn3.append(n3)
            if timepoint(p)==1:
                n1=1
                sum1=p.bld_glu
                if p.bld_glu >=4.4 and p.bld_glu<=8:
                    n1r =1
                elif p.bld_glu>8:
                    n1r=0
                    n1b=1
                else:
                    n1r=0
                    n1b=0                    
                n2=0 
                n3=0
                sum2=0
                sum3=0
            elif timepoint(p)==2:
                
                n2=1
                sum2=p.bld_glu
                if p.bld_glu >=6 and p.bld_glu<=10:
                    n2r =1
                elif p.bld_glu>10:
                    n2r=0
                    n2b=1
                else:
                    n2r=0
                    n2b=0             
                n1=0 
                n3=0
                sum1=0
                sum3=0                
            elif timepoint(p)==3:
                
                n3=1
                sum3=p.bld_glu
                if p.bld_glu>8:
                    n3b=1
                else:
                    n3b=0
                n2=0 
                n1=0
                sum2=0
                sum1=0  
            average1=0
            average2=0
            average3=0    

    pdn1.append(n1)
    pdn2.append(n2)
    pdn3.append(n3)
    
    print_excel("premeal每个pd", premeal,3)
    print_excel("postmeal每个pd", postmeal,3)
    print_excel("presleep每个pd", presleep,3)
    
    print_excel("premeal pd总数", [len(premeal)],2)
    print_excel("postmeal pd总数", [len(postmeal)],2)
    print_excel("presleep pd总数", [len(presleep)],2)
    
    print_excel("premeal平均值达标数，每个值达标数，测过餐前的pd数", [rpre,rapre,mpre],2)
    print_excel("postmeal平均值达标数，每个值达标数，测过餐后的pd数", [rpost,rapost,mpost],2)
    print_excel("both平均值达标数，每个值达标数，同时测过餐前餐后的pd数", [rall,raall,mall],2)
    
    print_excel("premeal至少一个值不达标的pd数，测过餐前的pd数", [nrpre,mpre],2)
    print_excel("postmeal至少一个值不达标的pd数，测过餐后的pd数", [nrpost,mpost],2)
    print_excel("bed至少一个值不达标的pd数，测过睡前的pd数", [nrbed,mbed],2)
    
    print_excel("premeal每个pd里餐前测量次数", pdn1) #每个pd里餐前测量次数
    print_excel("postmeal每个pd里餐后测量次数", pdn2)
    print_excel("presleep每个pd里睡前测量次数", pdn3)
    
    print("the project costs: ", time.time()-start_time)
        
if __name__ == "__main__":
    PDvalue_main("total.xlsx", "nei", #startDate="2018-07-01", endDate="2019-06-30", 
                 district=['廿六病区','三十五病区', '三十六病区','三十八病区' , '三十九病区', '四十病区', '四十一病区'])
