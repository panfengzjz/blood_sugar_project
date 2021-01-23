#coding:utf-8
import xlrd
import xlwt
from xlutils.copy import copy
from dayCount_print import print_excel

class patient():
    def __init__(self, info):
        self.name = info[0]
        self.age = info[1]
        self.hos_area= info[4]
        self.bld_glu = info[7]
        self.opr_dat = info[8]
        self.opr_tim = info[13]

def recordFreq_main(filename, sheetname, startDate="", endDate="", district=[]):
    fd_excel = xlrd.open_workbook(filename) #打开文件
    table = fd_excel.sheet_by_name(sheetname)    #读取sheet0
    workbook = xlwt.Workbook(encoding = 'ascii') #创建
    #worksheet = workbook.add_sheet('sheet2')    
    print("LOADED FILE")
    max_row = table.nrows
    n=1
    s=[]
    for r in range(2,max_row):
        p = patient(table.row_values(r))    
        former=patient(table.row_values(r-1))
        
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

        if (former.name==p.name):
            n +=1
        else:
            s.append(n)
            n=1
    print_excel("每人记录条数",s,2)
    print_excel("总人数",[len(s)],2)

if __name__ == "__main__":
    recordFreq_main("wai.xlsx", "Sheet", startDate="2018-07-01", endDate="2019-06-30", 
                    district=['2病区', '4病区'])
