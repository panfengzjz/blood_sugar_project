#coding:utf-8
import xlrd
import xlwt
from xlutils.copy import copy

class patient():
    def __init__(self, info):
        self.name = info[0]
        self.age = info[1]
        self.bld_glu = info[7]
        self.opr_dat = info[8]
        self.opr_tim = info[13]
        
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


def recordFreq_main(filename, sheetname):
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
        if (former.name==p.name):
            n +=1
        else:
            s.append(n)
            n=1
    print_excel("每人记录条数",s,2)
    print_excel("总人数",[len(s)],2)

if __name__ == "__main__":
    recordFreq_main("wai.xlsx", "Sheet")
