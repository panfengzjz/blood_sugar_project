#coding: utf-8
import openpyxl
import time
def ward_classify(ward):
    res = ""
    if (ward=="10病区" or ward=="11病区"or ward=="12病区"or ward=="16病区"or ward=="17病区"or ward=="18病区"or ward=="1病区"or 
        ward=="2病区"or ward=="34病区"or ward=="35病区"or ward=="37病区"or ward=="3病区"or ward=="44病区"or ward=="45病区"or 
        ward=="46病区"or ward=="47病区"or ward=="4病区"or ward=="5病区"or ward=="6病区"or ward=="7病区"or ward=="8病区"or ward=="9病区"or ward=="日间病部"
        or ward=="十三病区"or ward=="十四病区"):
        res="wai"
    elif (ward=="19病区" or ward=="20病区"or ward=="22病区"or ward=="25病区"or ward=="24病区"or ward=="27病区"or ward=="28病区"or ward=="29病区"
          or ward=="30病区"or ward=="31病区"or ward=="41病区"or ward=="42病区"or ward=="43病区" or ward=="二十一病区"):
        res="nei"
    elif (ward=="外监" or ward=="急ICU"or ward=="肝外监"or ward=="心外监"):
        res="jian"
    elif (ward=="23病区" ):
        res="endo"
    return res

# src_name: 原始表格名称
# ret_name: 最终生成的新excel名称
def ward_exchange(src_name, ret_name1,ret_name2,ret_name3,ret_name4):
    wb2 = openpyxl.Workbook()
    wb2.save(ret_name1)
    wb3 = openpyxl.Workbook()
    wb3.save(ret_name2)
    wb4 = openpyxl.Workbook()
    wb4.save(ret_name3)   
    wb5 = openpyxl.Workbook()
    wb5.save(ret_name4)    
    print('新建成功')

    wb1 = openpyxl.load_workbook(src_name)
    wb2 = openpyxl.load_workbook(ret_name1)
    wb3 = openpyxl.load_workbook(ret_name2)
    wb4 = openpyxl.load_workbook(ret_name3)
    wb5 = openpyxl.load_workbook(ret_name4)
    
    sheet1 = wb1[wb1.sheetnames[0]]
    sheet2 = wb2[wb2.sheetnames[0]]
    sheet3 = wb3[wb3.sheetnames[0]]
    sheet4 = wb4[wb4.sheetnames[0]]
    sheet5 = wb5[wb5.sheetnames[0]]
    max_row = sheet1.max_row        #最大行数
    max_column = sheet1.max_column  #最大列数
    
    i = 2
    n1=2
    n2=2
    n3=2
    n4=2
    
    while(i < max_row+1):
        ward=sheet1.cell(i, 5).value
        ward_class=ward_classify(ward)
        
        if (ward_class=="wai"):
            for col in range(1, 15):
                sheet2.cell(n1, col).value = sheet1.cell(i, col).value             
            n1+=1  
        elif (ward_class=="nei"):
            for col in range(1, 15):
                sheet3.cell(n2, col).value = sheet1.cell(i, col).value
                
            n2+=1   
        elif (ward_class=="jian"):
            for col in range(1, 15):
                sheet4.cell(n3, col).value = sheet1.cell(i, col).value             
            n3+=1  
        elif (ward_class=="endo"):
            #print("endo")
            for col in range(1, 15):
                sheet5.cell(n4, col).value = sheet1.cell(i, col).value             
            n4+=1
            #print(n4)
        i += 1


    wb2.save(ret_name1)
    wb3.save(ret_name2)
    wb4.save(ret_name3)
    wb5.save(ret_name4)
    #保存数据
    wb1.close()
    wb2.close()
    wb3.close()
    wb4.close()
    wb5.close()
    print("功能完成")

if __name__ == "__main__":
    start=time.time()
    src_name = "new-huge-from201801.xlsx"
    ret_name1 = "wai.xlsx"   
    ret_name2 = "nei.xlsx"  
    ret_name3 = "jian.xlsx" 
    ret_name4="endo.xlsx"
    #patientList = make_list(fileName)   #这次生成一个数组，方便pop
    #reorder_excel(fileName, reslName)
    ward_exchange(src_name, ret_name1,ret_name2,ret_name3,ret_name4)
    print(time.time()-start)