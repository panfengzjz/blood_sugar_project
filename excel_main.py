#coding: utf-8
from dayCount_neiwai import neiwai_main     # file 1.1 住院天数变化 餐前高血糖 内外
from dayCount_jian import jian_main         # file 1.2 住院天数变化 不分时刻 总的>10  监专用
from patientDay_PDvalue import PDvalue_main # file 2 输出pd均值、pd个数、达标情况、pd里测量次数 内外
from patientDay_PDvalue_jian import PDvalue_main_jian #2.2 不分时刻pd总数、pd个数、至少一次低血糖pd  内外
from patientDay_PDvalue_jian_spcf import PDvalue_main_jian_spcf #2.3 输出pd均值、pd个数、达标情况、pd里测量次数+不分时刻pd总数、pd个数、至少一次低血糖pd  监
from incidenceRate import rate_main         # file 3 低血糖发生与复测。总single条数。通用
from divideData import divide_main          # file 4.1 分界8-20 内外 
from divideData_jian import divide_main_jian          # file 4.2 分界8-20 监
from recordFreq import recordFreq_main      # file 5 每人测的条数。通用

if __name__ == "__main__":
    excelname="huge.xlsx"
    sheetname="nei"
    #分时刻
    #neiwai_main(excelname, sheetname)
    PDvalue_main(excelname, sheetname, startDate="2017-06-01", endDate="2017-06-31")
    #PDvalue_main_jian(excelname, sheetname)    
    divide_main("patientday.xls", "Sheet1")
    
    #不分时刻
    #jian_main(excelname, sheetname)
    #PDvalue_main_jian_spcf(excelname, sheetname)
    #divide_main_jian("patientday.xls", "Sheet1")
    
    #通用
    #rate_main(excelname, sheetname)
    #recordFreq_main(excelname, sheetname)