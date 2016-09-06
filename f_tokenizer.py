import xlrd
import re
import openpyxl

mrf_numbers = xlrd.open_workbook('C:\\Users\\nk91009578.NCOC\\Desktop\\report_Bolat.xlsx')
mrf_sheet = mrf_numbers.sheet_by_index(0)
mrfs = []
for rownum in range(mrf_sheet.nrows):
    row = mrf_sheet.row_values(rownum)
    for c_el in row:
        m = re.search('(?<=#)\w+', c_el)
        a = m.group(0)
        mrfs.append(a)
    
