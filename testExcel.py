import xlrd
import os
import xlwt

rd = xlrd.open_workbook('.\list.xlsx') # the name of excel
sheet = rd.sheet_by_name('Sheet1')  #sheet name
rowsCount = sheet.nrows
cellsCount = sheet.ncols

arr=[];
j=0
while j < cellsCount:
    if sheet.cell(0,j).value =='length':
        arr.insert(0,j) 
    if sheet.cell(0,j).value =='width':
        arr.insert(1,j)
    if sheet.cell(0,j).value =='high':
        arr.insert(2,j) 
    j=j+1

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('myworksheet');

i=1 
while i < rowsCount:
    worksheet.write(i, 0 ,sheet.cell(i, 0).value);
    worksheet.write(i, 1 ,str(int(sheet.cell(i, arr[0]).value))+ '*'+ str(int(sheet.cell(i, arr[1]).value))
     +'*'+ str(int(sheet.cell(i, arr[2]).value)));
    i=i+1

workbook.save('.\listNew.xls')

#print(sheet.name, rowsCount, sheet.ncols) 

