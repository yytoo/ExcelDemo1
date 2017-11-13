import xlrd
import xdrlib ,sys
import xlwt


def open_excel(file):
    try:  
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)
       
def isEqual(sheetsCopy,cellValueN):
    for rowIndex in range(0,10):
        cellValue=sheetsCopy.cell_value(rowIndex,1)
        cellValue=str(cellValue)
        if(cellValueN==cellValue):
           return 0
        if cellValue=='':
           break 
    return 1
       
def excel_table_byindex():
    data = open_excel('test.xlsx')
    wb=xlwt.Workbook()
    sheetsCopy = wb.add_sheet('Sheet1')
    wb.save('test2.xls')
    sheets = data.sheet_by_name("Sheet1")
    for rowIndex in range(0,10):
        cellValue=sheets.cell_value(rowIndex,1)
        cellValue = str(cellValue)
        cellValue=cellValue.strip()
        print(isEqual(sheetsCopy,cellValue))
        sheetsCopy.write(rowIndex,1,cellValue)
        if cellValue=='@_@':
            break
        rowIndex = rowIndex + 1
        print cellValue
    wb.save('test2.xls')
    
    
excel_table_byindex()
        
    