from enum import Enum
import qrcode
import xlrd

input_location = "input.xlsx"
class ExcelCols(Enum):
    ID=0
    EMAIL=3
    NAME=4
    EMPNO=6
    DINNER=7
    DRINK=8
export_params = [ExcelCols.ID,ExcelCols.NAME]

def generateQR(data,fileName):
    img = qrcode.make(data)
    img.save("output/"+fileName+".png")


def readExcelInput():
    wb = xlrd.open_workbook(input_location)
    sheet = wb.sheet_by_index(0)
    return sheet

sheet = readExcelInput()

for i in range(sheet.nrows):
    result= ""
    for param in export_params:
        result+= str(sheet.cell_value(i,param._value_))
    generateQR(result,str(sheet.cell_value(i,ExcelCols.NAME._value_)))
    print(result)