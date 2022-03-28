from enum import Enum
import qrcode
import xlrd
import base64
import os
from Crypto.Cipher import AES

input_location = "input.xlsx"

key_length = 16
secret_key = 'TESTTESTTESTTEST'
class ExcelCols(Enum):
    ID=0
    EMAIL=3
    NAME=4
    EMPNO=6
    DINNER=7
    DRINK=8
export_params = [ExcelCols.ID,ExcelCols.NAME]

def gen_encoded_secret():
    data_string = secret_key.encode("utf-8")
    encoded_key = base64.b64encode(data_string)
    return encoded_key

def pad_str(str, c):
    result = str + (c * (16-len(str)%16))
    return result

def encrypt_str(msg, key):
    decoded_key = base64.b64decode(key)
    aes_cipher = AES.new(decoded_key)
    encrypted_str = aes_cipher.encrypt(pad_str(msg, '%'))
    return base64.b64encode(encrypted_str)

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
    encoded_sec = gen_encoded_secret()
    en_str = encrypt_str(result, encoded_sec)
    generateQR(en_str,str(sheet.cell_value(i,ExcelCols.NAME._value_)))
    print(en_str)