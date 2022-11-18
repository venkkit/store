
from openpyxl import load_workbook
import xlrd
from datetime import datetime
import datetime
import os




dir = os.path.join('bills')
if not os.path.exists(dir):
    os.mkdir(dir)


# Read templet
workbook = load_workbook(filename="template.xlsx")
sheet_T = workbook.active
 

# Read Data sheet
loc = ("data.xlsx")
wb = xlrd.open_workbook(loc)
data_sheet = wb.sheet_by_index(0)

cols = data_sheet.ncols
rows = data_sheet.nrows
Invoice = int(input("Please enter last invoice number : "))

for i in range(1,rows):
    Date_v = (data_sheet.cell_value(i, 5))
    conv = (Date_v - 25569) * 86400.0
    Date = str((datetime.datetime.utcfromtimestamp(conv)).date())



    Invoice = Invoice+1
    invoice = str(Invoice)
    name = data_sheet.cell_value(i, 0)
    email = data_sheet.cell_value(i, 1)
    mobile = int(data_sheet.cell_value(i, 2))
    amount = data_sheet.cell_value(i, 3)
    quantity = data_sheet.cell_value(i, 4)



    # print(Invoice)
    # print(name)
    # print(email)
    # print(mobile)

    #bill sheets 

    sheet_T["A11"] = "Bill To : " + name
    sheet_T["A9"] = "Invoice No :" + invoice
    sheet_T["K24"] =  amount
    sheet_T["B12"] =  email
    sheet_T["B13"] =  str(mobile)
    sheet_T["A10"] = "Invoice Date : " + Date
    sheet_T["A30"] = Date
    sheet_T["I19"] = quantity

    #save the file
    workbook.save(filename="bills/"+name+".xlsx")
    print(invoice,name,"bill generated ")

print("---------*** All bills are completed ***---------")