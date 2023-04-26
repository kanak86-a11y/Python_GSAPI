from file_ops import fileoperations
from gsheet import Googlesheetsapi2
filename = input("Enter your filename:")
file_r = fileoperations.excel_read_file(filename)
file_ = file_r.copy()
read_gsheet = Googlesheetsapi2.call_sheet_api(file_r[1:])
print(read_gsheet)
status_of_file_delete = fileoperations.remove_file(filename)
print(status_of_file_delete)
print('file is copied')
print('file is deleted')

# outsp = input("Put the output here: ")
# output = outsp.replace(',')
# print("output:", output)
# import openpyxl
# def excel_read_file():
#     wb = openpyxl.load_workbook('Test_Data_Logging_1.xlsx')
#     # Sheet = wb.sheetnames
#     print(wb.active.title)
#     sh1 = wb['data1']
#     row = sh1.max_row
#     column = sh1.max_column
#     local_excel_data = []
#     for i in range(1, row+1):
#         row_data = []
#         for j in range(1, column+1): 
#             row_data.append(sh1.cell(row=i, column=j).value)
#         local_excel_data.append(row_data)
#     print(local_excel_data)
# read_ = excel_read_file()

# filess = fileoperations.os_path

        # print(sh1.cell(row=i, column=j).value)
# print(len(row_data),type(row_data), row_data)
# print(len(local_excel_data), type(local_excel_data), local_excel_data)
















# service_account_file = 'leafy.json'
# scopes = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
# spreadsheet_id = '1hzv14EF-TppuBWQAVYQ9s2kcHi12yPJH5oVcPqGfM0c'
# range_name = 'sheet1'
# sheets_api = Googlesheetsapi2(service_account_file,scopes,spreadsheet_id,range_name)
# if sheets_api.is_connected():
#     sheets_api.write_data2121('xbusdcbxw')
#     data = sheets_api.read_data()
#     print(data,'/n/nconnected')
# else:
#     print('not connected')
# test_append = ['test']
# resu = Googlesheetsapi2.write_data2121(sheets_api, 'test')

# files = excel.delete_file()
# print('file is deleted')

