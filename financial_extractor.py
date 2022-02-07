import os
import openpyxl as xl
import re

#Format is [account, date, description, debit, credit, balance, category, sub category, month, year]

this_folder = os.path.dirname(os.path.abspath(__file__))
pc_log = os.path.join(this_folder, 'report.csv')
pc_file = open(pc_log)
pc_info = pc_file.readlines()

xl_path = os.path.join(this_folder, 'export.xlsx')
xl_file = xl.Workbook()

sheet = xl_file['Sheet']

def extract_pc(r, info_file): #extract row 'r' information from PC Financial
    info = re.findall('"([^"]*)"', info_file[r])
    description = info[0]
    transaction_type = info[1]
    transaction_date = info[3]
    transaction_amount = -float(info[5])
    return["PC Financial Card", transaction_date, description, transaction_amount, transaction_type]

def print_pc(r, data_list, destination_sheet):
    destination_sheet['A' + str(r)].value = data_list[0]
    destination_sheet['B' + str(r)].value = data_list[1]
    destination_sheet['C' + str(r)].value = data_list[2]
    destination_sheet['D' + str(r)].value = data_list[3]

print_row = 1

for i in range(len(pc_info)-1,1,-1):
    pc_data = extract_pc(i, pc_info)
    if pc_data[4] != "PAYMENT":
        print_pc(print_row, pc_data, sheet)
        print_row += 1


xl_file.save(xl_path)
