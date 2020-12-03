import pandas as pd
import os
import shutil
import openpyxl

# import company information
company_info = pd.read_excel('company_info.xlsx')


# make a function to copy the template, fix the text
def make_transfer_letter(company_name,representative_name,company_letter_address):
    shutil.copyfile('letter.xlsx','送付状_{company_name}.xlsx'.format(company_name=company_name))
    excel=openpyxl.load_workbook('送付状_{company_name}.xlsx'.format(company_name=company_name))
    sheet = excel['送付状']
    sheet['A7']=company_name
    sheet['A8']=representative_name
    sheet['F7']=company_letter_address
    excel.save('送付状_{company_name}.xlsx'.format(company_name=company_name))
# itterate through all companies
for index,row in company_info.iterrows():
    make_transfer_letter(row['company_name'],row['representative_name'],row['company_letter_address'])
