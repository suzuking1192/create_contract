import pandas as pd
from docx import Document
import os
import shutil

# Get necessary information from excel
company_info = pd.read_excel('company_info.xlsx')

# make a copy of template
# make a function to read the information of word
def make_contract(company_name,company_home_address,ceo_name,contract_year,contract_month,contract_day,contract_name):
    shutil.copyfile('templatecontract.docx','覚書_{company_name}.docx'.format(company_name=company_name))
    document = Document('覚書_{company_name}.docx'.format(company_name=company_name))

# make a function to add the company name
    for paragraph in document.paragraphs:
        if '（以下「甲」という）' in paragraph.text:
            paragraph.text = paragraph.text.replace('（以下「甲」という）','{company_name}（以下「甲」という）'.format(company_name=company_name))

# make a function to add contract date
        if 'JapanWork利用規約（企業向け）' in paragraph.text:
            paragraph.text = paragraph.text.replace('JapanWork利用規約（企業向け）',contract_name)

        if '　年' in paragraph.text:
            paragraph.text = paragraph.text.replace('　年', '{year}年'.format(year=int(contract_year)))

        if '　月' in paragraph.text:
            paragraph.text = paragraph.text.replace('　月', '{month}月'.format(month=int(contract_month)))

        if '　日' in paragraph.text:
            paragraph.text = paragraph.text.replace('　日', '{day}日'.format(day=int(contract_day)))

# make a function to add company information

        if '甲　　　' in paragraph.text:
            paragraph.text = paragraph.text.replace('甲', '甲　{company_home_address}\n 　　　　　　　　　　　　　　　　　　{company_name}\n　　　　　　　　　　　　　　　　　　{ceo_name}'.format(company_home_address=company_home_address,company_name=company_name,ceo_name=ceo_name))

# make a function to save the change
    document.save('覚書_{company_name}.docx'.format(company_name=company_name))

# iterate through all companies
for index,row in company_info.iterrows():
    make_contract(row['company_name'],row['company_home_address'],row['ceo_name'],row['contract_year'],row['contract_month'],row['contract_day'],row['contract_name'])

# test
