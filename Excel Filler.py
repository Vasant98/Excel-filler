# -*- coding: utf-8 -*-
"""
Created on Thu Apr 21 14:16:05 2022

@author: WangSheng
"""


import pandas as pd
import os
from openpyxl import Workbook, load_workbook

wb = load_workbook(r'C:\Users\WangSheng\Downloads\OneDrive\COC_Template.xlsx')

input_data = os.path.normpath(os.path.expanduser
                             (r"C:\Users\WangSheng\Downloads\OneDrive\New_V2 Product Lot & SN release date.csv"))

df = pd.read_csv(input_data)
input_df = df[['SN', 'DATE']]


for i in input_df.index:
    SN = df['SN'][i]
    DATE = df['DATE'][i]
    
    ws = wb.active
    ws['A18'].value = SN
    ws['F41'].value = DATE
    
    wb.save('F-826-01-004_Rev 0_COC_' + SN + '.xlsx')
