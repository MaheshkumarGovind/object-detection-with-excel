# -*- coding: utf-8 -*-
"""
Created on Wed Aug 24 12:23:51 2022

@author: Mahesh
"""

import openpyxl


workbook  = openpyxl.load_workbook("E:\Mahesh AI\Mahesh Excel.xlsx")
sheet = workbook.active 
sheet.tittle = 'changed'




sheet['A1'] = "10"
sheet['B1'] = "10"
sheet['C1'] = "10"
sheet['A2'] = "10"
sheet['B2'] = "10"
sheet['A3'] = "10"
sheet['B3'] = "10"

sheet.insert_rows(1)
sheet.insert_cols(2)

workbook.save('E:\Mahesh AI\Mahesh Excel.xlsx')