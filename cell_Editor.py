# -*- coding: utf-8 -*-
"""
Created on Tue Feb 06 15:16:23 2018

@author: kmoudgalya
"""

import os
import openpyxl as xl

location = raw_input("Enter the directory path where the files whose cells are to be edited are located: ")

filenames = os.listdir(location)

#the specific part
sheet_name = raw_input("Enter the sheet name where the cell to be edited is located, such 'Sheet 1': ")
cell_location = raw_input("Enter the location of the cell which is to be edited: ")
for fname in filenames:
    wb = xl.load_workbook(os.path.join(location,fname))
    ws = wb[sheet_name]
    ws[cell_location] = ""
    
    wb.save(os.path.join(location,"CE"+fname))
    

