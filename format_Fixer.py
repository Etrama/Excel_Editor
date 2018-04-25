# -*- coding: utf-8 -*-
"""
Created on Wed Feb 07 18:49:38 2018

@author: kmoudgalya
"""
import openpyxl as xl
import os

print "Make sure that you have an empty file with all the necesary formatting \
    ready to go before this script is run."

location = raw_input(r"Enter the location where files to be copied into format are located: ")
formatfilelocation = raw_input("Enter the format file location: ")        

filenames = os.listdir(location)
for fname in filenames:
    print "File being copied: " + fname
    
    wb1 = xl.load_workbook(os.path.join(location,fname))
    ws1 = wb1.worksheets[0]
    
    #the first argument into os.pah.join cannot be a string literal,
    #gives an IO Error when it is.
    
    wb2 = xl.load_workbook(os.path.join(formatfilelocation,"excel format file.xlsx"))
    
    i = 0
    for ws1 in wb1.worksheets:
        ws2 = wb2.worksheets[i]
        for row in ws1:
            for cell in row:
                ws2[cell.coordinate].value = cell.value
        if i<=9:
            i = i + 1
    
    ws2 = wb2.worksheets[0]    
    wb2.save(os.path.join(location,"FF"+fname))
    print "New file created successfully."
    print ""