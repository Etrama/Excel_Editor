# -*- coding: utf-8 -*-
"""
Created on Sun Feb 04 16:53:44 2018

@author: kmoudgalya
"""

#file should not be open in excel, else you'll get a permission denied error..
import openpyxl
import os
#ppts are nto in scope here!!!
location = raw_input(r"Enter the directory path where the files to be edited are located: ")
# progresbar installed with:
#conda install -c anaconda progressbar 
#use this in windows prompt, without python open.
filenames = os.listdir(location)

for fname in filenames:
    print "Removing version tab from: " + fname
    workbook = openpyxl.load_workbook(os.path.join(location,fname))
    std = workbook.get_sheet_by_name('Version')
    workbook.remove_sheet(std)
    workbook.save(os.path.join(location,fname))


