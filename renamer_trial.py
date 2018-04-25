# -*- coding: utf-8 -*-
"""
Created on Sat Feb 03 23:59:30 2018

@author: kmoudgalya
"""

import os 
#dirs = os.listdir(path)
#os.path.join will make directory path reading platform agnostic as windows uses one kind of slash and linux might ue another kind of slash
#prefix = raw_input(" Enter the common prefix the files have: ")

location = raw_input(r"Enter the directory path where the files to be renamed are located:(The outlook items folder) ")

location2 = raw_input(r"Enter the directory of pdfs: ")

#newprefix = raw_input("Enter the new prefix for these files: ")
filenamessrc = os.listdir(location)
filenamesdest = os.listdir(location2)

i = 0 
for fname in filenamesdest:
    #if fname.startswith(prefix):
    fname_src = filenamessrc[i]
    i = i + 1
    new_name = fname_src[i]
    os.rename(os.path.join(location2,fname),os.path.join(location2,new_name))
