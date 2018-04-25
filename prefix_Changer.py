# -*- coding: utf-8 -*-
"""
Created on Sat Feb 03 23:59:30 2018

@author: kmoudgalya
"""

import os 
#dirs = os.listdir(path)
#os.path.join will make directory path reading platform agnostic as windows uses one kind of slash and linux might ue another kind of slash
prefix = raw_input(" Enter the common prefix the files have: ")

location = raw_input(r"Enter the directory path where the files to be renamed are located: ")

newprefix = raw_input("Enter the new prefix for these files: ")

filenames = os.listdir(location)

for fname in filenames:
    if fname.startswith(prefix):
        os.rename(os.path.join(location,fname),os.path.join(location,fname.replace(prefix,newprefix)))
