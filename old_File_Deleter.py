# -*- coding: utf-8 -*-
"""
Created on Wed Feb 07 15:59:51 2018

@author: kmoudgalya
"""

import os

location = raw_input(r"Enter the directory where you want to delete files from: ")
prefix = raw_input("Enter the prefix for the files that are to be deleted:")

for fname in os.listdir(location):
    if fname.startswith(prefix):
        os.remove(os.path.join(location,fname))
