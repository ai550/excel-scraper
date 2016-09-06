from openpyxl import load_workbook
from os.path import exists
import re
import os



def dir_search(directory):
    for file in os.listdir(directory):
        if(os.path.isfile(directory + "\\" + file)):
            # if file.endswith(".xlsx"):
            print("---File: " + file)    
        if(os.path.isdir(directory + "\\" + file)):
            print("Directory: " + file)
            dir_search(directory + "\\" + file)


dir_search("Z:\\04. Assets and Configuration Management\\02. Assets\\01. Contracts\\From SAP\\2005-0351")
