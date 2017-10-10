This script unzips all of the zip archives in the folder, renames my main math excel file, where all the calculations are being done, to the name of the excel file contained in the zip archive, and copies data from the source excel file to my main math file. 

```Python 
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 19 08:06:25 2017

@author: builduser
"""

import zipfile
import os
import shutil

from openpyxl import load_workbook,Workbook

import openpyxl
from OpenPyXLHelperFunctions import data_from_range, data_to_range

```
OpenPyXLHelperFunctions is a simple script that makes the openpyxl package a little easier to use for what I need it. 


```Python
reelSrcRange = 'C9:L30'
reelDestRange = 'B5'
```
sets the ranges where the source data is coming from and going to


```Python
def unzip_all(cwd):
    os.chdir(cwd)
    fileList = []
    for file in os.listdir():
        try:
            name,extension = file.split('.')
        except:
            name = file
            extension = 'apple'
            
        if extension == 'zip':
            name = file.split('_')[0]
            fileList.append(name + '.xlsm')
            with zipfile.ZipFile(file,'r') as filename:
                filename.extractall()
        

    return(fileList)
    
```
loops through all of the files in the directory, checks if the extension is a zip, and if so, uses the ZipFile package to unzip the file. There is a try - except which will handle the cases where there is no extension.

```Python
def copy_and_rename_all(fileList):
    filez = []
    for file in fileList:
        src = 'Math.xlsx'
        dest = 'DGE - ' + file.split('.')[0] + '.xlsx'
        
        filez.append(dest)
        
        shutil.copy(src,dest)
        
    return(filez)
```
Renames the math file where all of the calculations take place, to the name of the zipfile
    
```Python
def copy_and_paste(fileList,filez):
    if len(fileList) == len(filez):
        for index in range(len(fileList)):
            wb = load_workbook(fileList[index], data_only = True)
            data = data_from_range(reelSrcRange,'Strips',wb)
            wb.close()
            
            wb = load_workbook(filez[index], data_only= True)
            data_to_range(data,reelDestRange,wb.get_sheet_by_name('Table'),wb)
            
            wb.save(filez[index])

```
Uses the openpyxl package to copy data from the developers math sheets to my math sheets


```Python
fileList = unzip_all(cwd)
print('unzipped')
filez = copy_and_rename_all(fileList)
print('renamed')
copy_and_paste(fileList,filez)

```

This script was simple to create and edit to change for different purposes. It saves me quite a bit of time because there are >20 files. 
