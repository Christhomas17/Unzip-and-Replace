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



```Python
#This is what I call OpenPyXLHelperFunctions. Openpxl is great but I wanted to be able to input ranges using A1 notation so I created these helper functions.
import os
from openpyxl import load_workbook,Workbook
from openpyxl.utils import get_column_interval
import openpyxl
import re

import pandas as pd
import numpy as np

#gets data from a range
#sample execution
	#wb = load_workbook('AVV045751.xlsx',read_only = True, data_only = True)

	#strips = data_from_range('C9:J442','Strips',wb)

# wb = load_workbook('Sample.xlsx', data_only = True)
# ws = wb.get_sheet_by_name('Data')
# rng = 'B4:C7'


def data_from_range(RngAsString,WsAsString,wb):    
    ws = wb.get_sheet_by_name(WsAsString)
    rng = RngAsString
    
    try:
        start,end = rng.split(':')
    except:
        start = rng
        end = start
#    StartCol,EndCol = Get_Ranges(rng)
    
    data = []
    
    for row in ws[start:end]:
        data.append([cell.value for cell in row])
        
    df = pd.DataFrame(data)
    
    return(df)

def data_to_range(DataAsDataframe,FirstCell,ws,wb):
    data = DataAsDataframe
    rng = FirstCell
    
    col =  "".join(re.findall("[A-Z]",rng, flags = re.I))
        
    col = col_letter_to_number(col)
    row = int("".join(re.findall("[0-9]",rng)))      
        
    try:
        NumRows,NumCols = data.shape           
        
        Rows = np.arange(row, row + NumRows,1)
        Cols = np.arange(col, col + NumCols,1)
        
    
        for XIndex,x in enumerate(Cols):
            for YIndex,y in enumerate(Rows):           
                ws.cell(row = y, column = x).value = data.iloc[YIndex,XIndex]
    except:
        ws.cell(row = row, column = col).value = data
   
def Clear_Range(RangeAsString,ws,wb):
    for row in ws[RangeAsString]:
        for cell in row:
            cell.value = None

####################################
#Helper Functions
###################################            
def Get_Ranges(RangeAsString):
    rng = RangeAsString
    
    
    #in case there is only one range being given
    try:
        start,end = rng.split(':')
    except:
        start = rng
        end = start   

        
    StartCol =  "".join(re.findall("[A-Z]",start, flags = re.I))
    EndCol =  "".join(re.findall("[A-Z]",end, flags = re.I))
    
    return[StartCol,EndCol]            

def col_letter_to_number(ColLetter):
    letter = ColLetter.lower()
    
    alphabet = list('abcdefghijklmnopqrstuvwxyz')
    
    
    if len(letter) == 1:
        ColNum = alphabet.index(letter)+1
                               
        return(ColNum)
                        
    else:
        Set = letter[0]
        SetNum = alphabet.index(Set) + 1
        SetNum = SetNum * 26
        
        Letter = letter[1]
        
        
        ColNum = alphabet.index(Letter) + 1
                               
        return(SetNum + ColNum)

```
