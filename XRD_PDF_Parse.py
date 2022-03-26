# This program will parse the information from a powder difraction file (PDF)
# that has been saved as MS-excell sheet.
import pandas as pd
import numpy as np
import os

#%%
f_name = (os.listdir('File_To_Parse'))#list of files in "file_To-Parse"
f_howmany = range(len(f_name))

#Reads each sheet from the files into dictionary with each sheet becomes a df
# for x in f_howmany:
#     df = pd.read_excel(os.path.join('File_To_Parse', f_name[x]), header=None, usecols='A',nrows=27, sheet_name=None)
    
x = 0 # select index of file you want from 'f_name' list
SheetDict = pd.read_excel(os.path.join('File_To_Parse', f_name[x]), header=None,
                   usecols='A',nrows=27, sheet_name=None)

# get list of keys in SheetDict
key_list = []
key_length = []
for x in SheetDict.keys():
  key_list.append(x)

  
PDF_info = pd.DataFrame()
test = []

for key, value in SheetDict.items():
    s = value[0]
    test.append(s)
    
#%% Test
PDF_Number = []
PDF_Name = []
for x in range(len(key_list)):
    Ugg = SheetDict[key_list[x]]
    temp = str(Ugg.loc[0,0])
    PDF_Number.append(temp[12:24])
    temp = str(Ugg.loc[1, 0])
    PDF_Name.append(temp[6:])
