# This program will parse the information from a powder difraction file (PDF)
# that has been saved as MS-excell sheet.
import pandas as pd
import numpy as np
import os

#%%
f_name = (os.listdir('File_To_Parse'))#list of files in "file_To-Parse"
f_howmany = range(len(f_name))

# print("List of files")
print(pd.DataFrame(f_name, columns=['File Name']))

#Reads each sheet from the files into dictionary with each sheet becomes a df
# for x in f_howmany:
#     df = pd.read_excel(os.path.join('File_To_Parse', f_name[x]), header=None, usecols='A',nrows=27, sheet_name=None)
    
x = int(input("Enter index of file to parse")) # select index of file you want from 'f_name' list
SheetDict = pd.read_excel(os.path.join('File_To_Parse', f_name[x]), header=None,
                   usecols='A',nrows=27, sheet_name=None)

# get list of keys in SheetDict
key_list = []
key_length = []
for x in SheetDict.keys():
  key_list.append(x)

    
#%% Put desired data into df
 
PDF_info = pd.DataFrame(columns= ['PDF #', 'Name', 'Formula', 'Crystal System',
                                  'Referance', 'Notes' ])

PDF_Number = []
PDF_Name = []
for x in range(len(key_list)):
    Ugg = SheetDict[key_list[x]]
    
    temp = str(Ugg.loc[0,0])
    PDF_Number = temp[12:23]
    
    temp = str(Ugg.loc[1, 0])
    PDF_Name = temp[6:]
    
    temp = str(Ugg.loc[2, 0])
    PDF_Formula = temp[9:]
    
    temp = str(Ugg.loc[15, 0])
    PDF_System = temp[16:]
    
    temp = str(Ugg.loc[12, 0])
    PDF_Ref = temp[11:]
    
    PDF_Notes = str(Ugg.loc[26, 0])
    
    PDF_info = PDF_info.append({'PDF #' : PDF_Number, 'Name' : PDF_Name, 
                                'Formula' : PDF_Formula, 'Crystal System' : PDF_System,
                                  'Referance' : PDF_Ref, 'Notes' : PDF_Notes},
                               ignore_index = True)
save_as = input("Name for CSV output")
# # Uncomment to save file
PDF_info.to_csv(os.path.join('Parsed_files', save_as + '.csv'))
