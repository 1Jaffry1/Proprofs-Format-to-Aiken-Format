import numpy as np
import pandas as pd
import docx as dox

#insert filepath below:
filepath = "/Users/muhammad/Projects/Mafaheem/Level 2/Exam Level 2, Book 1.xlsx" 

#insert targetpath below:
targetpath = "/Users/muhammad/Projects/Mafaheem/Level 2/Exam 1-16partial.txt"



data = pd.read_excel(filepath, sheet_name = f"Exam 1-16", usecols = 'B:H', header=None, skiprows=6)
f = open(targetpath , 'w')
i = 0
m=1
while m<13:
    for j in range(7):
        if j == 0:
            f.write(str(data.iloc[i,j])+'\n')
        elif j == 1:
            f.write('A. '+str(data.iloc[i,j])+'\n')
        elif j == 2 and str(data.iloc[i,j]) != 'nan':
            f.write('B. '+str(data.iloc[i,j])+'\n')
        elif j == 3 and str(data.iloc[i,j]) != 'nan':
            f.write('C. '+str(data.iloc[i,j])+'\n')
        elif j == 4 and str(data.iloc[i,j]) != 'nan':
            f.write('D.: '+str(data.iloc[i,j])+'\n')
        elif j == 5 and str(data.iloc[i,j]) != 'nan':
            f.write('E. '+str(data.iloc[i,j])+'\n')
        elif j == 6:
            f.write('ANSWER: '+str(data.iloc[i,j])+'\n\n')
    i += 1
    m+=1
    try:
        if str(data.iloc[i,0]) == 'nan':
            break
    except:
        break
f.close()


print(f"PROCESS COMPLETE! FILE CREATED AT {targetpath}")
