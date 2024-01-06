import numpy as np
import pandas as pd
data = pd.read_excel('readFromThis.xlsx',sheet_name = 'FromThisSheet', usecols = 'B:H', header=None, skiprows=6)

f = open('writeToThis.txt', 'w')
# for i in range(28):
i = 0
while True:
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
            f.write('Answer: '+str(data.iloc[i,j])+'\n\n')
    i += 1
    try:
        if str(data.iloc[i,0]) == 'nan':
            break
    except:
        break
f.close
print(data)
