import pandas as pd
status = True
#you may run the file and then provide details during prompt
#or edit the code below and run, then enter 1 to run without prompt

#insert filepath below:
filepath = "BookPathHere, sample: /users/admin/file.xlsx  ,  or , file.xlsx" 

sheetName = "SheetNameHere"

#insert targetpath below:
targetpath = "Path of text file to be created, sample /users/admin/output.txt or output.txt"

if int(input("Enter 1 to run the script without the prompt, if you want the prompt interface, enter 0:\n")) == 0:
    print("----------------------------------\nNOTE: Enter all paths without any preceeding or following \"'\" charachters")
    filepath = input("----------------------------------\nEnter the path of the file you want to Extract data from (format should be xlsx):\n\t")
    sheetName = input("----------------------------------0\nEnter the name of the sheet you want to extract data from:\n\t")
    while True:
        try:
        #You can manipulate the extractor's settings in the line below:
            data = pd.read_excel(filepath, sheet_name = sheetName, usecols = 'B:H', header=None, skiprows=6)
            
        except (FileNotFoundError):
            print("----------------------------------")
            print("ERROR: The File Path you provided was not found!")
            filepath = input("\nEnter the path of the file you want to Extract data from (format should be xlsx):\n\t")
            continue

        except (ValueError):
            print("----------------------------------")
            print("ERROR: The sheet you provided was not found!")
            sheetName = input("\nEnter the name of the sheet you want to extract data from:\n\t")
            continue
        
        break

targetpath = input("----------------------------------\nEnter the path and name of the text file you want your data to be written to, \nif such a file does not exist, the script will create it:\nSample: /users/admin/myfile.txt\n\n\t")

f = open(targetpath , 'w')
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
            f.write('ANSWER: '+str(data.iloc[i,j])+'\n\n')
    i += 1
    try:
        if str(data.iloc[i,0]) == 'nan':
            break
    except:
        break
f.close()
print(f"PROCESS COMPLETE! FILE CREATED AT {targetpath}")