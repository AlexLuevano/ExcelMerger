import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import filedialog
from sys import exit
import os
from datetime import datetime
from time import strftime, gmtime

window = Tk()
files = filedialog.askopenfilenames(parent = window, title = 'Escoge archivos')
files_str = str(files)
files_list = list(files)
window.destroy()
window.mainloop()

#print(files_list)

if os.path.exists('C:\\temporary'):
    pass
else:
    print('C:\\temporary directory does not exist. Do you want to create it? Y/N?')
    answer = input()
    if answer == 'Y' or 'y' or 'Yes' or 'YES' or 'Yes':
        os.mkdir(r'C:\\temporary')

    elif answer == 'N' or 'n' or 'no' or 'NO' or 'No':
        print('Ok. Finishing execution')
        exit(0)

    else:
        print('Invalid answer. Finishing execution')
        exit(0)

now = datetime.now()
date = now.strftime("%Y-%m-%d_%H-%M-%S")
#path = os.path.join('C:\\Users',os.getlogin(),'Desktop','xls_merged_' + str(date)+'.xlsx')
path = os.path.join('C:\\temporary','xls_merged_' + str(date)+'.xlsx')
df = pd.DataFrame()
writer = pd.ExcelWriter(path,engine='xlsxwriter')

for f in files_list:
    if f[-4:] == '.csv':
        data = pd.read_csv(f, index_col=False, header=0, na_filter = False)
        df = df.append(data)
        print(f'{f} merged')
    else:
        try:
            data = pd.read_excel(f,sheet_name='Entities', index_col=False, header=1, na_filter = False)
            df = df.append(data)
        except:
            print('Sheet "Entitites" not found. Merging the first sheet of the file instead.')
            data = pd.read_excel(f,index_col=False, na_filter = False)
            df = df.append(data)

print('All files read and merged')
df.to_excel(writer,sheet_name='Entities',index = False, header = True)
writer.save()
print('File saved. Please review file at C:\\temporary')
#writer.close()
exit(0)