import tkinter as tk
from tkinter import filedialog
import os
import sys
import pandas as pd
import numpy as np


def get_absolute_paths():
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(title="ZθとGBのxlsxファイルを選択", filetypes=(("xlsxファイル", "*.xlsx*"),))

    absolute_paths = []
    for file_path in file_paths:
        absolute_path = os.path.abspath(file_path)
        absolute_paths.append(absolute_path)

    print(absolute_paths)

    if (len(absolute_paths) != 2):
        print('ファイルは2こにして');
        return

    file1 = pd.read_excel(absolute_paths[0], sheet_name='Sheet1')
    _file1 = file1.to_numpy()
    file1_check = [row[3] for row in _file1]
    
    # 最初がZだったとき
    if(file1_check[1] == 'Z'):
        df_z = pd.read_excel(absolute_paths[0], sheet_name='Sheet2')
        list_z = df_z.to_numpy()
        
        # 2こ目のファイル読み取り
        file2 = pd.read_excel(absolute_paths[1], sheet_name='Sheet1')
        _file2 = file2.to_numpy()
        file2_check = [row[3] for row in _file2]
        
        # 2こ目がZだったとき 
        if(file2_check[1] == 'Z'):
            print('ZとZになってる')
            return
        # 2こ目がGだったとき
        elif(file2_check[1] == 'G'):
            df_gb = pd.read_excel(absolute_paths[1], sheet_name='Sheet2')
            list_gb = df_gb.to_numpy()
        else:
            print('違うファイルになってる')
            return
        
    # 最初がGだったとき
    elif(file1_check[1] =='G'):
        df_gb = pd.read_excel(absolute_paths[0], sheet_name='Sheet2')
        list_gb = df_gb.to_numpy()
        
        file2 = pd.read_excel(absolute_paths[1], sheet_name='Sheet1')
        _file2 = file2.to_numpy()
        file2_check = [row[3] for row in _file2]
        
        # ２こ目がZだったとき
        if(file2_check[1] == 'Z'):
            df_z = pd.read_excel(absolute_paths[1], sheet_name='Sheet2')
            list_z = df_z.to_numpy()
        # ２こ目がGだったとき
        elif(file2_check[1] == 'G'):
            print('GBとGBになってる')
            return
        else:
            print('違うファイルになってる')
            return
            
    else:
        print('違うファイルになってる')
        return
        
    f_z = [row[2] for row in list_z]
    f_z.pop(0)
    z = [row[4] for row in list_z]
    z.pop(0)
    fa = f_z[z.index(max(z))]

    f_gb = [row[2] for row in list_gb]
    f_gb.pop(0)
    g = [row[4] for row in list_gb]
    g.pop(0)
    b = [row[5] for row in list_gb]
    b.pop(0)

    Ym0 = max(g)
    fr = f_gb[g.index(Ym0)]
    f1 = f_gb[b.index(max(b))]
    f2 = f_gb[b.index(min(b))]

    print('共振周波数fr',fr, '\tHz', sep='\t')
    print('反共振周波数fa' ,fa, '\tHz', sep='\t')
    print('象限周波数f1', f1, '\tHz', sep='\t')
    print('象限周波数f2', f2, '\tHz', sep='\t')
    print('アド最大値Ym0', Ym0, 'S', sep='\t')

    Qm = fr/(f2-f1)
    Rm = 1/Ym0
    wa = 2 * np.pi * fa
    wr = 2 * np.pi * fr
    Lm = Qm*Rm/wr
    Cm = 1/(wr*wr*Lm)
    Cd = wr*wr*Cm/(wa*wa - wr*wr)

    print('\nQm\t',round(Qm))
    print('Rm\t',round(Rm,2), '\tOhm')
    print('Lm\t',round(Lm,3), '\tH')
    print('Cm\t',round(Cm*10e11, 2), '\tpF')
    print('Cd\t',round(Cd*10e11), '\tpF')





if __name__ == "__main__":
    get_absolute_paths()
    
    print('終了する場合は何か押す')
    input()
    sys.exit()