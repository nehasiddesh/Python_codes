# -*- coding: utf-8 -*-
import pandas as pd
from tkinter import filedialog
from tkinter import *
import openpyxl
import xlrd
import numpy as np
import warnings

from numpy.core.defchararray import upper

warnings.filterwarnings("ignore")

root= Tk()
root.title('T-Space File Generator')
root.geometry("300x200")

root.filename1=filedialog.askopenfilename(initialdir="/Documents", title="Select Inter HO File",filetypes=(("All files","*.*"),("png files","*.png")))

cell_id_input = input("Enter cell names:")
cell_id_list = list(str(cell_id_input).split(","))

print("Reading:",root.filename1.split('/')[-1])
input_site_data=pd.read_excel(root.filename1)

for cell in cell_id_list:
    print(cell)
    if upper(cell.split("_")[1][-1]) == 'A':
        print("inside A")
        print(cell)
    if upper(cell.split("_")[1][-1]) == 'B':
        print("inside B")
        print(cell)
    if upper(cell.split("_")[1][-1]) == 'C':
        print("inside C")
        print(cell)





input_site_data.to_excel('dummy_1.xlsx',index=False)
print("******HO Attempts sheet generated******")


label= Label(root, text="T-Space Files Generated!!").pack()
root.mainloop() 
