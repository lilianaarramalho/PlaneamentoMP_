import tkinter.ttk as ttk
from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
import subprocess
import pathlib

cwd=os.getcwd()

def runmain():

    global path

    path = pathlib.Path(__file__).parent.absolute()

    os.chdir(path)

    print("A correr algoritmo de planeamento")

    root.destroy()

    subprocess.run(['main.exe'])

    input("\n\nPress enter to exit.")

root=Tk()

root.geometry('400x200')

root.title("Planeameno Amorim")

grupo_1=ttk.LabelFrame(root,text='Calcular data')


var=IntVar()
rb_Method0=ttk.Radiobutton(grupo_1,text='Gerar resultados',variable=var,value=0)
rb_Method1=ttk.Radiobutton(grupo_1,text='Gerar cumprimento do plano',variable=var,value=1)
var.set(0)

btn_run=ttk.Button(grupo_1,text="Correr",command=lambda: runmain())

grupo_1.grid(row=2,columnspan=6,padx=5,pady=10,ipadx=5,ipady=5)
rb_Method0.grid(row=3, column=1, sticky='WE', padx=2, pady=5)
rb_Method1.grid(row=3, column=2, sticky='WE', padx=2, pady=5)
btn_run.grid(row=4, column=4, sticky="W", pady=3)

col_count, row_count = root.grid_size()

for col in range(col_count):
    root.grid_columnconfigure(col, weight=2)

for row in range(row_count):
    root.grid_rowconfigure(row, weight=1)

if var==0:
    print('debug')

root.mainloop()



