#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov 21 22:14:25 2020

@author: troy
"""

from tkinter import *
from tkinter import ttk

root = Tk()
root.geometry("600x200")
root.title("Intake Information")

root.grid_columnconfigure((0,1), weight=1)

Label3 = Label(root, text="First Name")
Label4 = Label(root, text="Last Name")
Label5 = Label(root, text="City")
Label6 = Label(root, text="Loan Number")

Entry3 = Entry(root)
Entry4 = Entry(root)
Entry5 = Entry(root)
Entry6 = Entry(root)
button = ttk.Button(root, text='Get Documents')

Label3.grid(row=3, column=0)
Entry3.grid(row=3, column=1, sticky="ew")
Label4.grid(row=4, column=0)
Entry4.grid(row=4, column=1, sticky="ew")
Label5.grid(row=5, column=0)
Entry5.grid(row=5, column=1, sticky="ew")
Label6.grid(row=6, column=0, sticky="ew")
Entry6.grid(row=6, column=1, sticky="ew")
button.grid(row=10, column=1, sticky="ew")

def print_go():
    print(Entry3.get())
    
button.config(command=print_go)

root.mainloop()