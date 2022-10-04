#!/usr/bin/env python3
import excel
import textsplit
import tkinter as tk
from tkinter import simpledialog

#GUI using Tkinter

window = tk.Tk()

window.withdraw()

try:
    user_input = simpledialog.askstring(title = "TXT Path", prompt = "Enter filepath for the Test Coverage Report text file: ")
    pname = simpledialog.askstring(title = "TXT Path", prompt = "Enter Project Name: ")
    date = simpledialog.askstring(title = "TXT Path", prompt = "Enter Date: ")
    time = simpledialog.askstring(title = "TXT Path", prompt = "Enter Time: ")

    #Calling main functions
    textsplit.textsplit(user_input)
    excel.createexcel(pname, date, time)
except:
    print("Inaccurate Filepath or Canceled Process")



