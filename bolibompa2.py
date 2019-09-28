import csv
import datetime
from decimal import *
import re
import tkinter
from tkinter import messagebox
from tkinter import filedialog

getcontext().prec = 2 #behövs för att tidskonverteringen för flammcuerna ska funka

main_win = tkinter.Tk()
main_win.geometry("0x0")
main_win.sourceFile = 'undefined'

main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir= "/", title='Välj Finale-filen')


def finaleToFlamecues()
    