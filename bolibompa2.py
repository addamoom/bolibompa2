import csv
import datetime
from decimal import *
import re
import tkinter
from tkinter import messagebox
from tkinter import filedialog



main_win = tkinter.Tk()
main_win.geometry("0x0")
main_win.sourceFile = 'undefined'

main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir= ".", title='Välj Finale-filen')
main_win.shortcutfile = filedialog.askopenfilename(parent=main_win, initialdir= ".", title='Välj shortcut filen')

def dmxMatchFound()


def finaleToDmxcues()
    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka
    flammquer = []
    i = 0
    with open(main_win.finaleFile, newline='', encoding='utf-8') as finalef:
        finalefile = csv.reader(finalef, delimiter=',')
        for f_row in finalefile:
            if len(f_row) > 2: #sök inte igenom skabbiga tomma rader
                with open(main_win.shortcutFile, newline='', encoding='utf-8') as shorts:
                    sc = csv.reader(shorts, delimiter=',')
                    for sc_row in sc:
                        if sc_row[0] == f_row[14] and sc_row[1] == f_row[21]:
                            dmxMatchFound()




