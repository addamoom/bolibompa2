import csv
import datetime
from decimal import *
import re
import tkinter
from tkinter import messagebox
from tkinter import filedialog


main_win = tkinter.Tk()

main_win.geometry("0x0")
main_win.finaleFile = 'undefined'
main_win.shortcutFile = 'undefined'

main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir= ".", title='Välj Finale-filen')
main_win.shortcutFile = filedialog.askopenfilename(parent=main_win, initialdir= ".", title='Välj shortcut filen')

pyrocues = []
dmxques = []


def finale_import():
    #prereq: finalefile, shortcutfile
    #output: pyrocues array, dmxques

    i = 0
    flag = 0
    with open(main_win.finaleFile, newline='', encoding='utf-8') as finalef:
        finalefile = csv.reader(finalef, delimiter=',')
        for f_row in finalefile:
            if len(f_row) > 2: #sök inte igenom skabbiga tomma rader
                with open(main_win.shortcutFile, newline='', encoding='utf-8') as shorts: #borde testa att flytta ut denna till första open statementet, nu öppnas och stängs filen jätteofta
                    sc = csv.reader(shorts, delimiter=',')
                    for sc_row in sc:
                        if sc_row[0] == f_row[14] and sc_row[1] == f_row[21]:
                            dmxques.append({
                                "tid": f_row[2],
                                "shortcut": sc_row[2],
                                "pos": f_row[14],
                                "effekt":  f_row[21]
                            })
                            flag = 1
                            print("Dmxque added")
                            break
            if flag == 0:
                pyrocues.append(f_row) #skulle kunna filtrera bort onödiga saker ur den här arrayen
                print ("pyrocue added")

def write_dmxcues():
    #prereq: dmxques fylld (finale_import())
    #output: csv-fil med dmxcues formaterade för lightfactory, lägger den i samma dir som finale filen

    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka

    with open(re.sub('\.csv$', '', main_win.finaleFile) + 'toLightfactory.csv', 'w+', newline='') as csvfile:
        fieldnames = ['namn', 'tid', 'shortcut', '?']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        for q in dmxques:
            tid = Decimal(q['tid'])
            tidfloor = int(tid)

            frames = (tid - tidfloor) * 24
            tidhms = str(datetime.timedelta(seconds=tidfloor))
            tidhms = tidhms + ":" + str(frames)

            writer.writerow({'namn': q['pos'] + " " + q['effekt'], 'tid': tidhms, 'shortcut': q['shortcut'], '?': ''})

finale_import()
write_dmxcues()






    




















