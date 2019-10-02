import csv
import datetime
from decimal import *
import re
import tkinter
from tkinter import messagebox
from tkinter import filedialog
import names
import openpyxl

main_win = tkinter.Tk()
main_win.geometry("0x0")
main_win.finaleFile = 'undefined'
main_win.shortcutFile = 'undefined'

main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj Finale-filen')
main_win.shortcutFile = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj shortcut filen')

wb_bulk = openpyxl.load_workbook(filename='../Bulklager.xlsx')
main_win.gfflager = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj GFFs lista')

wb_gff = openpyxl.load_workbook(main_win.gfflager)
pyrocues = []
dmxques = []


def finale_import():
    # prereq: finalefile, shortcutfile
    # output: pyrocues array, dmxques

    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka

    i = 0
    flag = 0
    with open(main_win.finaleFile, newline='', encoding='utf-8') as finalef:
        finalefile = csv.reader(finalef, delimiter=',')
        for f_row in finalefile:
            if len(f_row) > 2:  # sök inte igenom skabbiga tomma rader
                with open(main_win.shortcutFile, newline='', encoding='utf-8') as shorts:  # borde testa att flytta ut denna till första open statementet, nu öppnas och stängs filen jätteofta
                    sc = csv.reader(shorts, delimiter=',')
                    for sc_row in sc:
                        if sc_row[0] == f_row[14] and sc_row[1] == f_row[21]:
                            dmxques.append({
                                "tid": f_row[2],
                                "shortcut": sc_row[2]
                            })
                            flag = 1
                            print("Dmxque added")
                            break
            if flag == 0:
                #           art.nr              pris            beskrivning
                f_cell = f_row[21] + ',' + f_row[26] + ',' + f_row[10]
                pyrocues.append(f_cell)  # skulle kunna filtrera bort onödiga saker ur den här arrayen
                print("pyrocue added")
                print(len(pyrocues))
            flag = 0        #bitches love återställda flaggor

def write_dmxcues():
    # prereq: dmxques fylld (finale_import())
    # output: csv-fil med dmxcues formaterade för lightfactory, lägger den i samma dir som finale filen
    with open(re.sub('\.csv$', '', main_win.finaleFile) + 'toLightfactory.csv', 'w+', newline='') as csvfile:
        fieldnames = ['namn', 'tid', 'shortcut', '?']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        for q in dmxques:
            tid = Decimal(q['tid'])
            tidfloor = int(tid)

            frames = (tid - tidfloor) * 25
            tidhms = str(datetime.timedelta(seconds=tidfloor))
            tidhms = tidhms + ":" + str(frames)

            writer.writerow({'namn': names.get_first_name(), 'tid': tidhms, 'shortcut': q['shortcut'], '?': ''})


finale_import()
write_dmxcues()

plocka_eget = []
search_gff = []
plocka_gff = []
errors = []

def pyro_cues_to_list():

    if pyrocues:
        pyrocues.pop(0)
        pcs = csv.reader(pyrocues, delimiter=',')
        ws_bulk = wb_bulk['Blad1']
        ws_gff = wb_gff['Wholesale - Product list - Exte']
        for row_cues in pcs:
            for row_bulk in ws_bulk:
                if row_cues[0] == row_bulk[0].value:
                    if row_bulk[3].value > 0:
                        plocka_eget.append(row_cues)
                        row_bulk[3].value = row_bulk[3].value-1
                    else:
                        search_gff.append(row_cues)

        if search_gff:
            for a in search_gff:
                for row_gff in ws_gff:
                    if a[0] == row_gff[3].value:
                        if row_gff[6].value > 0:
                            plocka_gff.append(a)
                            row_gff[3].value = row_gff[6].value-1
                        else:
                            errors.append(a)




    wb_bulk.save('../Bulklager.xlsx')
    wb_gff.save('NewGFF.xlsx')



def write_plocklistor():
    with open('bulklista.txt', 'w') as bl:
        bl.write('Art.nr, Styckpris, Pjäs')
        for row_i in plocka_eget:
            bl.write('\n' + row_i[0] + ', ' + row_i[1] + ', ' + row_i[2])

    with open('gff_lista.txt', 'w') as gffl:
        gffl.write('Art.nr, Styckpris, Pjäs')
        for row_j in plocka_gff:
            gffl.write('\n' + row_j[0] + ', ' + row_j[1] + ', ' + row_j[2])

    with open('errors.txt', 'w') as errorsfile:
        errorsfile.write('Art.nr, Styckpris, Pjäs')
        for row_k in errors:
            errorsfile.write('\n' + row_k[0] + ', ' + row_k[1] + ', ' + row_k[2])

pyro_cues_to_list()
write_plocklistor()