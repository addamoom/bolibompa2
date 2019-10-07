import csv
import datetime
from decimal import *
import re
import tkinter
from tkinter import filedialog
from tkinter import simpledialog
import openpyxl

main_win = tkinter.Tk()
main_win.geometry("0x0")
main_win.finaleFile = 'undefined'
main_win.shortcutFile = 'undefined'

# Använd denna vid utveckling
main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj Finale-filen')
main_win.gfflager = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj GFFs lista')

# Använd denna på pyttedatorn
# main_win.finaleFile = filedialog.askopenfilename(parent=main_win, initialdir="/Users/Pyroman/Desktop", title='Välj Finale-filen')
# main_win.gfflager = filedialog.askopenfilename(parent=main_win, initialdir="/Users/Pyroman/Desktop", title='Välj GFFs lista')
shortcutFile = open('shortcuts.csv', 'r')

wb_bulk = openpyxl.load_workbook(filename='Bulklager.xlsx')

wb_gff = openpyxl.load_workbook(main_win.gfflager)

ws_bulk = wb_bulk['Bulklager']
ws_gff = wb_gff['Wholesale - Product list - Exte']

ign_1m = int(simpledialog.askstring("", "Hur många 1m-tändare(svart)?",
                                    parent=main_win))
ign_5m = int(simpledialog.askstring("", "Hur många 5m-tändare(orange)?",
                                    parent=main_win))
ign_old = int(simpledialog.askstring("", "Hur många gamla tändare?",
                                     parent=main_win))

pyrocues = []
dmxques = []


def finale_import():
    # prereq: finalefile, shortcutfile
    # output: pyrocues array, dmxques

    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka

    flag = 0
    with open(main_win.finaleFile, newline='', encoding='utf-8') as finalef:
        finalefile = csv.reader(finalef, delimiter=',')
        for f_row in finalefile:
            if len(f_row) > 2:  # sök inte igenom skabbiga tomma rader
                with open('shortcuts.csv', newline='',
                          encoding='utf-8') as shorts:  # borde testa att flytta ut denna till första open statementet, nu öppnas och stängs filen jätteofta
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
                #           art.nr              pris            beskrivning
                f_cell = f_row[21] + ',' + f_row[26] + ',' + f_row[10] + ',' + '1'
                pyrocues.append(f_cell)  # skulle kunna filtrera bort onödiga saker ur den här arrayen
                # print("pyrocue added")
                # print(len(pyrocues))
            flag = 0  # bitches love återställda flaggor


def write_dmxcues():
    # prereq: dmxques fylld (finale_import())
    # output: csv-fil med dmxcues formaterade för lightfactory, lägger den i samma dir som finale filen
    with open(re.sub('\.csv$', '', main_win.finaleFile) + 'toLightfactory.csv', 'w+', newline='') as csvfile:
        fieldnames = ['namn', 'tid', 'shortcut', '?']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        for q in dmxques:
            tid = Decimal(q['tid'])
            tidfloor = int(tid)

            frames = round((tid - tidfloor) * 25)
            tidhms = str(datetime.timedelta(seconds=tidfloor))
            tidhms = tidhms + ":" + str(frames)

            writer.writerow({'namn': q['pos'] + " " + q['effekt'], 'tid': tidhms, 'shortcut': q['shortcut'], '?': ''})


finale_import()
write_dmxcues()

plocka_eget = []
igniters_list = []
errors = []


def pyro_cues_to_list():
    nflag = 0  # används för att fylla i plocklistan
    mflag = 0  # används för att avgöra om en rad ska sökas efter i gffs lista
    if pyrocues:

        pyrocues.pop(0)             #tar bort raden med rubriker

        pcs = csv.reader(pyrocues, delimiter=',')
        for row_cues in pcs:
            for row_bulk in ws_bulk:
                if row_cues[0] == 'BB':
                    row_cues[1] = '0'
                    plocka_eget.append(row_cues)
                    break
                if row_cues[0] == row_bulk[0].value:
                    mflag = 1
                    if row_bulk[3].value > 0:
                        if plocka_eget:
                            for row_pe in plocka_eget:
                                if row_cues[0] == row_pe[0]:
                                    row_pe[3] = str(int(row_pe[3]) + 1)
                                    row_bulk[3].value = row_bulk[3].value - 1
                                    row_pe[1] = row_bulk[6].value
                                    nflag = 1
                                    break
                            if not nflag:
                                plocka_eget.append(row_cues)
                            row_bulk[3].value = row_bulk[3].value - 1
                            nflag = 0
                        else:
                            plocka_eget.append(row_cues)
                    else:
                        search_gff_lager(row_cues)
            if not mflag:
                search_gff_lager(row_cues)
            mflag = 0

    igniters_to_list()
    wb_bulk.save('../Bulklager.xlsx')
    wb_gff.save('NewGFF.xlsx')


def search_gff_lager(row):
    # om produtken finns i lista, men ej i lager
    # om produkten inte finns i lista

    error_flag = 1

    for row_gff in ws_gff:
        if row[0] == row_gff[3].value:
            error_flag = 0
            if row_gff[6].value > 0:
                if row_gff[12].value is None:
                    row_gff[12].value = 1
                else:
                    row_gff[12].value = int(row_gff[12].value) + 1
            else:
                errors.append(row)
    if error_flag:
        errors.append(row)

def write_plocklistor():
    with open('bulklista.txt', 'w') as bl:
        bl.write('Art.nr, Styckpris, Pjäs, Antal, Totalpris')
        for row_i in plocka_eget:
            print(row_i)
            bl.write('\n' + row_i[0] + ', \t ' + str(row_i[1]) + ', \t' +
                     row_i[2] + ', \t' + row_i[3] + ', \t' + str(float(row_i[1]) * float(row_i[3])))
        for row_j in igniters_list:
            bl.write('\n' + row_j)



    with open('errors.txt', 'w') as errorsfile:
        errorsfile.write('Art.nr, Styckpris, Pjäs, Antal, Totalpris')
        for row_k in errors:
            errorsfile.write('\n' + row_k[0] + ', ' + row_k[1] + ', ' + row_k[2] + ', ' + row_i[3])



def igniters_to_list():
    price_1m = 0
    price_5m = 0
    price_old = 0

    for row_bulk in ws_bulk:
        if row_bulk[0].value == 'PYROT-IGN-1M':
            row_bulk[3].value = row_bulk[3].value - ign_1m
            price_1m = row_bulk[6].value

        elif row_bulk[0].value == 'PYROT-IGN-5M':
            row_bulk[3].value = row_bulk[3].value - ign_1m
            price_5m = row_bulk[6].value

        elif row_bulk[0].value == 'PYROT-IGN-GAMLA':
            row_bulk[3].value = row_bulk[3].value - ign_1m
            price_old = row_bulk[6].value

    igniters_list.append('PYROT-IGN-1M' + ',' + str(price_1m) + ',' + 'Eltändare 1m, Svart' + ',' + str(ign_1m) + ','
                        + str(price_1m*ign_1m))
    igniters_list.append('PYROT-IGN-5M' + ',' + str(price_5m) + ',' + 'Eltändare 5m, Orange' + ',' + str(ign_5m) + ','
                        + str(price_5m*ign_5m))
    igniters_list.append('PYROT-IGN-GAMLA' + ',' + str(price_old) + ',' + 'Eltändare Gamla' + ',' + str(ign_old) + ','
                        + str(price_old*ign_old))


pyro_cues_to_list()
write_plocklistor()


#todo: totalpris från bulk och totalpris från gff