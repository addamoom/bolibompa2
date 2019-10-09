import csv
import datetime
from decimal import *
import re
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog
import openpyxl
from tkinter.ttk import Treeview

main_win = Tk()
main_win.minsize(width=800, height=600)
main_win.title("Bolibompa 3.0")

main_win.finale_file = ''
gff_file = ''
create_order = 0

pyrocues = []
dmxques = []
shortcutFile = open('shortcuts.csv', 'r')


plocka_eget = []
plocka_gff = []
errors = []


wb_gff = ''
wb_bulk = openpyxl.load_workbook(filename='Bulklager.xlsx')

ws_gff = ''
ws_bulk = wb_bulk['Bulklager']


table = ''
folder_bulk = ''
folder_gff = ''
folder_error = ''

def import_finale():
    main_win.finale_file = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj Finale-filen')

    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka

    flag = 0
    with open(main_win.finale_file, newline='', encoding='utf-8') as finalef:
        finale_file = csv.reader(finalef, delimiter=',')
        for f_row in finale_file:
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
                                "effekt": f_row[21]
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
    print('färdig')


def import_gff():
    global gff_file
    global wb_gff
    global ws_gff
    gff_file = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj GFF-filen')
    wb_gff = openpyxl.load_workbook(gff_file)
    ws_gff = wb_gff['Wholesale - Product list - Exte']


def write_dmxcues():
    # prereq: dmxques fylld (finale_import())
    # output: csv-fil med dmxcues formaterade för lightfactory, lägger den i samma dir som finale filen
    with open(re.sub('\.csv$', '', main_win.finale_file) + 'toLightfactory.csv', 'w+', newline='') as csvfile:
        fieldnames = ['namn', 'tid', 'shortcut', '?']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        for q in dmxques:
            tid = Decimal(q['tid'])
            tidfloor = int(tid)

            frames = round((tid - tidfloor) * 25)
            tidhms = str(datetime.timedelta(seconds=tidfloor))
            tidhms = tidhms + ":" + str(frames)

            writer.writerow({'namn': q['pos'] + " " + q['effekt'], 'tid': tidhms, 'shortcut': q['shortcut'], '?': ''})




def search_assortment():
    nflag = 0  # används för att fylla i plocklistan
    mflag = 0  # används för att avgöra om en rad ska sökas efter i gffs lista

    pyrocues.pop(0)  # ta bort rad med rubriker

    pcs = csv.reader(pyrocues, delimiter=',')
    for row_cues in pcs:
        if row_cues[0] == 'BB':
            row_cues[1] = '0'
            plocka_eget.append(row_cues)
        else:
            for row_bulk in ws_bulk:
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
                row[1] = float(row_gff[9].value)
                plocka_gff.append(row)

            else:
                errors.append(row)
    if error_flag:
        errors.append(row)




def kbk():
    write_dmxcues()
    print('klar')


def scan_list():
    if main_win.finale_file and gff_file:
        print('Båda filerna finns')
        search_assortment()
        if plocka_eget:
            display_lists(folder_bulk, plocka_eget)
        if plocka_gff:
            display_lists(folder_gff, plocka_gff)
        if errors:
            display_lists(folder_error, errors)
    else:
        messagebox.showinfo("Varning!", "Du måste välja filer först")

def display_lists(folder, list):

    for row in list:
        table.insert(folder, "end", text=row[0], values=[row[1], row[2]])



def init(main_win):

    global table, folder_bulk, folder_gff, folder_error
    info_frame = Canvas(main_win, height=700)

    info_scroll = Scrollbar(info_frame)
    info_scroll.pack(side=RIGHT, fill=Y)

    table = Treeview(info_frame)                        #gör denna global
    table["columns"] = ("one", "two", "three")
    table.column("#0", width=150, minwidth=150, stretch=NO)
    table.column("#1", width=150, minwidth=150, stretch=NO)
    table.column("#2", width=500, minwidth=200)

    table.heading("#0", text="Art. NR", anchor=W)
    table.heading("#1", text="Pris", anchor=W)
    table.heading("#2", text="Beskrivning", anchor=W)

    # Level 1
    folder_bulk = table.insert("", 1, text="Bulklager")
    folder_gff = table.insert("", 2, text="GFF")
    folder_error = table.insert("", 3, text="Error")
    # table.insert(folder_bulk, "end", text="Pangpang", values=("13", "Den säger pang"))              #skapa metod av detta med folder som parameter
    # table.insert(folder_gff, "end", text="Pangpong", values=("14", "Den säger inte pang"))
    # table.insert(folder_error, "end", text="Laser", values=("14", "Den suger"))
    table.pack(side=TOP)  # , fill=X)


    button_frame = Frame(main_win)
    button_frame.pack(side=BOTTOM)

    bttn_import_finale = Button(button_frame, text="Importera Finale-fil", command=import_finale)
    bttn_import_finale.pack(side=LEFT)

    bttn_import_gff = Button(button_frame, text="Importera GFF-prislista", command=import_gff)
    bttn_import_gff.pack(side=LEFT)

    bttn_search_list = Button(button_frame, text="Analysera lista", command=scan_list)
    bttn_search_list.pack(side=LEFT)

    bttn_transact = Button(button_frame, text="KÖR!", command=kbk)
    bttn_transact.pack(side=RIGHT)

    global create_order
    chck_bttn_create_order = Checkbutton(button_frame, text="Skapa beställning", onvalue=1, offvalue=0)
    chck_bttn_create_order.pack(side=RIGHT)

    info_frame.pack(side=TOP)


init(main_win)

main_win.mainloop()
