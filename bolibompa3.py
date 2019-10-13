import csv
import datetime
import os
from decimal import *
import re
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog
import openpyxl
from tkinter.ttk import Treeview

from openpyxl import Workbook

main_win = Tk()
main_win.minsize(width=800, height=400)
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

total_bulk = 0
total_gff = 0

enter_factor = ''

analyzed = 0

def import_finale():
    main_win.finale_file = filedialog.askopenfilename(parent=main_win, initialdir=".", title='Välj Finale-filen')

    global plocka_eget, plocka_gff, errors
    getcontext().prec = 2  # behövs för att tidskonverteringen för flammcuerna ska funka
    dmxques.clear()
    pyrocues.clear()
    plocka_eget = []
    plocka_gff = []
    errors = []

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
                if f_row[21]:
                    #           art.nr              pris            beskrivning
                    f_cell = f_row[21] + ',' + '0' + ',' + f_row[10] + ',' + '1' + ',' + '0'
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

    global total_bulk

    pcs = csv.reader(pyrocues, delimiter=',')
    for row_cues in pcs:
        if row_cues[0] == 'BB':
            row_cues[1] = '0'
            plocka_eget.append(row_cues)
        else:
            for row_bulk in ws_bulk:
                if row_cues[0] == row_bulk[0].value:  # är det rätt pjäs?
                    mflag = 1
                    if row_bulk[3].value > 0:  # finns den i lager?
                        row_cues[1] = float(row_bulk[6].value)
                        row_cues[4] = row_cues[1]
                        total_bulk = total_bulk + float(row_cues[1])
                        print(row_cues[1])
                        if plocka_eget:  # är listan tom?
                            for row_pe in plocka_eget:
                                if row_cues[0] == row_pe[0]:  # matcha artnr
                                    row_pe[3] = str(int(row_pe[3]) + 1)  # öka antalet
                                    row_bulk[3].value = row_bulk[3].value - 1  # subtrahera från listan
                                    row_pe[4] = str(int(row_pe[4]) + int(row_pe[1]))  # öka totalpris
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
    multi_flag = 0
    error_multi = 0
    global total_gff

    for row_gff in ws_gff:
        if row[0] == row_gff[3].value:
            row[1] = float(row_gff[9].value)
            row[4] = row[1]
            error_flag = 0
            if row_gff[6].value > 0:  # finns i lager?
                total_gff = total_gff + row[1]
                if row_gff[12].value is None:
                    row_gff[12].value = 1
                else:
                    row_gff[12].value = int(row_gff[12].value) + 1

                for row_plocklista in plocka_gff:
                    if row_plocklista[0] == row[0]:
                        row_plocklista[3] = str(int(row_plocklista[3]) + 1)
                        row_plocklista[4] = str(int(row_plocklista[4]) + int(row_plocklista[1]))
                        multi_flag = 1
                if not multi_flag:
                    print(row)
                    plocka_gff.append(row)
                multi_flag = 0
            else:
                error_flag = 1

    if error_flag:
        row[1] = ''
        row[4] = ''
        if errors:
            for row_error in errors:
                if row_error[0] == row[0]:
                    row_error[3] = str(int(row_error[3]) + 1)
                    error_multi = 1
            if not error_multi:
                errors.append(row)
        else:
            errors.append(row)


def kbk_pyro():
    if analyzed:
        varning = messagebox.askyesno("Varning", "Du kommer att skriva över filer nu, vill du fortsätta?")
        if varning:
            location = filedialog.askdirectory()
            print_list(location)
            path = os.path.join(location, 'GFF_Order.xlsx')
            wb_gff.save(path)
            messagebox.showinfo("Färdigt!", "Nu finns det listor. Coolt va?")
    else:
        messagebox.showinfo("Fel!", 'Klicka på "Visa Lista" först  , din smurf!')

def kbk_flames():
    if dmxques:
        write_dmxcues()
        messagebox.showinfo("Spännande!", "Flammfilen finns nu ")



def scan_list():
    global analyzed
    if main_win.finale_file and gff_file:
        print('Båda filerna finns')
        search_assortment()
        if plocka_eget:
            display_lists(folder_bulk, plocka_eget)
            print(total_bulk)
            table.item('folder_bulk', values=['', '', '', total_bulk])
        if plocka_gff:
            display_lists(folder_gff, plocka_gff)
            table.item('folder_gff', values=['', '', '', total_gff])

        if errors:
            display_lists(folder_error, errors)
        analyzed = 1
    else:
        messagebox.showinfo("Varning!", "Du måste välja filer först")


def display_lists(folder, list):
    for row in list:
        table.insert(folder, "end", text=row[0], values=[row[1], row[2], row[3], row[4]])


def print_list(location):
    wb1 = Workbook()
    path = os.path.join(location, 'plocklista.xlsx')

    ws1 = wb1.active
    ws1.title = 'Pjäser'
    ws1.column_dimensions['A'].width = 20
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 60
    ws1.column_dimensions['D'].width = 20
    ws1.column_dimensions['E'].width = 20

    ws1.append(['Art.nr', 'Enhetspris', 'Beskrivning', 'Antal', 'Totalt pris'])
    ws1.append([''])
    ws1.append(['Från Bulk'])
    if plocka_eget:
        for row in plocka_eget:
            ws1.append(row)

    ws1.append([''])
    ws1.append(['Från GFF'])
    if plocka_gff:
        for row in plocka_gff:
            ws1.append(row)

    wb1.save(filename=path)


def re_init():
    global folder_bulk, folder_gff, folder_error, total_bulk, total_gff, pyrocues, dmxques, shortcutFile
    global plocka_eget, plocka_gff, errors, wb_gff, ws_gff, ws_bulk, analyzed, wb_bulk

    total_gff = 0
    total_bulk = 0

    table.delete(*table.get_children())
    folder_bulk = table.insert("", 1, 'folder_bulk', text="Bulklager", values=['', '', '', total_bulk], tags='folder')
    folder_gff = table.insert("", 2, 'folder_gff', text="GFF", values=['', '', '', total_gff], tags='folder')
    folder_error = table.insert("", 3, 'folder_error', text="Error", tags='folder')

    table.tag_configure('folder', font='bold')

    plocka_eget = []
    plocka_gff = []
    errors = []

    shortcutFile = open('shortcuts.csv', 'r')

    wb_gff = ''
    wb_bulk = openpyxl.load_workbook(filename='Bulklager.xlsx')

    ws_gff = ''
    ws_bulk = wb_bulk['Bulklager']

    analyzed = 0

    messagebox.showinfo("Klar!", "Din session är nu rensad. Var god välj nya filer")


def init(main_win):
    global table, folder_bulk, folder_gff, folder_error, enter_factor
    info_frame = Frame(main_win, height=700)

    info_scroll = Scrollbar(info_frame)
    info_scroll.pack(side=RIGHT, fill=Y)

    table = Treeview(info_frame)  # gör denna global
    table["columns"] = ("one", "two", "three", "four", "five")
    table.column("#0", width=150, minwidth=150, stretch=NO)
    table.column("#1", width=150, minwidth=150, stretch=NO)
    table.column("#2", width=500, minwidth=200)
    table.column("#3", width=150, minwidth=50, stretch=NO)
    table.column("#4", width=150, minwidth=50, stretch=NO)

    table.heading("#0", text="Art. NR", anchor=W)
    table.heading("#1", text="Pris", anchor=W)
    table.heading("#2", text="Beskrivning", anchor=W)
    table.heading("#3", text="Antal", anchor=W)
    table.heading("#4", text="Totalpris", anchor=W)

    # Level 1
    folder_bulk = table.insert("", 1, 'folder_bulk', text="Bulklager", values=['', '', '', total_bulk], tags='folder')
    folder_gff = table.insert("", 2, 'folder_gff', text="GFF", values=['', '', '', total_gff], tags='folder')
    folder_error = table.insert("", 3, 'folder_error', text="Error", tags='folder')

    table.tag_configure('folder', font='bold')

    table.pack(side=TOP)  # , fill=X)

    button_frame = Frame(main_win)
    button_frame.pack(side=BOTTOM)

    bttn_import_finale = Button(button_frame, text="Importera Finale-csv", command=import_finale)
    bttn_import_finale.pack(side=LEFT)

    bttn_import_gff = Button(button_frame, text="Importera GFF-prislista", command=import_gff)
    bttn_import_gff.pack(side=LEFT)

    bttn_search_list = Button(button_frame, text="Visa pjäser", command=scan_list)
    bttn_search_list.pack(side=LEFT)

    bttn_transact = Button(button_frame, text="Skapa plocklista", command=kbk_pyro)
    bttn_transact.pack(side=RIGHT)

    bttn_flames = Button(button_frame, text="Skapa flammlista", command=kbk_flames)
    bttn_flames.pack(side=RIGHT)

    bttn_clear = Button(button_frame, text="Rensa", command=re_init)
    bttn_clear.pack(side=RIGHT)

    info_frame.pack(side=TOP)


init(main_win)

main_win.mainloop()
