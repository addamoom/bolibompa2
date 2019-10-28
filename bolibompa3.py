import csv
import datetime
import os
from decimal import *
import re
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog, ttk
import openpyxl
from tkinter.ttk import Treeview
from tkinter.ttk import *
import ntpath
from openpyxl import Workbook
from ttkthemes import ThemedTk

# splashscreenbös
splashscreen = Tk()
splashscreen.overrideredirect(True)
splashscreen.geometry('432x432')
splashscreen.wait_visibility(splashscreen)
# splashscreen.attributes('-alpha', 0.3)
canvas = Canvas(splashscreen, height=432, width=432, bg="yellow")
canvas.pack()
image = PhotoImage(file="BombermanSigil.gif")

canvas.create_image(216, 216, image=image)

splashscreen.after(3000, splashscreen.destroy)
splashscreen.mainloop()

# width = splashscreen.winfo_screenwidth()
# height = splashscreen.winfo_screenheight()
# splashscreen.geometry('%dx%d+%d+%d' % (width*0.2, height*0.2, width*0.2, height*0.2))
# image_file = "BombermanSigil.gif"
# image = tkinter.PhotoImage(file=image_file)
# canvas = tkinter.Canvas(splashscreen, height=height*0.8, width=width*0.8, bg="brown")
# canvas.create_image(width*0.8/2, height*0.8/2, image=image)
# canvas.pack()

# splashscreen.after(2000, splashscreen.destroy)
# splashscreen.mainloop()

main_win = ThemedTk(theme="black")  # Tk()
main_win.minsize(width=800, height=300)
main_win.title("Bolibompa3")

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
ws_bulk = wb_bulk[wb_bulk.sheetnames[0]]

table = ''
folder_bulk = ''
folder_gff = ''
folder_error = ''

antal_bulk = 0
antal_gff = 0
antal_error = 0

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
        finale_file = csv.reader(finalef, delimiter='\t')
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
                    f_cell = f_row[21] + '\t' + '0' + '\t' + f_row[10] + '\t' + '1' + '\t' + '0'
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


def write_dmxcues(filepath):
    # prereq: dmxques fylld (finale_import())
    # output: csv-fil med dmxcues formaterade för lightfactory, lägger den i samma dir som finale filen
    fname = os.path.join(filepath, get_file_name(str(main_win.finale_file)))

    with open(re.sub('\.txt$', '', fname) + 'toLightfactory.csv', 'w+', newline='') as csvfile:
        fieldnames = ['namn', 'tid', 'shortcut', '?']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        for q in dmxques:
            tid = Decimal(q['tid'])
            tidfloor = int(tid)

            frames = round((tid - tidfloor) * 25)
            tidhms = str(datetime.timedelta(seconds=tidfloor))
            tidhms = tidhms + ":" + str(frames)

            writer.writerow({'namn': q['pos'] + " " + q['effekt'], 'tid': tidhms, 'shortcut': q['shortcut'], '?': ''})


def get_file_name(path):
    head, tail = ntpath.split(path)
    return tail


def search_stock(list_of_cues):
    global antal_error
    for row_cue in list_of_cues:
        if is_in_bulk(row_cue):
            continue
        elif is_in_gff(row_cue):
            continue
        else:
            row_cue[1] = '0'
            row_cue[4] = '0'
            antal_error += 1
            errors.append(row_cue)


# when reading an excel-cell, every value is read as a string, which is why they are parsed

def is_in_bulk(row_cue):
    global total_bulk, antal_bulk
    for row_bulk in ws_bulk:
        if row_cue[0] == row_bulk[0].value and row_bulk[3].value > 0:  # check if in stock
            row_bulk[3].value = int(row_bulk[3].value) - 1  # remove one from stock
            row_cue[1] = float(row_bulk[6].value)  # set price
            row_cue[4] = row_cue[1]  # set total price (used later)
            plocka_eget.append(row_cue)  # add to plocklista
            total_bulk += row_cue[1]
            antal_bulk += 1
            return True  # go back to searchthingy

    return False


def is_in_gff(row_cue):
    global total_gff, antal_gff
    for row_gff in ws_gff:
        if row_cue[0] == row_gff[3].value and row_gff[6].value > 0:  # check if in stock
            row_gff[6].value = int(row_gff[6].value) - 1  # remove one from stock
            row_cue[1] = float(row_gff[9].value)  # set the price
            row_cue[4] = row_cue[1]  # total price (used later)

            if row_gff[12].value is None:  # increase order number
                row_gff[12].value = 1
            else:
                row_gff[12].value += 1

            plocka_gff.append(row_cue)  # add the row to the plocklista
            total_gff += row_cue[1]
            antal_gff += 1
            return True  # go back to searchthingy

    return False


def sum_list(lista):
    lista.sort()
    summed_list = []
    index = 0
    summed_list.append(lista[0])        # add the first element of the old to the new
    lista.pop(0)                        # and remove it from the old
    for row in lista:                   # loop through the old list
        if row[0] == summed_list[index][0]:
            summed_list[index][3] = int(summed_list[index][3]) + 1
            summed_list[index][4] = float(summed_list[index][4]) + float(summed_list[index][1])
        else:
            summed_list[index][4] = round(summed_list[index][4], 2)
            summed_list.append(row)
            index += 1

    return summed_list


def kbk_pyro():
    if analyzed:
        varning = messagebox.askyesno("Varning", "Du kommer att skriva över filer nu, vill du fortsätta?")
        if varning:
            location = filedialog.askdirectory()
            print_list(location)
            path = os.path.join(location, 'GFF_Order.xlsx')
            wb_gff.save(path)
            wb_bulk.save('Bulklager.xlsx')

            messagebox.showinfo("Färdigt!", "Nu finns det listor. Coolt va?")
    else:
        messagebox.showinfo("Fel!", 'Klicka på "Visa Lista" först  , din smurf!')


def kbk_flames():
    if dmxques:
        path = filedialog.askdirectory()
        write_dmxcues(path)
        messagebox.showinfo("Spännande!", "Flammfilen finns nu ")


def scan_list():
    global analyzed, ign_1m, ign_5m, ign_old, plocka_eget, plocka_gff, errors
    if main_win.finale_file and gff_file:
        print('Båda filerna finns')
        # search_assortment()

        pyrocues.pop(0)
        pcs = csv.reader(pyrocues, delimiter='\t')
        search_stock(pcs)

        if plocka_eget:
            plocka_eget = sum_list(plocka_eget)
            display_lists(folder_bulk, plocka_eget)
            table.item('folder_bulk', values=['', '', antal_bulk, round(total_bulk, 2)])
        if plocka_gff:
            plocka_gff = sum_list(plocka_gff)
            display_lists(folder_gff, plocka_gff)
            table.item('folder_gff', values=['', '', antal_gff, round(total_gff, 2)])
        if errors:
            errors = sum_list(errors)
            display_lists(folder_error, errors)
            table.item('folder_error', values=['', '', antal_error, ''])
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


def add_ign():
    global total_bulk, ws_bulk, antal_bulk
    ign_1m = simpledialog.askinteger(title="", prompt="Hur många eltändare 1m (Svart)?")
    ign_5m = simpledialog.askinteger(title="", prompt="Hur många eltändare 5m  (Orange)?")
    ign_old = simpledialog.askinteger(title="", prompt="Hur många gamla eltändare?")

    i1m = ign_1m * 9
    i5m = ign_5m * 9
    iold = ign_old

    igniters = []

    # Varning för fulkod. Känsliga programmerare bör blunda
    if ign_1m > 0:
        antal_bulk += ign_1m
        igniters.append(
            ['P-IGN-1M'] + ['9'] + ['Eltändare 1m (svart)'] + [str(ign_1m)] + [str(i1m)])

    if ign_5m > 0:
        antal_bulk += ign_5m
        igniters.append(
            ['P-IGN-5M'] + ['9'] + ['Eltändare 5m (Orange)'] + [str(ign_5m)] + [str(i5m)])

    if ign_old > 0:
        antal_bulk += ign_old
        igniters.append(
            ['P-IGN-G'] + ['1'] + ['Eltändare Gamla'] + [str(ign_old)] + [str(iold)])

    for row in ws_bulk:
        if row[0].value == 'P-IGN-1M':
            row[3].value -= ign_1m
            print(row[3].value)
        elif row[0].value == 'P-IGN-5M':
            row[3].value -= ign_5m
        elif row[0].value == 'P-IGN-G':
            row[3].value -= ign_old

    if igniters:
        total_bulk += (i1m + i5m + iold)
        total_bulk = round(total_bulk, 2)
        display_lists(folder_bulk, igniters)
        plocka_eget.extend(igniters)
        print('hallå')
        table.item('folder_bulk', values=['', '', antal_bulk, total_bulk])


def re_init():
    global folder_bulk, folder_gff, folder_error, total_bulk, total_gff, pyrocues, dmxques, shortcutFile
    global plocka_eget, plocka_gff, errors, wb_gff, ws_gff, ws_bulk, analyzed, wb_bulk, ign_old, ign_1m, ign_5m

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
    ws_bulk = wb_bulk[wb_bulk.sheetnames[0]]

    analyzed = 0

    ign_1m = 0
    ign_5m = 0
    ign_old = 0
    messagebox.showinfo("Klar!", "Din session är nu rensad. Var god välj nya filer")


def init(main_win):
    global table, folder_bulk, folder_gff, folder_error, enter_factor
    info_frame = Frame(main_win)
    info_frame.pack(fill='both', expand=TRUE)

    #  info_scroll = Scrollbar(info_frame)
    #  info_scroll.pack(side=RIGHT)

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

    table.pack(side=TOP, expand=True)  # , fill=X)

    button_frame = Frame(main_win)

    bttn_add_ign = Button(button_frame, text="Lägg till eltändare", command=add_ign)
    bttn_add_ign.pack(side=TOP)

    bttn_import_finale = Button(button_frame, text="Importera Finale-TXT", command=import_finale)
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

    info_frame.pack(side=TOP, fill=Y)
    button_frame.pack(side=BOTTOM, expand=TRUE)


init(main_win)

main_win.mainloop()
