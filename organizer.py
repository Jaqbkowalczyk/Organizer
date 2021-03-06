# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import logging
from tkinter import *  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from typing import List, Any
from shutil import copy2
import datetime
import openpyxl
from openpyxl.styles import Font
from openpyxl.comments import Comment
import codecs
import csv
import itertools as it

BUDGET_EXCEL_FILE = 'Budżet Domowy 2021.xlsx'
#BUDGET_EXCEL_FILE = 'test.xlsx'
KEYWORDS_FILE = 'D:\\Python\\Organizator Wydatków\\keywords.csv'
CATEGORIES_FILE = 'D:\\Python\\Organizator Wydatków\\categories.csv'
RAPORT_FILE = 'Raport_'
logging.basicConfig(level=logging.CRITICAL, format='%(asctime)s - %(levelname)s - %(message)s')
os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


# logging.disable(logging.CRITICAL)


def openfile():
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    return filename


def readexcelfile(filename):
    workbook = openpyxl.load_workbook(filename)
    sheets = workbook.sheetnames
    return workbook, sheets


def makecategoriesfile(filename):
    workbook, sheets = readexcelfile(filename)
    categories = workbook.get_sheet_by_name(sheets[1])
    os.chdir('D:\\Python\\Organizator Wydatków')
    file = open(CATEGORIES_FILE, 'w')
    for i in it.chain(range(15, 30), range(36, 46), range(48, 58), range(60, 70), range(72, 82), range(84, 94),
                      range(96, 106), range(108, 118), range(120, 130), range(132, 142), range(144, 154),
                      range(156, 166), range(168, 178), range(180, 190), range(192, 202), range(204, 214)):
        position = 'B' + str(i)
        cell = categories[position]
        if cell.value not in ['.', '']:
            if i < 30:
                file.write(str(i + 43) + ',' + cell.value + '\n')
                # logging.debug(f'Dodajemy 43 do {position} wartość komórki: {cell.value}, nowa komórka: B{str(i + 43)}')
            else:
                file.write(str(i + 44) + ',' + cell.value + '\n')
                # logging.debug(f'Dodajemy 44 do {position} wartość komórki: {cell.value}, nowa komórka: B{str(i + 44)}')
    file.close()
    os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


def makekeywordsfile(filename):
    workbook, sheets = readexcelfile(filename)
    categories = workbook.get_sheet_by_name(sheets[1])
    os.chdir('D:\\Python\\Organizator Wydatków')
    file = open(KEYWORDS_FILE, 'w')
    for i in it.chain(range(15, 30), range(36, 46), range(48, 58), range(60, 70), range(72, 82), range(84, 94),
                      range(96, 106), range(108, 118), range(120, 130), range(132, 142), range(144, 154),
                      range(156, 166), range(168, 178), range(180, 190), range(192, 202), range(204, 214)):
        position = 'B' + str(i)
        cell = categories[position]
        if cell.value not in ['.', '']:
            if i < 30:
                file.write(str(i + 43) + '\n')
                logging.debug(f'Dodajemy 43 do {position} wartość komórki: {cell.value}, nowa komórka: B{str(i + 43)}')
            else:
                file.write(str(i + 44) + '\n')
                logging.debug(f'Dodajemy 44 do {position} wartość komórki: {cell.value}, nowa komórka: B{str(i + 44)}')
    file.close()
    os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


def readcsvfile(filename, **kwargs):
    # function opens csv file in encoding dependant of bank, if bank is None it opens file with standard utf-8 encoding
    filename = filename
    if 'bank' in kwargs:
        if kwargs['bank'] == 'mbank':
            with codecs.open(filename, 'r', 'ansi') as data:
                csv_data = csv.reader(data, delimiter=';')
                data_list = list(csv_data)
        elif kwargs['bank'] == 'millennium':
            with codecs.open(filename, 'r', 'utf-8-sig') as data:
                text = data.read()
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(text)
                data.seek(0)
                csv_data = csv.reader(data, delimiter=dialect.delimiter)
                data_list = list(csv_data)
        elif kwargs['bank'] == 'santander':
            with codecs.open(filename, 'r', 'utf-8') as data:
                text = data.read()
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(text)
                data.seek(0)
                csv_data = csv.reader(data, delimiter=dialect.delimiter)
                data_list = list(csv_data)
        else:
            pass
    elif 'encoding' in kwargs:
        if kwargs['encoding'] == 'ansi':
            with codecs.open(filename, 'r', 'ansi') as data:
                csv_data = csv.reader(data)
                data_list = list(csv_data)
        if kwargs['encoding'] == 'utf-8':
            with codecs.open(filename, 'r', 'utf-8') as data:
                csv_data = csv.reader(data)
                data_list = list(csv_data)
        if kwargs['encoding'] == 'utf-8-sig':
            with codecs.open(filename, 'r', 'utf-8-sig') as data:
                csv_data = csv.reader(data)
                data_list = list(csv_data)
    else:
        with codecs.open(filename, 'r', 'utf-8') as data:
            csv_data = csv.reader(data)
            data_list = list(csv_data)
    return data_list


def checkbank(filename):
    # function that checks the bank name from folder tree
    folderpath = os.path.dirname(filename)
    bank = os.path.basename(folderpath).lower()
    logging.debug(f'Bank name from folder: {bank}')
    return bank


def loopthroughcategories(title, positive=False):
    # function that checks each category and returns right one for the expense (which has the most keywords matching)
    category = -1
    category_list = []
    keywords = codecs.open(KEYWORDS_FILE, 'r', 'utf-8')
    keywords_data = csv.reader(keywords)
    keywords_list = list(keywords_data)
    for row in keywords_list:
        if positive:
            if int(row[0]) > 65:
                break
        for i in range(1, len(row)):
            if row[i].lower() in title:
                category_list.append(row[0])
            else:
                continue
    cat_set = set(category_list)
    temp = 0

    for cat in cat_set:
        cat_number = category_list.count(cat)
        if cat_number > temp:
            highest_category = cat
            temp = cat_number
        else:
            continue
        category = int(highest_category)
    return category


def checkandupdatecell(workbook, sheetname, cell, value, title, bankname=None):
    if cell[1] == '-' or cell[2] == '-':
        logging.debug('No cell - cell not updated')
        return None
    else:
        sheet = workbook[sheetname]
        cellobj = sheet[cell]
        cellvalue = cellobj.value
        logging.debug(f'Sheet: {sheetname} Cell: {cell} Value: {cellvalue}')
        if cellvalue is not None:
            if str(cellvalue)[0] == '=':
                values = cellvalue.split('+')
                if str(value) in values:
                    pass
                else:
                    cellvalue += '+' + str(value)
            else:
                if value == cellvalue:
                    pass
                else:
                    cellvalue = '=' + str(cellvalue) + '+' + str(value)
        else:
            cellvalue = value
        sheet[cell] = cellvalue
        cellobj = sheet[cell]
        if bankname == 'mbank':
            cellobj.font = Font(color="00FF0000")
        elif bankname == 'millennium':
            cellobj.font = Font(color="00660066")
        elif bankname == 'santander':
            cellobj.font = Font(color="00008080")
        else:
            pass
        comment = Comment(title, bankname)
        if cellobj.comment:
            if cellobj.comment == comment:
                pass
            else:
                comment = Comment(cellobj.comment.text + '\n' + comment.text, bankname)
                comment.width = 200
                comment.height = 300
                cellobj.comment = comment
        else:
            cellobj.comment = comment

    logging.debug(f'Sheet: {sheetname} Cell: {cell} New value: {cellvalue}')


def click_ok():
    i = 0
    chooseall = True
    for log in options:
        (click_date, click_title, click_sheet, click_cell, click_value) = log
        print(vars[i].get())
        for row in data_list:
            if len(row) == 1:
                if vars[i].get() == row[0]:
                    chooseall = False
                    logging.debug("Brak wyboru")
            elif len(row) > 1:
                if vars[i].get() == row[1]:
                    if len(click_cell) == 4:
                        new_cell = click_cell[0] + click_cell[1] + row[0]
                    else:
                        new_cell = click_cell[0] + row[0]
                    logging.debug(f'Tytuł: {click_title}, Numer Komórki: {new_cell}')
                    checkandupdatecell(workbook, click_sheet, new_cell, abs(click_value), click_title, bankname=bank)
                    raport_log.append((click_date, click_title, click_sheet, new_cell, click_value))
        i += 1
    if not chooseall:
        popup_window()
    else:
        shutdown(root)


def click_exit():
    popup_window(exit=True)


def shutdown(event):
    event.destroy()
    raport()
    workbook.save(BUDGET_EXCEL_FILE)
    sys.exit()


def popup_window(exit=False):
    window = Toplevel()
    if exit:
        label_text = 'Not everything has been saved in Excel, do you want to continue?'
    else:
        label_text = "Not everything is checked, do you want to continue?" \
                     "\n(You have to fill values manually in Excel file.)"
    label = Label(window, text=label_text)
    label.pack(fill='x', padx=50, pady=5)
    button_close = Button(window, text="Yes, I'll do it myself.", command=lambda: shutdown(window))
    button_close.pack(fill='x')
    button_back = Button(window, text="No, go back.", command=window.destroy)
    button_back.pack(fill='x')


def raport():
    category = ''
    categories_dict = {}
    os.chdir('D:\\Python\\Organizator Wydatków\\Raporty')
    text = '\t\tRaport z dnia ' + date_part +' dla banku: ' + bank + '\n' + '-'*60 + '\n'
    raport_filename = RAPORT_FILE + date_part + '_' + bank + '.csv'
    raport_path = os.getcwd() + raport_filename
    if os.path.isfile(raport_path):
        raport_filename = RAPORT_FILE + date_part + '_' + bank + '2' + '.csv'
    for cat_row in data_list:
        if len(cat_row) > 1:
            categories_dict[cat_row[0]] = cat_row[1]
        elif len(cat_row) == 1:
            continue
    logging.debug(categories_dict)
    file = open(raport_filename, 'w')
    for log in raport_log:
        (date, title, sheet, cell, value) = log
        for index, letter in enumerate(cell):
            if letter.isdigit():
                cat_num = cell[index:]
                category = categories_dict[cat_num]
                break
            else:
                pass
        if f'Data: {date}, Tytuł: {title}, ' \
           f'Arkusz: {sheet}, Komórka: {cell}, Kategoria: {category}, Wartość: {value}' in text:
            pass
        else:
            text += f'Data: {date}, Tytuł: {title}, ' \
                         f'Arkusz: {sheet}, Komórka: {cell}, Kategoria: {category}, Wartość: {value}' + '\n'
        logging.critical(f'Transakcja zapisana w raporcie: Data: {date}, Tytuł: {title}, '
                         f'Arkusz: {sheet}, Komórka: {cell}, Kategoria: {category}, Wartość: {value}')
    file.write(text)
    file.close()
    os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


class BankTrialBalance:
    data_list: List[Any]

    def __init__(self, filename, bank, workbook, sheets):
        self.filename = filename
        self.bank = bank
        self.workbook = workbook
        self.sheets = sheets
        self.data_list = []
        pass

    def datafromcsv(self):
        if self.bank == 'mbank':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, bank=self.bank)
            logging.debug(f'Data list: {self.data_list}')
        elif self.bank == 'santander':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, bank=self.bank)
        elif self.bank == 'millennium':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, bank=self.bank)
            logging.debug(f'Data list: {self.data_list}')
        else:
            print('Please specify banking account')

    def sheetandcellfromdate(self, month, day):
        sheet: str
        columnlist = [0, 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA',
                  'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM']
        if month == 1:
            sheet = 'Styczeń'
        elif month == 2:
            sheet = 'Luty'
        elif month == 3:
            sheet = 'Marzec'
        elif month == 4:
            sheet = 'Kwiecień'
        elif month == 5:
            sheet = 'Maj'
        elif month == 6:
            sheet = 'Czerwiec'
        elif month == 7:
            sheet = 'Lipiec'
        elif month == 8:
            sheet = 'Sierpień'
        elif month == 9:
            sheet = 'Wrzesień'
        elif month == 10:
            sheet = 'Październik'
        elif month == 11:
            sheet = 'Listopad'
        elif month == 12:
            sheet = 'Grudzień'
        else:
            sheet = ''

        return sheet, columnlist[day]


    def categorize(self):
        options = []
        if self.bank == 'mbank':
            start = False
            valueindex = 0
            titleindex = 0
            dateindex = 0
            month = 0
            day = 0
            cell = ''
            for index, row in enumerate(self.data_list):
                if not start:
                    if len(self.data_list[index]) > 0:
                        if self.data_list[index][0] == '#Data operacji':
                            start = True
                            for i in range(len(self.data_list[index])):
                                if self.data_list[index][i] == '#Kwota':
                                    valueindex = i
                                elif self.data_list[index][i] == '#Tytuł' or self.data_list[index][
                                    i] == '#Opis operacji':
                                    titleindex = i
                                elif self.data_list[index][i] == '#Data operacji':
                                    dateindex = i
                                else:
                                    continue
                        else:
                            continue
                    else:
                        continue
                else:
                    if len(self.data_list[index]) > 0:
                        title = self.data_list[index][titleindex].lower()
                        date = self.data_list[index][dateindex]
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace("PLN", "")
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace(",", ".")
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace(" ", "")
                        value = float(self.data_list[index][valueindex])
                        if self.data_list[index][titleindex-1].lower() == 'przelew na twoje cele':
                            category = 212
                        else:
                            if value >= 0:
                                category = loopthroughcategories(title, positive=True)
                            else:
                                category = loopthroughcategories(title)
                        month = int(date.split('-')[1])
                        day = int(date.split('-')[2])
                        sheet, column = self.sheetandcellfromdate(month, day)
                        cell = column + str(category)
                        checkandupdatecell(self.workbook, sheet, cell, abs(value), title, bankname=bank)
                        if category == -1:
                            value = float(self.data_list[index][valueindex])
                            options.append((date, title, sheet, cell, value))
                        else:
                            raport_log.append((date, title, sheet, cell, value))
                        logging.debug(f'Category: {category}, Cell: {cell}, Value: {value} Title {title} Date {date} '
                                      f'Month {month}')
                    else:
                        start = False
                        continue
        elif self.bank == 'santander':
            for index, row in enumerate(self.data_list):
                if len(self.data_list[index]) > 0 and index > 0:
                    title = self.data_list[index][2].lower()
                    date = self.data_list[index][1]
                    if len(self.data_list[index][5]) > 0:
                        self.data_list[index][5] = self.data_list[index][5].replace("\"", "")
                        self.data_list[index][5] = self.data_list[index][5].replace(",", ".")
                        value = float(self.data_list[index][5])
                        if value >= 0:
                            category = loopthroughcategories(title, positive=True)
                        else:
                            category = loopthroughcategories(title)
                    else:
                        self.data_list[index][6] = self.data_list[index][6].replace("\"", "")
                        self.data_list[index][6] = self.data_list[index][6].replace(",", ".")
                        value = float(self.data_list[index][6])
                        if value >= 0:
                            category = loopthroughcategories(title, positive=True)
                        else:
                            category = loopthroughcategories(title)
                    month = int(date.split('-')[1])
                    day = int(date.split('-')[0])
                    sheet, column = self.sheetandcellfromdate(month, day)
                    cell = column + str(category)
                    checkandupdatecell(self.workbook, sheet, cell, abs(value), title, bankname=bank)
                    if category == -1:
                        options.append((date, title, sheet, cell, value))
                    else:
                        raport_log.append((date, title, sheet, cell, value))
                    logging.debug(f'Category: {category}, Cell: {cell}, Value: {value} Title {title} Date {date} '
                                  f'Month {month}')
                else:
                    continue
        elif self.bank == 'millennium':
            for index, row in enumerate(self.data_list):
                for i in range(len(self.data_list[index])):
                    self.data_list[index][i] = self.data_list[index][i].replace("\"", "")
                if len(self.data_list[index]) > 0 and index > 0:
                    title = self.data_list[index][6].lower()
                    desc = self.data_list[index][5].lower()
                    if title == '' or title == ',':
                        title = desc
                    if len(self.data_list[index][7]) > 0:
                        self.data_list[index][7] = self.data_list[index][7].replace(",", ".")
                        value = float(self.data_list[index][7])
                        if value >= 0:
                            category = loopthroughcategories(title, positive=True)
                            category2 = loopthroughcategories(desc, positive=True)
                            if category != category2:
                                if category == -1:
                                    category = category2
                                    title = desc
                        else:
                            category = loopthroughcategories(title, positive=False)
                            category2 = loopthroughcategories(desc, positive=False)
                            if category != category2:
                                if category == -1:
                                    category = category2
                                    title = desc
                    else:
                        self.data_list[index][8] = self.data_list[index][8].replace(",", ".")
                        value = float(self.data_list[index][8])
                        if value >= 0:
                            category = loopthroughcategories(title, positive=True)
                            category2 = loopthroughcategories(desc, positive=True)
                            if category != category2:
                                if category == -1:
                                    category = category2
                                    title = desc
                        else:
                            category = loopthroughcategories(title, positive=False)
                            category2 = loopthroughcategories(desc, positive=False)
                            if category != category2:
                                if category == -1:
                                    category = category2
                                    title = desc
                    date = self.data_list[index][1]
                    month = int(date.split('-')[1])
                    day = int(date.split('-')[2])
                    sheet, column = self.sheetandcellfromdate(month, day)
                    cell = column + str(category)
                    checkandupdatecell(self.workbook, sheet, cell, abs(value), title, bankname=bank)
                    if category == -1:
                        options.append((date, title, sheet, cell, value))
                    else:
                        raport_log.append((date, title, sheet, cell, value))
                    logging.debug(f'Category: {category}, Cell: {cell}, Value: {value}, Description {desc}, '
                                  f'Title {title}, Date {date}, Month {month}')
                else:
                    continue
        self.workbook.save('example.xlsx')
        return options


class FullScreenApp(object):
    def __init__(self, master, **kwargs):
        self.master=master
        pad = 3
        self._geom = '200x200+0+0'
        master.geometry("{0}x{1}+0+0".format(
            master.winfo_screenwidth()-pad, master.winfo_screenheight()-pad))
        master.bind('<Escape>',self.toggle_geom)
    def toggle_geom(self,event):
        geom = self.master.winfo_geometry()
        print(geom, self._geom)
        self.master.geometry(self._geom)
        self._geom = geom


if __name__ == '__main__':
    dropdownlist = []
    labels = []
    dropdowns = []
    vars =[]
    raport_log = []
    i = 0
    os.chdir('D:\\Python\\Organizator Wydatków')
    root = Tk()
    root.iconbitmap('ico\\kasa.ico')
    os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')
    today = datetime.date.today()
    date_part = today.strftime("%d_%m_%Y")
    backup_path = 'C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse\\Backup\\' + date_part + '_' + BUDGET_EXCEL_FILE
    if os.path.isfile(backup_path):
        print('Backup already made today!')
        pass
    else:
        copy2(BUDGET_EXCEL_FILE, backup_path)
    app = FullScreenApp(root)
    root.title('Organizer Wydatków')
    root.resizable(height=None, width=None)
    lab = Label(root, text='Manualny wybór kategorii')
    lab.grid(row=0, column=0)
    data = readcsvfile(CATEGORIES_FILE, encoding='ansi')
    data_list = list(data)
    for row in data_list:
        if len(row) > 1:
            dropdownlist.append(row[1])
        elif len(row) == 1:
            dropdownlist.append(row[0])
    filename = openfile()
    workbook, sheets = readexcelfile(BUDGET_EXCEL_FILE)
    bank = checkbank(filename)
    # makekeywordsfile(BUDGET_EXCEL_FILE) # !!! uncommenting this will overwrite existing keywords file !!!
    # makecategoriesfile(BUDGET_EXCEL_FILE)
    trialbalance = BankTrialBalance(filename, bank, workbook, sheets)
    trialbalance.datafromcsv()
    logging.debug(f'Sheets in workbook: {sheets}')
    options = trialbalance.categorize()
    for log in options:
        i += 1
        variable = StringVar()
        variable.set(dropdownlist[0])
        vars.append(variable)
        (date, title, sheet, cell, value) = log
        labels.append(Label(root, text=f'Data: {date}, Tytuł: {title}, Wartość: {value}'))
        labels[i - 1].grid(row=i, column=0)
        dropdowns.append(OptionMenu(root, variable, *dropdownlist))
        dropdowns[i - 1].config(width=35)
        dropdowns[i - 1].grid(row=i, column=1)
        logging.debug(f'Data: {date}, Tytuł: {title}, Miesiąc: {sheet}, Komórka: {cell} Wartość: {value}')
    logging.debug(data_list)
    button_ok = Button(root, text="OK", command=click_ok)
    button_ok.grid(sticky=E, row=i+2, column=2)
    button_ok.config(width=15)
    button_exit = Button(root, text="Exit", command=click_exit)
    button_exit.grid(sticky=E, row=i + 2, column=3)
    button_exit.config(width=15)
    root.mainloop()
