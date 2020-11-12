# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import logging
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from typing import List, Any

import openpyxl
import codecs
import csv
import itertools as it

BUDGET_EXCEL_FILE = 'test.xlsx'
KEYWORDS_FILE = 'D:\\Python\\Organizator Wydatków\\keywords.csv'
CATEGORIES_FILE = 'D:\\Python\\Organizator Wydatków\\categories.csv'

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
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
    csv_data = None
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


def loopthroughcategories(title):
    # function that checks each category and returns right one for the expense (which has the most keywords matching)
    category = -1
    category_list = []
    keywords = codecs.open(KEYWORDS_FILE, 'r', 'utf-8')
    keywords_data = csv.reader(keywords, )
    keywords_list = list(keywords_data)
    for row in keywords_list:
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


def main():
    filename = openfile()
    workbook, sheets = readexcelfile(BUDGET_EXCEL_FILE)
    bank = checkbank(filename)
    # makekeywordsfile(BUDGET_EXCEL_FILE) # !!! uncommenting this will overwrite existing keywords file !!!
    # makecategoriesfile(BUDGET_EXCEL_FILE)
    trialbalance = BankTrialBalance(filename, bank, workbook, sheets)
    trialbalance.datafromcsv()
    logging.debug(f'Sheets in workbook: {sheets}')
    trialbalance.categorize()


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

    def checkandupdatecell(self, sheet, cell, value):
        if cell[1] == '-' or cell[2] == '-':
            return None
        else:
            sheet = self.workbook.get_sheet_by_name(sheet)
            cellobj = sheet[cell]
            cellvalue = cellobj.value
            logging.debug(f'Cell: {cell} value: {cellvalue}')
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
        pass

    def  askuser(self):

        return None

    def categorize(self):
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
                        if self.data_list[index][titleindex-1].lower() == 'przelew na twoje cele':
                            category = 212
                        else:
                            category = loopthroughcategories(title)
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace("PLN", "")
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace(",", ".")
                        self.data_list[index][valueindex] = self.data_list[index][valueindex].replace(" ", "")
                        value = abs(float(self.data_list[index][valueindex]))
                        month = int(date.split('-')[1])
                        day = int(date.split('-')[2])
                        sheet, column = self.sheetandcellfromdate(month, day)
                        cell = column + str(category)
                        self.checkandupdatecell(sheet, cell, value)
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
                    category = loopthroughcategories(title)
                    if len(self.data_list[index][5]) > 0:
                        self.data_list[index][5] = self.data_list[index][5].replace("\"", "")
                        self.data_list[index][5] = self.data_list[index][5].replace(",", ".")
                        value = abs(float(self.data_list[index][5]))
                    else:
                        self.data_list[index][6] = self.data_list[index][6].replace("\"", "")
                        self.data_list[index][6] = self.data_list[index][6].replace(",", ".")
                        value = abs(float(self.data_list[index][6]))
                    month = int(date.split('-')[1])
                    day = int(date.split('-')[0])
                    sheet, column = self.sheetandcellfromdate(month, day)
                    cell = column + str(category)
                    self.checkandupdatecell(sheet, cell, value)
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
                    category = loopthroughcategories(title)
                    category2 = loopthroughcategories(desc)
                    if category != category2:
                        if category == -1:
                            category = category2
                    if len(self.data_list[index][7]) > 0:
                        self.data_list[index][7] = self.data_list[index][7].replace(",", ".")
                        value = abs(float(self.data_list[index][7]))
                    else:
                        self.data_list[index][8] = self.data_list[index][8].replace(",", ".")
                        value = abs(float(self.data_list[index][8]))
                    date = self.data_list[index][1]
                    month = int(date.split('-')[1])
                    day = int(date.split('-')[2])
                    sheet, column = self.sheetandcellfromdate(month, day)
                    cell = column + str(category)
                    self.checkandupdatecell(sheet, cell, value)
                    logging.debug(f'Category: {category}, Cell: {cell}, Value: {value}, Description {desc}, '
                                  f'Title {title}, Date {date}, Month {month}')
                else:
                    continue
        self.workbook.save('example.xlsx')


if __name__ == '__main__':
    main()
