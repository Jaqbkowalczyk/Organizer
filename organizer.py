# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import logging
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import openpyxl
import codecs
import csv


BUDGET_EXCEL_FILE = 'test.xlsx'

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


def readcsvfile(filename, bank):
    filename = filename
    csv_data = None
    if bank == 'mbank':
        data = codecs.open(filename, 'r', 'ansi')
        csv_data = csv.reader(data, delimiter=";")
    elif bank == 'millennium':
        data = codecs.open(filename, 'r', 'utf-8-sig')
        csv_data = csv.reader(data)
    elif bank == 'santander':
        data = codecs.open(filename, 'r', 'utf-8')
        csv_data = csv.reader(data)
    else:
        pass
    data_list = list(csv_data)
    return data_list


def checkbank(filename):
    # function that checks the bank name from folder tree
    bank = None
    folderpath = os.path.dirname(filename)
    bank = os.path.basename(folderpath).lower()
    logging.debug(f'Bank name from folder: {bank}')
    return bank

def main():
    filename = openfile()
    workbook, sheets = readexcelfile(BUDGET_EXCEL_FILE)
    bank = checkbank(filename)
    trialbalance = BankTrialBalance(filename, bank)
    trialbalance.datafromcsv()
    logging.debug(f'Sheets in workbook: {sheets}')


class BankTrialBalance:
    def __init__(self, filename, bank):
        self.filename = filename
        self.bank = bank
        self.data_list = None
        pass

    def datafromcsv(self):
        if self.bank == 'mbank':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, self.bank)
            logging.debug(f'Data list: {self.data_list}')
        elif self.bank == 'santander':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, self.bank)
        elif self.bank == 'millennium':
            logging.debug(f'Retriving data from {self.bank} csv file')
            self.data_list = readcsvfile(self.filename, self.bank)
            logging.debug(f'Data list: {self.data_list}')
        else:
            print('Please specify banking account')


if __name__ == '__main__':
    main()
