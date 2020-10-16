import os
import logging
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import openpyxl

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


def openfile():
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    return filename


def readexcelfile():
    workbook = openpyxl.load_workbook('test.xlsx')
    sheets = workbook.sheetnames
    return workbook, sheets


def checkbank(filename):
    pass

def main():
    filename = openfile()
    workbook, sheets = readexcelfile()
    logging.debug(f'Sheets in workbook: {sheets}')  # todo set debug here from video and check sheets list`1


class BankTrialBalance(filename=None, bank=None):
    def __init__(self, filename, bank):
        self.filename = filename
        self.bank = bank
        pass

    def datafromcsv(self):
        if self.bank == 'mbank':
            pass
        elif self.bank == 'santander':
            pass
        elif self.bank == 'millenium':
            pass
        else:
            print('Please specify banking account')

    pass


if __name__ == '__main__':
    main()
