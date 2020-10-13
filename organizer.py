# import os
import logging
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
# os.chdir('C:\\Users\\jaqbk\\OneDrive\\Dokumenty\\Finanse')


def openfile():

    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    return filename


def main():
    filename = openfile()


if __name__ == '__main__':
    main()
