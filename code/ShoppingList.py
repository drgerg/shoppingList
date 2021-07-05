#!/usr/bin/env python3
# 
# ShoppingList.py - 2021 by Gregory A. Sanders (dr.gerg@drgerg.com)
# Read the Shopping List.xlsx spreadheet and print the list to the receipt printer.
##
##

from tkinter import *
import tkinter as tk
import openpyxl as xl
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
from tkinter.font import Font
from datetime import datetime
import time,sys,csv,os, textwrap
# import pyexcel as pxl
import xlrd as pxl
import openpyxl as xl
import os, warnings, unicodedata, csv, xlrd, pathlib
from escpos import *
from escpos.printer import Network

version = "v0.8.7"

def main():
    noteText = getNotes()
    if noteText == "\n":
        text4.delete('1.0',END)
        text4.insert('1.0', 'Notes: ')
    else:
        text4.delete('1.0',END)
        text4.insert('1.0', noteText)
    window.update()
    exportFolder = os.path.expanduser("~")
    file = exportFolder + "\OneDrive\House\Shopping List.xlsx"
    colName = "Qty"
    final = makeShoppingList(file,exportFolder,colName)
    #
    ## MAKE THE STRING FOR DISPLAY AND PRINTING. SKIP 'NONE' CELLS.
    #
    finalStr = ""
    for tup in final:
        tupString = '(' + str(tup[0]) + ') '
        if tup[1] != None:
            tupString = tupString + tup[1] + ' '
        if tup[2] != None:
            tupString = tupString + tup[2] + ' '
        tupString = tupString + tup[3] + "\n"
        finalStr = finalStr + tupString
    listFrame = LabelFrame(window, text="Selected Items:")
    listFrame.grid(column=1, row=2, padx=6, sticky='w')
    #
    ##  SHOW THE LIST IN THE LEFT HAND TEXT BOX.
    #
    final_output = tk.Text(listFrame, height = 30, width = 40)
    final_output.grid(column=1, row=2, sticky='ns')
    final_output.insert('1.0',finalStr)
    final_output['state'] = 'disabled'
    text3.delete("1.0", END)
    text3.insert(INSERT, "This is your list. If it's empty, go put your quantities in the spreadsheet. Hit the Print button to print.")

    button_prn = ttk.Button(controlsFrame, text="Print", command=lambda:printIt(finalStr))                   # "Print" button
    button_prn.grid(column=0, row=7, padx=10, pady=10, sticky='n')                       # Place Print button in grid

    window.update()

def getNotes():
    noteText = text4.get('1.0', 'end')
    return noteText

def printIt(final):
    kitchen = Network("192.168.1.87") #Printer IP Address
    kitchen.set(align='center',width=2,height=2)
    kitchen.text('The List\n')
    tnow = datetime.now()
    tnow = tnow.strftime("%B %d, %Y %H:%M:%S")
    kitchen.set(align='center', width=1,height=1)
    kitchen.text(tnow + '\n\n')
    stuff = getNotes()
    stuff = textwrap.fill(stuff, width=48)
    kitchen.set(align='left')
    kitchen.text(stuff)
    kitchen.text("\n\n\n")
    kitchen.text(final)
    kitchen.text("\n\n\n")
    kitchen.cut()
    text2.delete("1.0", END)
    text2.insert("1.0", "Your list should be on the printer.")
    keepGoing = messagebox.askyesno("Hold Up.", "Exit and Keep Quantities?")
    if keepGoing == 1:
        exit()

    else:
        text2.delete("1.0", END)
        main()


def exCPcontinue(xlsxfile,exportFolder):
    ####################################################################
    ## CUSTOM SOURCE FILE: PICK COLUMN TO MATCH, COLUMNS TO PRINT     ##
    ####################################################################
    # Open the right file
    filename = xlsxfile
    folder = exportFolder
    ## these next two lines are a sad way to stop getting a
    ## warning about the lack of a default style
    with warnings.catch_warnings(record=True):  
        warnings.simplefilter("always")  
    wb1 = xl.load_workbook(filename, data_only=True)
    ws1 = wb1.worksheets[0]
    # Go to the correct column
    colName = setCtrlVals()[0]      # Get colName value.
    cnumChk = []
    for c in ws1[1]:            # Make a list of the column names in this worksheet.
        cnumChk.append(c.value)
                                ## Display the columns in a pick list so the user can pick one to match from.
        # Pick list code goes here.
    if colName in cnumChk:      # Make sure colName is in worksheet.
        for c in ws1[1]:
            if c.value == colName:
                cv = c.col_idx  # cv stores column index.
                                #
                                ## MAKE A LIST OF ROWS WITH MORE THAN 1 IN QTY.
                                #
        final = []
        for row in ws1.iter_rows(min_row=2, min_col=cv, max_col=cv):    # Check cell validity
            for cell in row:
                if cell.value is not None:
                    if cell.value != 0:
                        if cell.value != " ":
                            if cell.value >= 1:
                                final.append(cell.value)    # Add valid cell contents to cnums.
                                                                # colName NOT in WS - HOLD UP - offer to start over.
    else:
        text2.delete("1.0", END)
        text2.insert("1.0", colName + " was not found in " + filename)
        keepGoing = messagebox.askyesno("Hold Up.", "Column Name\n" + colName + "\nwas not found\n\nStart over?")
        if keepGoing == 1:
            window.mainloop()
        else:
            exit()
    return final

#
## COMPILE THE SHOPPING LIST
#
def makeShoppingList(file,exportFolder,colName):
    # Open the right file
    filename = file
    folder = exportFolder
    final = []                                  # list for final output
    ## these next two lines are a sad way to stop getting a
    ## warning about the lack of a default style
    with warnings.catch_warnings(record=True):  
        warnings.simplefilter("always")  
    wb1 = xl.load_workbook(filename, data_only=True)
    for ws1 in wb1.worksheets:
        for c in ws1[1]:            # Make a list of the column names in this worksheet.
            if c.value == colName:
                cv = c.col_idx  # cv stores column index.
        ##
        ### LOOK THROUGH THE SPREADSHEET AND APPEND EVERY ROW THAT HAS 1 OR MORE IN THE QTY COLUMN.
        ##
        for row in ws1.iter_rows(min_row=2, min_col=cv, max_col=4, values_only=True):    # Check cell validity
            cell = row[0]
            if cell is not None:
                if cell != 0:
                    if cell != " ":
                        if cell >= 1:
                            final.append(row)    # Add valid cell contents to final.
    return final

#
## CLEAR SELECTIONS FROM THE SHOPPING LIST SPREADSHEET
#
def clearTheList():
    exportFolder = os.path.expanduser("~")
    file = exportFolder + "\OneDrive\House\Shopping List.xlsx"
    colName = "Qty"
    clearSL = setCtrlVals()[1]
    if clearSL == 1:
        # Open the right file
        filename = file
        folder = exportFolder
        ## these next two lines are a sad way to stop getting a
        ## warning about the lack of a default style
        with warnings.catch_warnings(record=True):  
            warnings.simplefilter("always")  
        wb1 = xl.load_workbook(filename, data_only=True)
        for ws1 in wb1.worksheets:
            thisrow = 2
            for row in ws1.iter_rows(min_row=2, min_col=1, max_col=1):    # Check cell validity
                ws1.cell(row=thisrow, column=1).value = None
                thisrow = thisrow + 1
        # saving the destination excel file 
        wb1.save(str(file))
        text2.delete("1.0", END)
        text2.insert("1.0", "All quantities are now cleared from the spreadsheet.")
        main()
    else:
        text2.delete("1.0", END)
        text2.insert("1.0", "The checkbox was not checked, so I didn't do anything.")
        keepGoing = messagebox.askyesno("Hold Up.", "Start over?")
        if keepGoing == 1:
            text2.delete("1.0", END)
            main()
        else:
            exit()
#
## WRAP UP AND DISPLAY
#


def exit():
    sys.exit()

##
##  EVERYTHING SOUTH OF HERE IS THE 'window.mainloop' UNDEFINED BUT YET DEFINED QUASI-FUNCTION
##
window = Tk()  # Create the root window.  'root' is the common name, but I named this one 'window'.
window.title("Shopping List Data Compiler and Printer")  # Set window title
winWd = 1000  # Set window size and placement
winHt = 800
x_Left = int(window.winfo_screenwidth() / 2 - winWd / 2)
y_Top = int(window.winfo_screenheight() / 2 - winHt / 2)
window.geometry(str(winWd) + "x" + str(winHt) + "+{}+{}".format(x_Left, y_Top))
window.config(background="white")  # Set window background color
window.columnconfigure(0, weight=1)
window.columnconfigure(1, weight=1)
window.columnconfigure(2, weight=1)
window.columnconfigure(3, weight=1)
window.columnconfigure(4, weight=1)
window.columnconfigure(5, weight=1)
window.columnconfigure(6, weight=1)
window.rowconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
window.rowconfigure(2, weight=1)
window.rowconfigure(3, weight=1)
window.rowconfigure(4, weight=1)
window.rowconfigure(5, weight=1)
window.rowconfigure(6, weight=1)
label_file_explorer = Label(
    window,  # Create a File Explorer label
    text="Shopping List Data Compiler and Printer",
    width=winWd,
    font=18,
    justify="center",
    fg="green",
)
#
##
#
def featureNotReady():
    messagebox.showinfo(title='Not Yet', message='That feature is not ready.')
#
## DEFINE THE ABOUT WINDOW
#
def aboutWindow():
    aw = Toplevel(window)
    aw.title("About")
    awinWd = 400  # Set window size and placement
    awinHt = 400
    x_Left = int(window.winfo_screenwidth() / 2 - awinWd / 2)
    y_Top = int(window.winfo_screenheight() / 2 - awinHt / 2)
    aw.config(background="white")  # Set window background color
    aw.geometry(str(awinWd) + "x" + str(awinHt) + "+{}+{}".format(x_Left, y_Top))
    aw.iconbitmap('./ico/shoppinglist_icon.ico')
    awlabel = Label(aw, font=18, text ="About" + version)
    awlabel.grid(column=0, columnspan=3, row=0, sticky="n")  # Place label in grid
    aw.columnconfigure(0, weight=1)
    aw.rowconfigure(0, weight=1)
    aboutText = Text(aw, height=20, width=170, bd=3, padx=10, pady=10, wrap=WORD, font=nnFont)
    aboutText.grid(column=0, row=1)
    aboutText.insert(INSERT, "This is all under construction.\n\n- Greg Sanders\n\nThis app is written in Python and compiled using PyInstaller.\n\n"
"Check out more of my projects at www.drgerg.com.")
#
## DEFINE THE HELP WINDOW
#
def helpWindow():
    hw = Toplevel(window)
    hw.title("Help")
    hwinWd = 400  # Set window size and placement
    hwinHt = 600
    x_Left = int(window.winfo_screenwidth() / 2 - hwinWd / 2)
    y_Top = int(window.winfo_screenheight() / 2 - hwinHt / 2)
    hw.config(background="white")  # Set window background color
    hw.geometry(str(hwinWd) + "x" + str(hwinHt) + "+{}+{}".format(x_Left, y_Top))
    hw.iconbitmap('./ico/shoppinglist_icon.ico')
    hwlabel = Label(hw, height=8, font=18, text ="Help")
    hw.columnconfigure(0, weight=1)
    hw.rowconfigure(0, weight=1)
    hw.rowconfigure(1, weight=1)

    helpText = Text(hw, height=40, width=80, bd=3, padx=6, pady=6, wrap=WORD, font=nnFont)
    helpsb = ttk.Scrollbar(hw, orient='vertical', command=helpText.yview)

    helpText['yscrollcommand'] = helpsb.set
    hwlabel.grid(column=0, columnspan=3, row=0, padx=10, pady=10)  # Place label in grid
    helpText.grid(column=0, row=1)
    helpsb.grid(column=1, row=1, sticky='ns')
    helpText.insert(INSERT, "At this point, we don't really know anything.\n"
    )
#
## MENU AND MENU ITEMS
#
Frame(window)
menu = Menu(window)
window.config(menu=menu)
nnFont = Font(family="Segoe UI", size=10)          ## Set the base font
fileMenu = Menu(menu, tearoff=False)
fileMenu.add_command(label="Item")
fileMenu.add_command(label="Exit", command=exit)
menu.add_cascade(label="File", menu=fileMenu)

editMenu = Menu(menu, tearoff=False)
editMenu.add_command(label="Help", command=helpWindow)
editMenu.add_command(label="About", command=aboutWindow)
menu.add_cascade(label="Help", menu=editMenu)
#
## VARIOUS ATTEMPTS AT GETTING WINDOWS ICONS TO WORK
#
# The secret was getting the path stuff right in the pyinstaller command line: pyinstaller --add-binary=".\ico\shoppinglist_icon.ico;ico" --noconsole --icon=shoppinglist_icon.ico NextNum.py
#
window.iconbitmap('./ico/shoppinglist_icon.ico')
#
## SET UP RADIO BUTTONS FOR COLUMN NAME SELECTION
#
# "Options" frames them nicely.
#
def setCtrlVals():
    colName = "Qty"
    # custcb = custFileVar.get()
    clearSL = clearSLVar.get()
    return colName, clearSL

controlsFrame = LabelFrame(window, text="Options")             # larger frame to hold Radio Button frame
controlsFrame.grid(column=0, row=2, padx=10, sticky='nw')
colNameVar = tk.StringVar(value="CBL_NO")
rbframe = LabelFrame(controlsFrame, text="Choose What You Will")  # Frame within a frame for Radio Buttons
rbframe.grid(column=0, row=2, padx=10, pady=10, sticky='n')
#
## Set up push-buttons
#
button_clear = ttk.Button(controlsFrame, text="Clear", command=clearTheList)                  # "Exit" button
button_clear.grid(column=0, row=8, padx=10, pady=10, sticky='n')                     # Place Exit button in grid
button_exit = ttk.Button(controlsFrame, text="Exit", command=exit)                  # "Exit" button
button_exit.grid(column=0, row=9, padx=10, pady=10, sticky='n')                     # Place Exit button in grid
#
## Set up check boxes
#
# custFileVar = IntVar(value=0)
# custFilesChkBox = Checkbutton(controlsFrame,text='Select Custom File.', variable=custFileVar, onvalue=1, offvalue=0, command=setCtrlVals)      # define it
# custFilesChkBox.grid(column=0, row=3, sticky='nw')                                                   # place it
clearSLVar = IntVar(value=0)
clearSLChkBox = Checkbutton(controlsFrame,text='Clear All Qtys.', variable=clearSLVar, onvalue=1, offvalue=0, command=setCtrlVals)      # define it
clearSLChkBox.grid(column=0, row=4, sticky='nw')                                                   # place it
#
## Set up minAvail text entry box
#
T4Frame = LabelFrame(window, text="Notes and Reminders:")             # larger frame to hold Radio Button frame
T4Frame.grid(column=2, row=2, padx=6, sticky='w')
#
## Set up text windows
#
text1 = Text(window, height=6, width=150, wrap=WORD, font=nnFont)
text2 = Text(window, height=2, width=150, font=nnFont)
text3 = Text(window, height=3, width=150, font=nnFont)
text4 = Text(T4Frame, height=28, width=40, wrap=WORD, font=nnFont)

label_file_explorer.grid(column=0, columnspan=7, row=0, sticky="n")  # Place label in grid

text1.grid(column=0, columnspan=7, row=1, padx=10)
text2.grid(column=0, columnspan=7, row=6, padx=10)
text3.grid(column=0, columnspan=7, row=7, padx=10)
text4.grid(column=1, row=1)
main()
window.mainloop()  # Run the (not defined with 'def') main window loop.
