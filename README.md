# shoppingList
ShoppingList is a no-cloud solution to your shopping woes. Private, secure and convenient, ShoppingList is revolutionizing the experience of the mundane grocery or Home Depot run.

ShoppingList is written in Python. Tkinter is the Python module that creates the graphical interface.

ShoppingList ('SL' from here on) reads from an LibreOffice Calc or Excel .xlsx file. SL reads data from the first four columns.

The first column must be named 'Qty', which is short for Quantity. Maybe in future releases that will be configurable, but for now, it seems like a pretty basic category to use for this sort of thing.


## THE COOLEST PART: The Printer.

SL is intended for use with a networked 80mm ESC/POS receipt printer. There are many under a hundred bucks to choose from. No more carrying that $900.00 phone around in your hand while navigating the treacherous isles of your local grocery store.  Instead, your list is on a small piece of paper that won't break if it gets dropped.

The python-escpos module is used to interface to with the printer.  It contains support for serial and USB protocols as well as network, but network is so simple that it is currently the default transport protocol. Perhaps in the future the other protocols can be added.

THE FILE MENU: There are three options: Configure, Select Database, and Exit.  
 - Configure lets you set the IP address for the recipt printer as well as set your own title text for the top of your printed shopping list.  This is stored in ShoppingList.ini which is in thesame folder with ShoppingList.exe.
 - - When you set the title text, try to keep it less than around 20 characters long.
 - Select Database opens a Windows file selection dialog.  When you click on the file of your choice, the location of that file is saved in ShoppingList.ini.


## FUNCTIONS EXPLAINED:

 - RELOAD: Once you've selected your database file and set your printer IP address, it's time to add some stuff to your list.  In your .xlsx database (aka, spreadsheet), add the quantities in the Qty column to the things you need to buy.  If you have SL open, after saving the database file you can press the 'Reload' button and your list will be displayed in the 'Selected items' box. In this way you can keep track of what you have added to your list as you go.

 - CLEAR ALL: If you want to clear your selections from your database after printing your list, there are two steps:  1) Check the 'Clear All Qtys' checkbox, and 2) press the 'Clear All' button.  If you press the button without checking the box,your selections will NOT be cleared, and you will see a info box telling you just that.

 - PRINT: This . . . well, it prints your list.  Afterwards you will see a box asking if you want to exit or not.  If you exit, your current selections will be retained in the database.  If you want to remove them, select 'No' and go back and use the Clear All function to remove them.

You can always find the current version number in the 'About' dialog.

