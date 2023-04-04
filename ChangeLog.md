# ShoppingList ChangeLog

**v1.8.1** - A non-integer in the Qty column was causing the program to stop with a vague error message.  Now it tells you what the problem is, and sends you into the spreadsheet to edit.

**v1.8** - I wanted SL to generally create a list title from the filename of the database I had loaded, so now that option exists in the Configuration menu.  I tweaked some other stuff here and there, too, in order to make getting from point A to B easier.

Improved the help file a little and brought it up to date.

Made a few adjustments that, in my opinion, make things work more smoothly.

**v1.7.2** - Added a button that opens the database using the system's default app for spreadsheets. Close the database after editing, or you'll get a popup reminding you to close it. This is to avoid permissions errors. (only one app can have the spreadsheet open at one time)

**v1.7.1** - Fixed an error that printed "None" instead of ignoring "None" cells.

**v1.7** - The Notes section previously did alright with wrapping, but in the actual list itself, not so much. So now, items in the list that are longer than the paper is wide are wrapped and indented on the second and subsequent lines.

**v1.6.1-Win** - Setup for Windows in a single .exe file using Inno Setup.

This has been tested only on Windows 10 64-bit.

Be aware, I do not have a certificate with a Trusted CA (expensive!), so you may be confronted with a warning when you install ShoppingList. If you feel you don't want to risk installing it, then please don't. Otherwise, click "More info" on the "Windows protected your PC" dialog, Then the click "Run anyway" button. The installation will continue at that point.

**v1.6.1** - My main beta-tester (my wife) discovered if she put a non-string value in one of the columns other than the Qty column, SL threw a fit and refused to continue.

I fixed it.

**1.6** - Now SL optionally creates a .pdf version of your list. You can actually use SL without a receipt printer if all you want is a list to carry on your phone. The .pdf works great for that.

Leave the printer IP address set to 192.168.254.254 (the default) and when you print, the .pdf will be saved to the same folder your shopping list database .xlsx file lives in.

**v1.1** - Cleaned up code and made the interface a bit prettier.
If you are upgrading, just copy the contents of this .zip file to the same location your earlier version lived.

**v1.0** - First public release.  This has been tested on a small number of Windows machines, and seems to work just fine.