ExcelSaveUTF8woBOM
==================

Saves data from active workbook in UTF-8 without Byte Order Mark (BOM)

Process:  
1. Opens save file dialog to get a path for saving the file
1. Copies the active sheet to a new workbook
1. Saves the new workbook in Unicode format to the path given
1. Closes the new workbook (thus avoiding permission conflict)
1. Converts the saved file from Unicode to UTF-8 without BOM
1. Tells user that file is saved

------------------

The request:  
> "We have to export this data in UTF-8 without a byte order mark, but we can't install any software. We do have Excel 2003..."

Required References:  
- Visual Basic For Applications
- Microsoft Excel 11.0 Object Library
- Microsoft ActiveX Data Objects 2.5 Library

If, when running, you get the following error:  
> Compile error: Variable not defined

You will have to install the references.

In Excel, press Alt-F11 or choose **Tools > Macro > Visual Basic Editor**  
Choose **Tools > References...**  
Make sure that all three of the Required References (above) are checked. Others will probably be checked; that is fine. The version number may be higher. That’s also fine.  
