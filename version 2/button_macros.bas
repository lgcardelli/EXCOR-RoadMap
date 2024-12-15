Attribute VB_Name = "button_macros"
Option Explicit

'the following macros will provide the logic to the buttons 


'macro naming format will be as follows: 'NameOfSheet_NameOfButton'
'name of button must match name of the shape that will house the macro

'Notes for Coding
'Entry Database Sheet Name = Sheet3(Entry DataBase)
    'Sheet3 Range = A2:M9999

'Entry List Sheet Name = Sheet2(Entry List)
    'Sheet2 Range = B5:N9999


Sub EntryList_syncButton()
'this is a helper macro that will copy over the data from the database to entry list
'this button should be removed when the app is handed over to the public
'this is here to support development of the app / testing

'variables
Dim LastRow As Long

    LastRow = Sheet3.Range("A9999").End(xlUp).Row 'this is to find the last row in the database
    Sheet2.Range("B5:N" & LastRow).Value = Sheet3.Range("A3:M" & LastRow).Value 'this copies the data from the DB to the Entry list 

End Sub