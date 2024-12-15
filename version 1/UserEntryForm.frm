VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserEntryForm 
   Caption         =   "ENTRY FORM"
   ClientHeight    =   11115
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13370
   OleObjectBlob   =   "UserEntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'this function is called when the userform initializes
'it ensures that the combo boxes get populated with the necessary data

        Call AddCategoriesToComboBox
        Call AddNumbersToComboBoxes
        Call AddStatusToComboBox
        
End Sub

Sub AddCategoriesToComboBox()
'this function populates the 'Type' combo box with the drop down options
'field1 is the variable name given to the combo box

        field1.AddItem "Capability"
        field1.AddItem "Hardware"
        field1.AddItem "Software"
        field1.AddItem "People"
        field1.AddItem "Process"
        field1.AddItem "Products"

End Sub

Sub AddNumbersToComboBoxes()
'this function populates the phase combo boxes with numbers

        Dim i As Integer 'counter
        Dim currentCombo As MSForms.ComboBox 'variable for current combobox
        
        'looping thow the combo boxes using the name + the counter as a check
        'these combo boxes go from field6-field11
        For i = 6 To 11
                Set currentCombo = Me.Controls("field" & i)
                currentCombo.AddItem "1"
                currentCombo.AddItem "2"
                currentCombo.AddItem "3"
                currentCombo.AddItem "4"
                currentCombo.AddItem "5"
                currentCombo.AddItem "6"
        Next i

End Sub

Sub AddStatusToComboBox()
'this function populates the status combo box with selections

        field5.AddItem "Concept"
        field5.AddItem "R&D"
        field5.AddItem "Testing"
        field5.AddItem "Fielding"
        
End Sub

Private Sub cancelButton_Click()
'this function handles the cancel button action
'this should simply unload the form

        Unload UserEntryForm

End Sub

Private Sub saveButton_Click()
        EntryForm_SaveButtonLogic
End Sub

Private Sub deleteButton_Click()
        EntryForm_DeleteButtonLogic
End Sub


