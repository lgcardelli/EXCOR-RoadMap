Attribute VB_Name = "entryForm_Macros"
Option Explicit

'variables
Dim entryRow As Long, entryCol As Long
Dim entryField As Control

Sub AddNewEntry()
'this function handles the transition from worksheet to userform
'this function is called by the addnewbutton

        'clear the selected row, if there is one
        entryForm.Range("Z1").ClearContents
        'hides the edit button
        entryForm.Shapes("editbutton").Visible = msoFalse
        'launches the User Form
        UserEntryForm.Show
    
End Sub

Sub EntryForm_EditButtonLogic()
'this logic controls what the edit button does on the worksheet

        'first make sure the user has selected a row with data just in case the button is clicked by accident
        With entryForm 'this is the worksheet where the entries live
                If .Range("Z1").Value = Empty Then 'Z1 confirms whether a row has been selected
                        MsgBox "Please select an entry to edit"
                        Exit Sub
                End If
                
        'if the user has a row selected and that row contains data, then do the following:
        entryRow = .Range("Z1").Value ' Remeber - EntryRow is a global variable and is set to the value of Z1 i.e. the current selected row
        
                With UserEntryForm 'this is calling the actual userForm
                        For entryCol = 4 To 14 'this is referencing the columns from the worksheet, the user enters data starting on col 3 i.e. type
                                Set entryField = .Controls("field" & (entryCol - 3)) 'think of this field'i' where i is the number to the title
                                entryField.Value = entryForm.Cells(entryRow, entryCol).Value 'this maps data from the worksheet to the userform
                        Next entryCol
                        .Show 'shows the user form when the looping is complete
                End With
        End With
        
End Sub

Sub EntryForm_SaveButtonLogic()
'this function controls the logic of the save button within the user form
'this will save the user inputs and emplace them on the worksheet

        'variables
        Dim prefix As String

        'first check whether the user has entered data in MVP fields
        Call EntryForm_CheckForEnteredData
        
        'determine prefix
        prefix = DeterminePrefix()
        
        With UserEntryForm
                
            'Find the last used row in column A with the same prefix and get the next incrementing number
                Dim lastRow As Long
                Dim lastEntry As String
                Dim nextNumber As Integer
        
                lastRow = entryForm.Cells(entryForm.Rows.Count, 2).End(xlUp).Row
        
                ' Loop through column A to find the highest number for the given prefix
                nextNumber = 0
                Dim i As Integer
                        For i = 1 To lastRow
                                'Left Function checks for value match
                                '.Cells(i,2) is the value in row i col b
                                '.Value 1 checks for the first value or in this case, the letter prefix
                                If Left(entryForm.Cells(i, 2).Value, 1) = prefix Then
                                        lastEntry = entryForm.Cells(i, 2).Value
                                        On Error Resume Next
                                        nextNumber = Application.WorksheetFunction.Max(nextNumber, CInt(Mid(lastEntry, 2)))
                                        On Error GoTo 0
                                End If
                        Next i
                        nextNumber = nextNumber + 1  ' Increment for the new entry

                ' Find first available row for a new entry
                        If entryForm.Range("Z1").Value = Empty Then
                                'entryRow = entryForm.Range("B9999").End(xlUp).Row + 1 'finds first available row
                                entryRow = entryForm.Cells(entryForm.Rows.Count, 2).End(xlUp).Row + 1 ' Find first available row in column B
                                entryForm.Cells(entryRow, 2).Value = prefix & nextNumber
                        Else
                                ' Code for updating an existing entry (as per your current logic)
                        End If
        
                ' Assign prefix and incremented number to column B
                'entryForm.Cells(entryRow, 2).Value = prefix & nextNumber
                
                'add the the time of entry
                If entryForm.Cells(entryRow, 3).Value = Empty Then
                        entryForm.Cells(entryRow, 3).Value = Now
                        'format the cells so the numbers read right
                        entryForm.Cells(entryRow, 3).NumberFormat = "YYYY-MM-DD HH:MM:SS"
                End If
                

                'Save other data to the database
                For entryCol = 4 To 14
                        Set entryField = .Controls("field" & entryCol - 3)
                        entryForm.Cells(entryRow, entryCol).Value = entryField.Value
                Next entryCol

                Unload UserEntryForm
                End With
                
End Sub

Sub EntryForm_DeleteButtonLogic()
        If MsgBox("Are you sure you want to delete?", vbYesNo, "Delete Product") = vbNo Then Exit Sub

        With entryForm
            If .Range("Z1").Value = Empty Then GoTo NotSaved
                    .Shapes("editbutton").Visible = msoFalse
                    entryRow = .Range("Z1").Value
                    .Range("Z1").ClearContents
                    .Range(entryRow & ":" & entryRow).EntireRow.Delete 'deletes the row
NotSaved:
                    Unload UserEntryForm
        End With

End Sub

Sub EntryForm_CheckForEnteredData()
'this function checks whether the user has entered the MVP data requirements

        With UserEntryForm
                'type
                If .field1.Value = Empty Then
                        MsgBox "Please select a type"
                        Exit Sub
                End If
                
                'title
                If .field2.Value = Empty Then
                        MsgBox "Please add a title"
                        Exit Sub
                End If
                
                'description
                If .field3.Value = Empty Then
                        MsgBox "Please add a description"
                        Exit Sub
                End If
                
                'status
                If .field5.Value = Empty Then
                        MsgBox "Please provide a status"
                        Exit Sub
                End If
                
        End With
                               
End Sub

Function DeterminePrefix() As String
'this function checks the type to determine the prefix for Col B
        
        Dim prefix As String
        
        With UserEntryForm
                Select Case .field1.Value
                        Case "Capability": prefix = "C"
                        Case "Software": prefix = "S"
                        Case "Hardware": prefix = "H"
                        Case "People": prefix = "P"
                        Case "Process": prefix = "R"
                        Case "Products": prefix = "D"
                Case Else
                        DeterminePrefix = False
            End Select
            
            'return value
            DeterminePrefix = prefix
            
        End With

End Function

Sub ForceAutoFitWrappedText()
    Dim ws As Worksheet
    Dim cell As Range

    ' Set the worksheet you want to apply the autofit on
    Set ws = entryForm

    ' Loop through each cell in the used range
    For Each cell In ws.Range("F4:F9999")
    If cell.WrapText Then
        cell.Rows.AutoFit
    End If
Next cell
End Sub
