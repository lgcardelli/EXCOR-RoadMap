Attribute VB_Name = "testing"
Option Explicit
Sub ToggleLegendVisibility()
    Dim ws As Worksheet
    Dim legend As Shape
    
    ' Set the worksheet where the legend is located
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "YourSheetName" with the actual sheet name
    
    ' Check if the grouped object "legend" exists
    On Error Resume Next
    Set legend = ws.Shapes("legend")
    On Error GoTo 0
    
    If legend Is Nothing Then
        MsgBox "The grouped object 'legend' was not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Toggle visibility
    legend.Visible = Not legend.Visible
End Sub

Sub ChangeCellColors()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    
    ' Set the worksheet where the range is located
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "YourSheetName" with the actual sheet name
    
    ' Define the range
    Set targetRange = ws.Range("N1:Z5")
    
    ' Loop through each cell in the range
    For Each cell In targetRange
        Select Case cell.Row
            Case 1
                cell.Interior.Color = RGB(0, 0, 255) ' Blue
            Case 2
                cell.Interior.Color = RGB(0, 255, 0) ' Green
            Case Else
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow
        End Select
    Next cell
End Sub

Sub ChangeNamedRangeColors()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range

    ' Set the worksheet where the named range exists
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace with your sheet name
    
    ' Reference the named range
    On Error Resume Next
    Set targetRange = ws.Range("color") ' Replace "MyColorRange" with your named range
    On Error GoTo 0

    ' Ensure the named range exists
    If targetRange Is Nothing Then
        MsgBox "Named range 'MyColorRange' not found!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Loop through each cell in the range
    For Each cell In targetRange
        Select Case cell.Row
            Case 1
                cell.Interior.Color = RGB(0, 0, 255) ' Blue
            Case 2
                cell.Interior.Color = RGB(0, 255, 0) ' Green
            Case Else
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow
        End Select
    Next cell

    MsgBox "Cell colors updated!", vbInformation, "Done"
End Sub

Sub ChangeNamedRangeColorsRelative()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim relativeRow As Long
    
    ' Set the worksheet where the named range exists
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace with your sheet name
    
    ' Reference the named range
    On Error Resume Next
    Set targetRange = ws.Range("color") ' Replace "MyColorRange" with your named range
    On Error GoTo 0

    ' Ensure the named range exists
    If targetRange Is Nothing Then
        MsgBox "Named range 'MyColorRange' not found!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Loop through each cell in the range
    For Each cell In targetRange
        ' Calculate the relative row number within the range
        relativeRow = cell.Row - targetRange.Rows(1).Row + 1
        
        ' Apply color based on the relative row
        Select Case relativeRow
            Case 1
                cell.Interior.Color = RGB(0, 0, 255) ' Blue
            Case 2
                cell.Interior.Color = RGB(0, 255, 0) ' Green
            Case Else
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow
        End Select
    Next cell

    MsgBox "Cell colors updated!", vbInformation, "Done"
End Sub




