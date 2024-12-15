Option Explicit

Sub ToggleLegendVisibility()
    Dim ws As Worksheet
    Dim legend As Shape
    
    ' Set the worksheet where the legend is located
    Set ws = ThisWorkbook.Sheets("Road Map") ' Replace "YourSheetName" with the actual sheet name
    
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
