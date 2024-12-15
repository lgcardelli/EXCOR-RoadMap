Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    'this makes the edit button not visible initially
    If Shapes("editbutton").Visible = True Then Shapes("editbutton").Visible = msoFalse

    'this if statement controls the highlighting of rows when the user selects a row
    If Not Intersect(Target, Range("B4:N9999")) Is Nothing And Range("B" & Target.Row).Value <> "" Then
        Range("Z1").Value = Target.Row
    
    'this will unhide the edit button
        With Shapes("editbutton")
            .Left = Range("O" & Target.Row).Left + 6
            '.Top = Range("K" & Target.Row).Top
            .Top = Range("O" & Target.Row).Top + (Range("M" & Target.Row).Height - .Height) / 2
            .Visible = msoCTrue
        End With
    Else
        Range("Z1").Value = 0
    End If
    
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TimestampColumn As Integer
    TimestampColumn = 3 ' Change to the column where you want the timestamp (e.g., 2 for Column B)

    If Not Intersect(Target, Me.Columns(2)) Is Nothing Then
        Application.EnableEvents = False
        Target.Offset(0, TimestampColumn - Target.Column).Value = Now
        Application.EnableEvents = True
    End If
End Sub



