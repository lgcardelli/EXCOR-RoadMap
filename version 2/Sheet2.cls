Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Call HighlightRow(Target)    
    
End Sub

Sub HighlightRow(ByVal Target As Range)
'this makes the edit button not visible initially
    'If Shapes("editbutton").Visible = True Then Shapes("editbutton").Visible = msoFalse
    'this if statement controls the highlighting of rows when the user selects a row
    If Not Intersect(Target, Range("B5:N9999")) Is Nothing And Range("B" & Target.Row).Value <> "" Then
        Range("A4").Value = Target.Row
    
    'this will unhide the edit button
        'With Shapes("editbutton")
            '.Left = Range("O" & Target.Row).Left + 6
            '.Top = Range("K" & Target.Row).Top
            '.Top = Range("O" & Target.Row).Top + (Range("M" & Target.Row).Height - .Height) / 2
            '.Visible = msoCTrue
        'End With
    Else
        Range("A4").Value = 0
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    'Call DisableCellEvents(Target)
    Call ApplyFilters(Target)

End Sub

Sub ApplyFilters(ByVal Target As Range)
'this functions handles the changes for the filtering
    If Target.CountLarge > 1 Then Exit Sub

    If Not Intersect(Target, Range("B3:N3")) Is Nothing Then

        If InStr(Target.Value, ":") = 0 Then

            If Not Intersect(Target, Range("B3")) Is Nothing Then
                Sheet3.Range("S3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("C3")) Is Nothing Then
                Sheet3.Range("T3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("D3")) Is Nothing Then
                Sheet3.Range("U3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("E3")) Is Nothing Then
                Sheet3.Range("V3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("F3")) Is Nothing Then
                Sheet3.Range("W3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("G3")) Is Nothing Then
                Sheet3.Range("X3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("H3")) Is Nothing Then
                Sheet3.Range("Y3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("I3")) Is Nothing Then
                Sheet3.Range("Z3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("J3")) Is Nothing Then
                Sheet3.Range("AA3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("K3")) Is Nothing Then
                Sheet3.Range("AB3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("L3")) Is Nothing Then
                Sheet3.Range("AC3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("M3")) Is Nothing Then
                Sheet3.Range("AD3").Value = "*" & Target.Value & "*"
            End If

            If Not Intersect(Target, Range("N3")) Is Nothing Then
                Sheet3.Range("AE3").Value = "*" & Target.Value & "*"            
            End If

        RunFilter
        End If
    End If
End Sub

Sub DisableCellEvents(ByVal Target As Range)
'prevents user from deleting data from the entry list

    If Not Intersect(Target, Range("B5:N9999")) Is Nothing Then
        If Target.Count >= 1 Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            End
        End If
    End If

End Sub