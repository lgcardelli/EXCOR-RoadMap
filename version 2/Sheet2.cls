Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
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
