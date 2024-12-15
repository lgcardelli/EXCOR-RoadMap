Attribute VB_Name = "filter_macros"

Option Explicit

Sub RunFilter()

    Dim LastRow As Long, LastCriteriaRow As Long, LastResultRow As Long

    Sheet2.Range("B5:N9999").ClearContents 'this clears data from entry list

    With Sheet3
        LastCriteriaRow = 3
        LastRow = .Range("A9999").End(xlUp).Row
        .Range("A2:M" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("S2:AE" & LastCriteriaRow), CopyToRange:=.Range("AI2:AU2"), Unique:= True

         LastResultRow = .Range("AI9999").End(xlUp).Row
         If LastResultRow < 3 Then GoTo NoResults
            Sheet2.Range("B5:N" & LastResultRow + 2).Value = .Range("AI3:AU" & LastResultRow).Value

        NoResults:

    End With
End Sub


Sub ClearFilter()

Dim LastRow As Long

    LastRow = Sheet3.Range("A9999").End(xlUp).Row

    Sheet2.Range("B5:N" & LastRow + 2).Value = Sheet3.Range("A3:M" & LastRow).Value
    Sheet2.Range("B3:N3").Value = Sheet1.Range("Z1:AL1").Value
    Sheet3.Range("S3:AE3").ClearContents

End Sub
