Sub trovaUguali()
    
    ' se i valori della colonna A sono presenti nella colonna B scrive una stringa in C
    
    Dim numRighe As Integer

    numRighe = ActiveSheet.UsedRange.Rows.Count
    
    For x = 1 To numRighe
        For y = 1 To numRighe
            If Range("A" & x).Value = Range("B" & y).Value Then
                Range("C" & x).Value = "trovato"
            End If
        Next y
    Next x

    MsgBox "Esecuzione terminata"

End Sub
