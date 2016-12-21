Sub Auto_Open()

ActiveSheet.EnableCalculation = False

' Dichiarazione variabili
Dim filtro, messaggio, fileName1, fileName2 As String
Dim targetSheet, sourceSheet As Worksheet
Dim sourceWorkbook1, sourceWorkbook2 As Workbook
Dim q As Integer

' prevengo il flickering
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' predispongo l'input
filtro = "Excel Files (*.xls; *.xlsx),*.xls;*.xlsx"
' setto la destinazione
Set targetWorkbook = Application.ActiveWorkbook

'///////////////////////////////////////
' apro il primo xls sorgente
fileName1 = Application.GetOpenFilename(filtro, , "Seleziona il primo file: ")
Set sourceWorkbook1 = Application.Workbooks.Open(fileName1)
' setto le coppie sorgente/destinazione
Set targetSheet = targetWorkbook.Worksheets(1)
Set sourceSheet = sourceWorkbook1.Worksheets(1)
' importo 1
sourceSheet.Copy targetSheet
' chiudo
sourceWorkbook1.Close

'///////////////////////////////////////
' apro il secondo xls sorgente
fileName2 = Application.GetOpenFilename(filtro, , "Seleziona il secondo file: ")
Set sourceWorkbook2 = Application.Workbooks.Open(fileName2)
' setto le coppie sorgente/destinazione
Set targetSheet = targetWorkbook.Worksheets(2)
Set sourceSheet = sourceWorkbook2.Worksheets(1)
' importo 2
sourceSheet.Copy targetSheet
' chiudo
sourceWorkbook2.Close

'///////////////////////////////////////
' SCRITTURA TABELLA

' rinomino fogli
targetWorkbook.Sheets(1).Name = "A"
targetWorkbook.Sheets(2).Name = "B"
' reset default

' setto nomi RS
Worksheets(3).Cells(2, 2) = Worksheets(1).Cells(3, 1)
Worksheets(3).Cells(2, 17) = Worksheets(2).Cells(3, 1)

' setto pageName
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 2) = Worksheets(1).Cells(21 + q, 2)
    Worksheets(3).Cells(4 + q, 17) = Worksheets(2).Cells(21 + q, 2)
Next q
    
' setto valori
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 13) = Worksheets(1).Cells(21 + q, 3)
    Worksheets(3).Cells(4 + q, 28) = Worksheets(2).Cells(21 + q, 3)
Next q

' calcolo diff tra pageName
For q = 0 To 50
    If Worksheets(3).Cells(4 + q, 2) = Worksheets(3).Cells(4 + q, 17) Then
        Worksheets(3).Cells(4 + q, 15) = True
    Else
        Worksheets(3).Cells(4 + q, 15) = False
    End If
Next q

' calcolo diff tra page views
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 30) = Worksheets(3).Cells(4 + q, 13) - Worksheets(3).Cells(4 + q, 28)
Next q

' calcolo diff percentuale tra pv
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 31) = Round(((Worksheets(3).Cells(4 + q, 30) / Worksheets(3).Cells(4 + q, 28)) * 100), 2)
Next q

'///////////////////////////////////////
' RESET DEFAULT

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox ("Fatto!")

End SubSub Auto_Open()

ActiveSheet.EnableCalculation = False

' Dichiarazione variabili
Dim filtro, messaggio, fileName1, fileName2 As String
Dim targetSheet, sourceSheet As Worksheet
Dim sourceWorkbook1, sourceWorkbook2 As Workbook
Dim q As Integer

' prevengo il flickering
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' predispongo l'input
filtro = "Excel Files (*.xls; *.xlsx),*.xls;*.xlsx"
' setto la destinazione
Set targetWorkbook = Application.ActiveWorkbook

'///////////////////////////////////////
' apro il primo xls sorgente
fileName1 = Application.GetOpenFilename(filtro, , "Seleziona il primo file: ")
Set sourceWorkbook1 = Application.Workbooks.Open(fileName1)
' setto le coppie sorgente/destinazione
Set targetSheet = targetWorkbook.Worksheets(1)
Set sourceSheet = sourceWorkbook1.Worksheets(1)
' importo 1
sourceSheet.Copy targetSheet
' chiudo
sourceWorkbook1.Close

'///////////////////////////////////////
' apro il secondo xls sorgente
fileName2 = Application.GetOpenFilename(filtro, , "Seleziona il secondo file: ")
Set sourceWorkbook2 = Application.Workbooks.Open(fileName2)
' setto le coppie sorgente/destinazione
Set targetSheet = targetWorkbook.Worksheets(2)
Set sourceSheet = sourceWorkbook2.Worksheets(1)
' importo 2
sourceSheet.Copy targetSheet
' chiudo
sourceWorkbook2.Close

'///////////////////////////////////////
' SCRITTURA TABELLA

' rinomino fogli
targetWorkbook.Sheets(1).Name = "A"
targetWorkbook.Sheets(2).Name = "B"
' reset default

' setto nomi RS
Worksheets(3).Cells(2, 2) = Worksheets(1).Cells(3, 1)
Worksheets(3).Cells(2, 17) = Worksheets(2).Cells(3, 1)

' setto pageName
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 2) = Worksheets(1).Cells(21 + q, 2)
    Worksheets(3).Cells(4 + q, 17) = Worksheets(2).Cells(21 + q, 2)
Next q
    
' setto valori
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 13) = Worksheets(1).Cells(21 + q, 3)
    Worksheets(3).Cells(4 + q, 28) = Worksheets(2).Cells(21 + q, 3)
Next q

' calcolo diff tra pageName
For q = 0 To 50
    If Worksheets(3).Cells(4 + q, 2) = Worksheets(3).Cells(4 + q, 17) Then
        Worksheets(3).Cells(4 + q, 15) = True
    Else
        Worksheets(3).Cells(4 + q, 15) = False
    End If
Next q

' calcolo diff tra page views
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 30) = Worksheets(3).Cells(4 + q, 13) - Worksheets(3).Cells(4 + q, 28)
Next q

' calcolo diff percentuale tra pv
For q = 0 To 50
    Worksheets(3).Cells(4 + q, 31) = Round(((Worksheets(3).Cells(4 + q, 30) / Worksheets(3).Cells(4 + q, 28)) * 100), 2)
Next q

'///////////////////////////////////////
' RESET DEFAULT

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox ("Fatto!")

End Sub
