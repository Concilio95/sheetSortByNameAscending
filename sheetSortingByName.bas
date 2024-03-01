Attribute VB_Name = "sheetSortingByName"
Sub sheetSortingByName()
    Application.ScreenUpdating = False
    Dim sheetCounter As Integer, i As Integer, j As Integer
    sheetCounter = Sheets.Count
    
    For i = 1 To sheetCounter - 1
        For j = i + 1 To sheetCounter
            If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub
