Sub sort()
Dim percentChange As Double
Dim tickerName As String
Dim yearChange As Double
Dim vol As Double
Dim yearOpen As Double
Dim yearClose As Double

yearChange = 0
percentChange = 0
vol = 0
yearOpen = 0
yearClose = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
'Calculates yearly change and percentage change--------------------------------------
    
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        yearOpen = Cells(i, 3).Value
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        yearClose = Cells(i, 6).Value
        yearChange = yearClose - yearOpen
        Range("J" & Summary_Table_Row).Value = yearChange
        percentChange = percentChange + (yearChange / yearOpen)
        Range("K" & Summary_Table_Row).Value = percentChange
    Else
        percentChange = 0
    End If
    
'Highlights positive, negative and no change-----------------------------------------

     If Range("J" & Summary_Table_Row).Value > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf Range("J" & Summary_Table_Row).Value = 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 6
    Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
'Gets ticker and volume--------------------------------------------------------------
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickerName = Cells(i, 1).Value
        vol = vol + Cells(i, 7).Value
        closePrice = closePrice + Cells(i, 6).Value
        Range("I" & Summary_Table_Row).Value = tickerName
        Range("L" & Summary_Table_Row).Value = vol
        Summary_Table_Row = Summary_Table_Row + 1
        vol = 0
        yearClose = 0
    Else
        vol = vol + Cells(i, 7).Value
    End If
    
Next i

Range("K2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "0.00%"

End Sub