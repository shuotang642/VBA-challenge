Attribute VB_Name = "Module1"
Sub Stock_Checker()

Dim Stock_Name As String
Dim Vol_Total As Double
Vol_Total = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Set sht = ActiveSheet


Dim lastrow As Long
lastrow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row


For i = 2 To lastrow


  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    Stock_Name = Cells(i, 1).Value
    Vol_Total = Vol_Total + Cells(i, 7).Value
    'open - close = differences and then adds up all and devided by total row'
    Year_Change = Year_Change + ((Cells(i, 3).Value - Cells(i, 6).Value) / Summary_Table_Row)

    Range("I" & Summary_Table_Row).Value = Stock_Name
    Range("J" & Summary_Table_Row).Value = Year_Change
    Range("K" & Summary_Table_Row).Value = Percentage_Change
    Range("L" & Summary_Table_Row).Value = Vol_Total

    ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    ' Reset the Brand Total
    Vol_Total = 0

  ' If the cell immediately following a row is the same brand...
  Else

    ' Add to the Brand Total
    Vol_Total = Vol_Total + Cells(i, 7).Value

  End If

Next i

End Sub


