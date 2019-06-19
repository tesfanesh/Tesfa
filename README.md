# Tesfa
Sub stock()

Cells(1, 9).Value = "Ticker"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"

Dim LastRow As Long
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
' MsgBox LastRow

Dim Ticker As String
Dim Volume As Double
Dim rowInd As Integer
Dim openVal As Double
Dim closeVal As Double

Volume = 0
rowInd = 2

For i = 2 To LastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set the open price
        openVal = Cells(i, 3).Value
        
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
        ' Set the Brand name
        Ticker = Cells(i, 1).Value

        ' Add to the Brand Total
        Volume = Volume + Cells(i, 7).Value
      
        ' Set the closing price
        closeVal = Cells(i, 6).Value

        ' Print the Credit Card Brand in the Summary Table
        Range("I" & rowInd).Value = Ticker
      
        ' Print yearly change in the Summary Table
        Range("J" & rowInd).Value = closeVal - openVal
        
        If openVal = 0 Then
            Range("K" & rowInd).Value = 0
        Else
            ' Print percent change in the Summary Table
            Range("K" & rowInd).Value = (closeVal - openVal) / openVal
        End If
        
        ' Print the Brand Amount to the Summary Table
        Range("L" & rowInd).Value = Volume

        ' Add one to the summary table row
        rowInd = rowInd + 1
      
        ' Reset the Brand Total
        Volume = 0
      
        ' Reset the open price
        openVal = 0
      
        ' Reset the close price
        closeVal = 0

    ' If the cell immediately following a row is the same brand...
    Else

        ' Add to the Brand Total
        Volume = Volume + Cells(i, 7).Value

    End If

Next i
  
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "I").End(xlUp).Row
End With
'MsgBox LastRow

Columns("K").NumberFormat = "0.00%"

For i = 2 To LastRow
    
    If Cells(i, 10).Value > 0 Then
    
        Cells(i, 10).Interior.ColorIndex = 4
    
    Else
        
        Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
Next i

End Sub
