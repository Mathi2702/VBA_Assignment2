Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ws As Worksheet
Dim i As Long
Dim tickersymbol As String
Dim yearlychange As Double
Dim openprice As Double
Dim closeprice As Double
Dim percentchange As Double
Dim totalvolume As Double
Dim Summary_Table As Long
Dim max_value, min_value, great_vol As Double
Dim max_tckr, min_tckr, vol_tckr As String
Dim llRow As Long
Dim lRow As Long

For Each ws In ThisWorkbook.Worksheets
'Header for the summary table
 ws.Range("L1").Value2 = "Ticker"
 ws.Range("M1").Value2 = "Yearly Change"
 ws.Range("N1").Value2 = "Percent Change"
 ws.Range("O1").Value2 = "Total Stock Volume"
 'Header for Bonus work
 ws.Range("S20").Value2 = "Greatest % Increase"
 ws.Range("s21").Value2 = "Greatest% Decrease"
 ws.Range("S22").Value2 = "Greatest totalvolume"
 ws.Range("v19").Value2 = "Ticker"
 ws.Range("w19").Value2 = "Value"
 
 tickersymbol = 0
 totalvolume = 0
 Summary_Table = 2
 max_value = 0
 min_value = 0
 great_vol = 0
 max_tckr = " "
 min_tckr = " "
 lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 openprice = ws.Cells(2, 3).Value2
 
 
'for loop for the row
     For i = 2 To lRow
            
         If ws.Cells(i + 1, 1).Value2 <> ws.Cells(i, 1).Value2 Then
            'specify the value for ticker and close price
             tickersymbol = ws.Cells(i, 1).Value2
             closeprice = ws.Cells(i, 6).Value2
            'calculation for yearly change
             yearlychange = closeprice - openprice
            'percentchange calculation
             percentchange = (yearlychange / openprice)
            'format the percentchange column
             ws.Range("N" & Summary_Table).Value2 = Format(percentchange, "Percent")
            'sum the total volume
             totalvolume = totalvolume + ws.Cells(i, 7).Value
            
            'specify the cell for the summary table
             ws.Range("L" & Summary_Table).Value2 = tickersymbol
             ws.Range("M" & Summary_Table).Value2 = yearlychange
             ws.Range("N" & Summary_Table).Value2 = percentchange
             ws.Range("O" & Summary_Table).Value2 = totalvolume
            'if condition for formatting
                If (yearlychange > o) Then
                   ws.Range("M" & Summary_Table).Interior.ColorIndex = 4
                Else
                   ws.Range("M" & Summary_Table).Interior.ColorIndex = 3
                End If
                   
            'populate tickersymbol, yealry challange,percent change and total volume to summaary_table_row
             Summary_Table = Summary_Table + 1
            
            'if statement to pull Greatest % Increase and % Decrease ,Total Volume
                If (percentchange > max_value) Then
                      max_value = percentchange
                      max_tckr = tickersymbol
                ElseIf (percentchange < min_value) Then
                       min_value = percentchange
                       min_tckr = tickersymbol
                End If
                If (totalvolume > great_vol) Then
                       great_vol = totalvolume
                       vol_tckr = tickersymbol
                End If
                
             'specify the cell value for open price
              openprice = ws.Cells(i + 1, 3).Value2
             'rest the total volume
              totalvolume = 0
             
          Else
                          
             'total volume
             totalvolume = totalvolume + ws.Cells(i, 7).Value
             
          
         End If
      Next i
     
'print the value
ws.Range("V20").Value = max_tckr
ws.Range("V21").Value = min_tckr
ws.Range("V22").Value = vol_tckr
ws.Range("W20").Value = Format(max_value, "Percent")
ws.Range("W21").Value = Format(min_value, "Percent")
ws.Range("W22").Value = great_vol
 
Next ws
    
End Sub




