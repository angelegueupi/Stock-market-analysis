 
Sub vba_challenge()

Dim ticker As String
   Dim yearperiod As Long
   Dim startTime As Date
   Dim highvolume As Integer
   Dim lowvolume As Integer
   Dim endTime  As Date
   Dim totalVolume As Double

Dim start_row As Long


For Each W In Worksheets

'Create a header row
   W.Cells(1, 9).Value = "Ticker"
   W.Cells(1, 10).Value = "yearly change"
   W.Cells(1, 11).Value = "percentage change"
   W.Cells(1, 12).Value = "total stock volume"
   
   
   Dim summaryTableRow As Integer
     
   summaryTableRow = 2
   start_row = 2
   'Get the number of rows to loop over
   RowCount = W.Cells(Rows.Count, "A").End(xlUp).Row
   
   ' loop from row 2 in column A out to the last row
    For Row = 2 To RowCount
    
    ' check to see if the ticker symbol changes
            If W.Cells(Row + 1, 1).Value <> W.Cells(Row, 1).Value Then
                ' if the ticker symbol changes, do ....
                ' first set the ticker symbol
                
                tickersymbol = W.Cells(Row, 1).Value
                
                
     ' add the last volume from the row
        totalVolume = totalVolume + W.Cells(Row, 7).Value
        
    ' add the ticker symbol to the I column in the summary table row
                W.Cells(summaryTableRow, 9).Value = tickersymbol
                
                If W.Cells(start_row, 3).Value = 0 Then
                For i = start_row To Row
                
                If W.Cells(i, 3).Value <> 0 Then
                start_row = i
                Exit For
                End If
                Next i
                
                End If
                
                
        'yearly change
                
        yearlyChange = W.Cells(Row, 6) - W.Cells(start_row, 3)
                
        'percent change
                
        percentchange = yearlyChange / W.Cells(start_row, 3)
                
        start_row = Row + 1
                
                
                
   ' add the yearly change to J column
               
    W.Cells(summaryTableRow, 10).Value = yearlyChange
                
    W.Cells(summaryTableRow, 10).NumberFormat = "0.00"
                
  ' add the percent change to K column
    W.Cells(summaryTableRow, 11).Value = percentchange
                
 ' formatting
 
    W.Cells(summaryTableRow, 11).NumberFormat = "0.0%"
                
    W.Cells(summaryTableRow, 12).Value = totalVolume
                
    W.Columns("L").AutoFit
                
   For i = 2 To RowCount
   

    W.Range("I1:L1").Font.Bold = True
    
    
    If W.Cells(i, 10) > 0 Then
    W.Cells(i, 10).Interior.Color = vbGreen
    
    Else
    W.Cells(i, 10).Interior.Color = vbRed
    End If
    
                
   Next i
           
                
    ' go to the next summary table row (add 1 on to the value of the summary table row)
                summaryTableRow = summaryTableRow + 1
                ' reset the ticker symbol total to 0
                totalVolume = 0
                yearlyChange = 0
                
                
            Else
                ' if the ticker symbol stays the same, do....
                ' add on to the total volume from the G column
                totalVolume = totalVolume + W.Cells(Row, 7).Value
            End If
    
    Next Row
    

Next W


End Sub
