Sub alphatesting():


 For Each ws In Worksheets
  ' Set an initial variable for holding the tickername
  Dim tickername As String

  ' Set an initial variable for holding the total per ticker
  Dim totalvolume As Double
  totalvolume = 0
 
 Dim yearlychange As Double
 Dim openindex As Long
 openindex = 2
 
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      ' Print the tickername in the Summary Table
      ws.Range("i1").Value = "Ticker"

      ' Print the volume andchange to the Summary Table
      ws.Range("l1").Value = "Total Volume"
      
      ws.Range("j1").Value = "Yearly Change"

  ' Loop through all tickernames
  For i = 2 To lastRow

    ' Check if we are still within the same tickername if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
      yearlychange = ws.Cells(i, "F") - ws.Cells(openindex, "C")
      ' Set the ticker name
     tickername = ws.Cells(i, 1).Value

      ' Add to the Brand Total
      totalvolume = totalvolume + ws.Cells(i, "g").Value

      ' Print the ticker in the Summary Table
      ws.Range("i" & Summary_Table_Row).Value = tickername
      
      ws.Range("j" & Summary_Table_Row).Value = yearlychange

      ' Print the totalvolume to the Summary Table
      ws.Range("l" & Summary_Table_Row).Value = totalvolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume Total
      totalvolume = 0
    openindex = i + 1
    ' If the cell immediately following a row is the same total...
    Else

      ' Add to the volume Total
      totalvolume = totalvolume + ws.Cells(i, "g").Value

    End If

  Next i

'set the rows to be looked at
For i = 2 To lastRow
'If greater than zero make it green, if less then make it red

If ws.Cells(i, 10).Value >= 0 Then
    
    ws.Cells(i, 10).Interior.ColorIndex = 4
Else
   ws.Cells(i, 10).Interior.ColorIndex = 3

End If
 
  
Next i

 
   
    
  
    
    ' Loop through each row and calculate the percent change
    For i = 2 To lastRow
        
        ' Get the old value and new value from columns B and C
        Dim oldValue As Double
        Dim newValue As Double
        Dim percentChange As Double
        
        oldValue = ws.Cells(i, "c").Value
        newValue = ws.Cells(i, "j").Value
        
        ' Calculate the percent change
        percentChange = (newValue - oldValue) / oldValue * 100
        
        ' Write the percent change in column D
        ws.Cells(i, "k").Value = percentChange
        
    Next i
    
    ' Optional: Format the percent change column as percentage
    ws.Range("k2:k" & lastRow).NumberFormat = "0.00%"
    
    Dim word As String
    word = "Percent Change"
    ws.Range("K1").Value = word
    
Next ws

End Sub


