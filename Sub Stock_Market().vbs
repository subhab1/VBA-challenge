Sub Stock_Market()

'Setting to run code in every worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
'Setting headers for the Summary Table for each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'Setting the summery table for part 3
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Grestest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
       
    
    
' variables:
    Dim tickername As String
    Dim tickervolume As Double
        tickervolume = 0

'Keeping track of the location for each ticker name in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    Dim open_price As Double
    open_price = ws.Cells(2, 3).Value
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double

       
'To Count the number of rows in the first column.
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through the rows by the ticker names
    For i = 2 To lastrow

'Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
    tickername = ws.Cells(i, 1).Value

    tickervolume = tickervolume + ws.Cells(i, 7).Value

'Print the ticker name in the summary table
    ws.Range("I" & summary_table_row).Value = tickername

'Print the volume  in the summary table
    ws.Range("L" & summary_table_row).Value = tickervolume
    
    
'calculation for yearly change.
    
' Setting closing price value
    close_price = ws.Cells(i, 6).Value
'open_price = ws.Cells(2, 3).Value

    yearly_change = (close_price - open_price)
              
'Print the yearly change in the summary table
    ws.Range("J" & summary_table_row).Value = yearly_change


    If (open_price = 0) Then

      
        percent_change = 0

    Else
                    
         percent_change = yearly_change / open_price
                
    End If


        ws.Range("K" & summary_table_row).Value = percent_change
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
   
'Reset the row counter. Add one to the summary_ticker_row
       summary_table_row = summary_table_row + 1

'Reset volume to zero
        tickervolume = 0

'Reset the opening price
        open_price = ws.Cells(i + 1, 3)
            
        Else
              
'Add the volume
        tickervolume = tickervolume + ws.Cells(i, 7).Value

            
        End If
        
        Next i
        
'calculation for second and third summary table
            Dim maxtickername As String
            maxtickername = " "
            Dim mintickername As String
            mintickername = " "
            Dim max_volume_name As String
            max_volume_name = "  "
            Dim max_percent As Double
            max_percent = 0
            Dim min_percent As Double
            min_percent = 0
            Dim max_volume As Double
            max_volume = 0
    
    
    
'Conditional formatting for column 9
    
    lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    

    For i = 2 To lastrow2
            If ws.Cells(i, 10).Value > 0 Then
               ws.Cells(i, 10).Interior.ColorIndex = 43
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        
            
    If (ws.Cells(i, 11).Value > max_percent) Then
      max_percent = ws.Cells(i, 11).Value
      maxtickername = ws.Cells(i, 9).Value
    ElseIf (ws.Cells(i, 11).Value < min_percent) Then
      min_percent = ws.Cells(i, 11).Value
      mintickername = ws.Cells(i, 9).Value
    
    ElseIf (ws.Cells(i, 12).Value > max_volume) Then
      max_volume = ws.Cells(i, 12).Value
      max_volume_name = ws.Cells(i, 9).Value
      
      
    
      End If
      
'print the value
      ws.Range("q2").Value = max_percent
      ws.Range("q2").NumberFormat = "0.00%"
      ws.Range("q3").Value = min_percent
      ws.Range("q3").NumberFormat = "0.00%"
      ws.Range("P2").Value = maxtickername
      ws.Range("p3").Value = mintickername
      ws.Range("q4").Value = max_volume
      ws.Range("o2").Value = "Greatest % increase"
      ws.Range("o3").Value = "Greatest % decrease"
      ws.Range("o4").Value = "Greatest Total Volume"
      ws.Range("p4").Value = max_volume_name
      
      
      
      
      
      
   Next i
    
    Next ws

End Sub
