Sub VBA_challenge()

'loop through all sheets
For Each ws In Worksheets

    ' Set up titles for summary tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Value"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
      
    ' Establish variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim annual_open As Double
    Dim annual_close As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim ticker_inc As String
    Dim ticker_dec As String
    Dim ticker_vol As String
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim greatest_vol As Double
        
    
    ' Keep track of each stock the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Set i as
    Dim i As Double
    
    'Set initial annual open
    annual_open = ws.Cells(2, 3).Value
    
    ' Set intial values for min and max percent and highest total
    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0
      
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Loop through all ticker symbols
        For i = 2 To LastRow
        
         ' Check if the ticker symbol in row below is different, if it is not....
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set Ticker Symbol
        ticker = ws.Cells(i, 1).Value
        
        'Identify annual close value
        annual_close = ws.Cells(i, 6)
        
        'Add to total_volume
        total_volume = total_volume + ws.Cells(i, 7).Value
            
        ' Calculate Yearly Change by subtracting closing price at the end of the year to the opening price of start of year and the percent change
        yearly_change = annual_close - annual_open
        percent_change = yearly_change / annual_open
                
        ' Print the Ticker symbol in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker

        ' Print the Yearly Change to the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = yearly_change

        ' Print the Percent Change to the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change)
      
        ' Print the Total Stock Value to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = total_volume
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Assign annual open for next stock
        annual_open = ws.Cells(i + 1, 3).Value
    
        'Reset total volume
        total_volume = 0
           
        Else
           
        ' Keep adding the stock volume
         total_volume = total_volume + ws.Cells(i, 7).Value

      End If
      
  Next i

    ' Setting color formatting for yearly change
    
    ' Set j as integer
    Dim j As Integer
    
    ' Determine the Last Row
    LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    ' Loop through yearly change
    For j = 2 To LastRow2
    
        ' Set conditional formatting
        If ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    
        Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
    
        End If
    
    Next j
        
' Set k as integer
    Dim k As Integer
      
    'Set resuts initial value
    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0
    
    'Loop through values
    For k = 2 To LastRow2
       
        ' Finding max percent change and associated ticker
        If ws.Cells(k, 11).Value >= greatest_inc Then
        greatest_inc = ws.Cells(k, 11).Value
        ticker_inc = ws.Cells(k, 9).Value
        
        Else
        greatest_inc = greatest_inc
        
        ' Print the info symbol in the second table
        ws.Range("P2").Value = ticker_inc
        ws.Range("Q2").Value = FormatPercent(greatest_inc)
        
        End If
     
        ' Finding min percent change and associated ticker
        If ws.Cells(k, 11).Value < greatest_dec Then
        greatest_dec = ws.Cells(k, 11).Value
        ticker_dec = ws.Cells(k, 9).Value
        
        Else
        greatest_dec = greatest_dec
        
        ' Print the info symbol in the second table
        ws.Range("P3").Value = ticker_dec
        ws.Range("Q3").Value = FormatPercent(greatest_dec)
        
        End If
  
        
        ' Finding greatest total volume and associated ticker
        If ws.Cells(k, 12).Value >= greatest_vol Then
        greatest_vol = ws.Cells(k, 12).Value
        ticker_vol = ws.Cells(k, 9).Value
        
        Else
        greatest_vol = greatest_vol
        
        ' Print the info symbol in the second table
        ws.Range("P4").Value = ticker_vol
        ws.Range("Q4").Value = greatest_vol
        
        End If
        
    Next k

' Autofit to display data
  ws.Columns("I:L").AutoFit
  ws.Columns("O:Q").AutoFit

Next ws

End Sub


