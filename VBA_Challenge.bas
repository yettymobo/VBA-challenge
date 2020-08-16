Attribute VB_Name = "Module1"
Sub StockTicker()

For Each ws In Worksheets

    ' Set variables
    Dim WorkseetName As String
 
    Dim lRow As Long
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim OpenValue As Double
    OpenValue = 0
  
    Dim CloseValue As Double
    CloseValue = 0
   
  
    Dim FirstRow As Boolean
    FirstRow = True
    
    Dim yearlyChange As Double
    yearlyChange = 0
 
    Dim TotalVol As Double
    TotalVol = 0
    
     

    'Keep track of the location of each Ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 1

    'Add headers
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"



    'Loop through all Ticker Symbols
    For i = 2 To lRow
    
           If FirstRow = True Then
            OpenValue = ws.Cells(i, 3).Value
          End If
          
          ticker = ws.Cells(i, 1).Value
          CloseValue = ws.Cells(i, 6).Value
          TotalVol = TotalVol + ws.Cells(i, 7).Value
          FirstRow = False
        

        'Check if we are still within the same ticker, if it is not then
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            ws.Range("I" & 1 + Summary_Table_Row).Value = ticker
            
            yearlyChange = CloseValue - OpenValue
            
            If OpenValue <> 0 Then
                percentChange = yearlyChange / OpenValue
            Else
               percentChange = 0
               
            End If
    
            
            'Print Remaining Values to the Summary Table
            
            ws.Range("J" & Summary_Table_Row + 1).Value = yearlyChange
            ws.Range("K" & Summary_Table_Row + 1).Value = Format(percentChange, "Percent")
            ws.Range("L" & Summary_Table_Row + 1).Value = TotalVol
             
            
            'Color Formatting
            If yearlyChange > 0 Then
              ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
              
            Else
                
              ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            FirstRow = True
            'Reset values
            TotalVol = 0
            
            
        
                    

        End If

    Next i
    
'Hard Solution'

    'Declare Variables
    Dim GreatestPcntIncrease As Double
        GreatestPcntIncrease = 0
        
    Dim GreatestPcntDecrease As Double
        GreatestPcntDecrease = 0
        
    Dim GreatestVol As Double
        GreatestVol = 0
        
    Dim GreatestPcntIncreaseTicker As String
    
    Dim GreatestPcntDecreaseTicker As String
    
    Dim GreatestVolTicker As String

    'Compute last row for iteration
    lRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Print Headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    
    'Calculate Greatest - Volume, Pecent Increase, Percent Decrease
    For i = 2 To lRow
          ticker = ws.Cells(i, 9).Value
          percentChange = ws.Cells(i, 11).Value
          TotalVol = ws.Cells(i, 12).Value
          
          If TotalVol > GreatestVol Then
            GreatestVol = TotalVol
            GreatestVolTicker = ticker
          
          End If
          
          If percentChange > 0 And percentChange > GreatestPcntIncrease Then
            GreatestPcntIncrease = percentChange
            GreatestPcntIncreaseTicker = ticker
            
          End If
          
          If percentChange < 0 And percentChange < GreatestPcntDecrease Then
            GreatestPcntDecrease = percentChange
            GreatestPcntDecreaseTicker = ticker
            
          End If
         
        Next i
        
                   
       'Print Values
       
        ws.Range("P2").Value = GreatestPcntIncreaseTicker
        ws.Range("Q2").Value = Format(GreatestPcntIncrease, "percent")
        
        ws.Range("P3").Value = GreatestPcntDecreaseTicker
        ws.Range("Q3").Value = Format(GreatestPcntDecrease, "percent")
        
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q4").Value = GreatestVol
        
Next ws

End Sub

