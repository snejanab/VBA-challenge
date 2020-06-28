Attribute VB_Name = "Module1"
Sub ticker_name()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    'Loop through all worksheets in the workbook
    ws.Activate
    
    
    
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set an initial variable for holding the ticker
  Dim ticker_name As String
  'Dim max As Double
  
     
  'Set an initial variables for holding the opening and closing prices
  Dim opening_price As Double
  Dim closing_price As Double
    
  opening_price = Range("C2").Value
   
        
  ' Set an initial variable for holding the total volume per ticker
  Dim Volume_Total As Double
  Volume_Total = 0
  ' Keep track of the location for each ticker in the summary table
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Range("K1").Value = "Ticker"
  Range("L1").Value = "Yearly Change"
  Range("M1").Value = "Percent Change"
  Range("N1").Value = "Volume"
 
  ' Loop through all rows
  For i = 2 To LastRow
    ' Check if we are still within the same ticker, if it is not...
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the Ticker name
      ticker_name = Cells(i, 1).Value
           
      
      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
      ' Print the ticker in the Summary Table
      Range("K" & Summary_Table_Row).Value = ticker_name
                
                         
      closing_price = Cells(i, 6).Value
      yearly_change = closing_price - opening_price
      Range("L" & Summary_Table_Row).Value = yearly_change
        
      If yearly_change >= 0 Then
        
        Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
           
      Else
           
        Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
            
      End If
        
      If opening_price <> 0 Then
        
        percent_change = yearly_change / opening_price
        Range("M" & Summary_Table_Row).Value = FormatPercent(percent_change, 2)
                
      Else
      
        Range("M" & Summary_Table_Row).Value = "N/A"
      
      End If
      
      Range("Q2").Value = "Greatest % Increase"
      Range("Q3").Value = "Greatest % Decrease"
      Range("Q4").Value = "Greatest Volume Increase"
      Range("R1").Value = "Ticker"
      Range("S1").Value = "Value"
        
        
      Range("S2") = FormatPercent(WorksheetFunction.max(Range("M2:M10000")), 2)
      Range("S3") = FormatPercent(WorksheetFunction.Min(Range("M1:M10000")), 2)
      Range("S4") = WorksheetFunction.max(Range("N1:N10000"))
        
      Columns("K:S").AutoFit
        
      opening_price = Cells(i + 1, 3).Value
        
                   
        
      ' Print the Volume_Total to the Summary Table
      Range("N" & Summary_Table_Row).Value = Volume_Total
        
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
          
      ' Reset the Volume Total
      Volume_Total = 0
      
       
            
         
     
    ' If the cell immediately following a row is the same ticker...
    
    Else
      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
                    
                
    End If
    
       
  Next i
   
        
    
    
    For j = 2 To LastRow
    
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Double
    
    max_percent = Range("S2").Value
    min_percent = Range("S3").Value
    max_volume = Range("S4").Value
    
    
    
    
        If Cells(j, 13).Value = max_percent Then
            Range("R2").Value = Cells(j, 11).Value
        End If
        
        If Cells(j, 13).Value = min_percent Then
            Range("R3").Value = Cells(j, 11).Value
        End If
            
        If Cells(j, 14).Value = max_volume Then
            Range("R4").Value = Cells(j, 11).Value
        End If
        
    Next j
    
    Next ws
    
End Sub


