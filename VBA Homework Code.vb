'* Create a script that will loop through all the stocks for one year and output the ~
' following information.

  ' The ticker symbol.

  ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  ' The total stock volume of the stock.

'You should also have conditional formatting that will highlight positive change in green and negative change in red.



Sub stock_data()

'Create variables to hold Ticker Symbol, Yearly Change, P_Change
    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim P_Change As Double
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    
'Define Open and Close Price
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    
    Dim Close_Price As Double
    Close_Price = Cells(2, 6).Value
    

'create a variable to define last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Create Summary Table Headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"


'Keep track of the location of each brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Loop through ticker symbol
    For i = 2 To lastrow
    
 ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set ticker symbol
     Ticker_Symbol = Cells(i, 1).Value
     
     'set total stock volume
     Total_Stock_Volume = Cells(i, 7).Value
     
     ' Print the Ticker Symbol in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Symbol
      
      ' Print the Total Stock Volume in the Summary Table
      Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
      
      'Add close price
      Close_Price = Cells(i, 6).Value
      
     'define yearly change
     Yearly_Change = (Close_Price - Open_Price)
     
      ' Print Yearly Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = Yearly_Change
    
     
     'set P_Change
     'take the yearly change / open price
     'error message add non-divisiability condition
     
     If (Open_Price = 0) Then
     
        P_Change = 0
        
    Else
    
        P_Change = (Yearly_Change / Open_Price)
        
    End If
    
    'Print p change in summary table formatt to %
    Range("L" & Summary_Table_Row).Value = P_Change
    Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
    


      ' Add one to the summary table row
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Holding more summary tables in case I need them
      'Summary_Table_Row = Summary_Table_Row + 1
      'Summary_Table_Row = Summary_Table_Row + 1
      
      
      ' Reset Yearly Change
      'Thinking I may need to change this variable to Total Stock Volume? Changed.
      Total_Stock_Volume = 0
      
      'Reset opening price
      Open_Price = Cells(i + 1, 3)
      
      
       ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i
  
 
 lastrow_summary_table_row = Cells(Rows.Count, 10).End(xlUp).Row
  
  'Color code for yearly change
  
   For i = 2 To lastrow_summary_table_row
            If Cells(i, 11).Value > 0 Then
            
            Cells(i, 11).Interior.ColorIndex = 10
            
            Else
            
            Cells(i, 11).Interior.ColorIndex = 3
                
            End If
            
    Next i
    

End Sub
