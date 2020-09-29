Sub stockmkt() 

‘Note: I did use only one script for the large stock market data base…as you can see, it worked just fine.

MsgBox ("Iniciando...")

For Each excelsheet
    
    excelsheet.Activate

'Headers
  
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly change"
    Cells(1, 12).Value = "Percent change"
    Cells(1, 13).Value = "Total Stock Volume"
      
'Variables
  
    Dim ticker As String
    Dim oppr, clopr, percent_change As Double
    Dim total_stock As LongLong
      
'Table
  
    Dim row_summary As Integer
    row_summary = 2

'Loop for tickers
  
    TotalRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalCol = Cells(1, Columns.Count).End(xlToLeft).Column
    oppr = 0
    total_stock = 0
 
    For i = 2 To TotalRow
           
'Conditional for oppr
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            oppr = Cells(i, 3).Value
        End If
              
                   
'Getting the closing price:
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            clopr = Cells(i, 6).Value
                                       
            ticker_Name = Cells(i, 1).Value
    
'Ticker in table
        Range("J" & row_summary).Value = ticker_Name
        
            yearly_change = clopr - oppr

'Yearly_Change in table
        Range("K" & row_summary).Value = yearly_change
        
        If yearly_change >= 0 Then
            Range("K" & row_summary).Interior.ColorIndex = 4
        Else
            Range("K" & row_summary).Interior.ColorIndex = 3
            
        End If
                
'Problem with zero values...
    
        If clopr = 0 Then
            percent_change = 0
                        
            Cells(i, 12).Value = 0
            Cells(i, 12).NumberFormat = "#.00%"
        
        ElseIf oppr = 0 Then
            percent_change = 0
                    
        Else
        
            percent_change = ((clopr / oppr) - 1)
        
        End If
             
'Percent change in the table
        Range("L" & row_summary).Value = percent_change
        Range("L" & row_summary).NumberFormat = "#.00%"
            
            total_stock = total_stock + Cells(i, 7).Value

'Total stock in the table
        Range("M" & row_summary).Value = total_stock

'Adding row
            row_summary = row_summary + 1
      
'Reset total stock
            total_stock = 0
            
        Else

        total_stock = total_stock + Cells(i, 7).Value

    End If
          
  Next i
         
    
'Challenges in table
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Range("Q2:Q3").NumberFormat = "#.00%"
                  
'Table increases…I used direct excel functions…

    Dim ticker2 As String
    Dim greatest, minimum, total_volume As Double
    
          
    TotalRow2 = Cells(Rows.Count, 12).End(xlUp).Row
    
    For t = 2 To TotalRow2
                
    If Cells(t, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & TotalRow2)) Then
                Cells(2, 16).Value = Cells(t, 10).Value
                Cells(2, 17).Value = Cells(t, 12).Value
        
        ElseIf Cells(t, 12).Value = Application.WorksheetFunction.min(Range("L2:L" & TotalRow2)) Then
    
                Cells(3, 16).Value = Cells(t, 10).Value
                Cells(3, 17).Value = Cells(t, 12).Value
                
     ElseIf Cells(t, 13).Value = Application.WorksheetFunction.Max(Range("M2:M" & TotalRow2)) Then
    
                Cells(4, 16).Value = Cells(t, 10).Value
                Cells(4, 17).Value = Cells(t, 13).Value
                          
        End If
                  
    Next t
      
Next excelsheet

MsgBox ("Acabamos...")

End Sub

