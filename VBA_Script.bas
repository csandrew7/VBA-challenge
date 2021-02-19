Attribute VB_Name = "Module1"
Sub StockAnalyse()

    'Loop through all worksheets
    
       For Each ws In Worksheets
       
    'Headers for summary table
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
  
    'Define Variables
            
        Dim Ticker As String
        Dim OP As Double
        Dim CP As Double
        Dim Yearly_Change As Double
        Dim Beginning_OP As Long

            Beginning_OP = 2
        
        Dim Percent_Change As Double
        Dim Total_Volume As Double
            
            Total_Volume = 0
        
        Dim LastRow As Double
            
            'Last_Row Formula
                
                LastRow = Cells(Rows.Count, 1).End(xlUp).Row
                    
        Dim summary_tablerow As Long
            
                summary_tablerow = 2
        
        For i = 2 To LastRow
        
    'Ticker Total Stock Volume
        
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
    'Checking if we are still in the same ticker
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
    'Set Ticker Symbol
            
            Ticker = ws.Cells(i, 1).Value
            
    'Print values to summary table
            
            ws.Range("I" & summary_tablerow).Value = Ticker
            ws.Range("L" & summary_tablerow).Value = Total_Volume
            
    'Reset Total Stock Volume
    
            Total_Volume = 0
                
    'Yearly change calculations
    
        OP = ws.Range("C" & Beginning_OP)
        CP = ws.Range("F" & i)
        Yearly_Change = CP - OP
        
    'Print yearly change value
    
        ws.Range("J" & summary_tablerow).Value = Yearly_Change
        
    'Percent Change Calculations
    
        If OP = 0 Then
            Percent_Change = 0
        Else
            OP = ws.Range("C" & Beginning_OP)
            Percent_Change = Yearly_Change / OP
        End If
    
    'Formatting range to include % and two decimal places
    
        ws.Range("K" & summary_tablerow).NumberFormat = "0.00%"
        ws.Range("K" & summary_tablerow).Value = Percent_Change
        
    'Conditional formatting postive in green, negative in red
    
        If ws.Range("J" & summary_tablerow).Value >= 0 Then
            ws.Range("J" & summary_tablerow).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_tablerow).Interior.ColorIndex = 3
        End If
        
    'Add one to summary tablerow and beginning_OP
    
        summary_tablerow = summary_tablerow + 1
        Beginning_OP = i + 1
            
        End If
            
        Next i
        
   Next ws
   
End Sub
                     

