Attribute VB_Name = "Module1"
Sub VBA_challenge()

Call stocksummary_table
Call conditional_formatting
Call bonus

End Sub
Sub stocksummary_table()

'Set ws as worksheet.
Dim ws As Worksheet

'Loop through all the sheets.
For Each ws In ThisWorkbook.Worksheets

'Hide the screen-updating.
Application.ScreenUpdating = False

'Set ticker variable to hold value for ticker code.
Dim Ticker As String
        
'Set stock volume variable to hold value for stock volume data and assign initial value.
Dim StockVolume As Double
StockVolume = 0
        
'Set row for summary table and set at line next the header.
Dim SummaryTable_Row As Double
SummaryTable_Row = 2
    
'Set open value variable to hold value for data and assign initial value.
Dim OpenValue As Double
OpenValue = 0
    
'Set close value variable to hold value for data and assign initial value.
Dim CloseValue As Double
CloseValue = 0
    
'Set yearly change variable that calculates the difference between opening and closing stock values.
Dim YearlyChange As Double
    
'Set percent change variable.
Dim PercentChange As Double
       
'Determine the last non-empty record in ticker column.
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
             
'Add data summary headers. Format to bold and autofit cells.
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("I1:L1").Font.Bold = True
ws.Range("I1:L1").Columns.AutoFit
                                                                                              
    'Loop through all ticker records.
    For i = 2 To LastRow
                                                                            
        'Check if the value in ticker cell does not match the next cell record.
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                                        
            'Assign value to ticker variable.
            Ticker = ws.Cells(i, 1)
                                                       
            'Assign value to closing value variable.
            CloseValue = ws.Cells(i, 6)
                                            
            'Add volume data to stock volume variable.
            StockVolume = StockVolume + ws.Cells(i, 7).Value
                                                                 
            'Calculate the difference between closing and open value.
            YearlyChange = CloseValue - OpenValue
                                                     
            'Calculate the percentage change between closing and open value and change number format.
            PercentChange = YearlyChange / OpenValue
            ws.Columns("K").NumberFormat = "0.00%"
                                                                 
            'Print the ticker and stock volume on the summary table.
            ws.Range("I" & SummaryTable_Row).Value = Ticker
            ws.Range("J" & SummaryTable_Row).Value = YearlyChange
            ws.Range("K" & SummaryTable_Row).Value = PercentChange
            ws.Range("L" & SummaryTable_Row).Value = StockVolume
                                            
            'Add one to summary table row.
            SummaryTable_Row = SummaryTable_Row + 1
                                        
            'Reset brand total
            StockVolume = 0
                
                                                                                                                
        'If the value in the ticker cell matches the next cell record.
        Else
                                                
            'Add to the stock volume total.
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
            'Assign value to opening value variable.
            OpenValue = ws.Cells(i, 3)
                                                                                                       
            End If
                                
        End If
    
    Next i
        
Next ws

End Sub

Sub conditional_formatting()

'Set ws as worksheet.
Dim ws As Worksheet

'Loop through all the sheets.
For Each ws In ThisWorkbook.Worksheets

'Hide the screen-updating.
Application.ScreenUpdating = False

'Create format range object
Dim FormatRange As Range

'Set format range cells
Set FormatRange = ws.Range("J:J")

'Determine the last non-empty record in yearly change column.
YC_LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'Loop through all yearly change records.
    For j = 2 To YC_LastRow
    
        'Check if the value in yearly change cell is negative.
        If ws.Cells(j, 10).Value < 0 Then
        
            'Change cells color to red if cell value is less than zero.
            ws.Cells(j, 10).Interior.ColorIndex = 3
        
        'Check if the value in yearly change cell is not negative (zero and positive).
        Else
            'Change cells color to green if cell value is equal or greater than zero.
            ws.Cells(j, 10).Interior.ColorIndex = 4

        End If

    Next j

Next ws

End Sub

Sub bonus()

'Set ws as worksheet.
Dim ws As Worksheet

'Loop through all the sheets.
For Each ws In ThisWorkbook.Worksheets

'Hide the screen-updating.
Application.ScreenUpdating = False

'Set variables to hold value for maximum and minimum percentage increase.
Dim MaxPercentage As Double
Dim MinPercetage As Double

'Set variable to hold value for maximum stock volume.
Dim MaxVolume As Double

'Set variable to hold value for ticker codes.
Dim Max_Ticker As String
Dim Min_Ticker As String
Dim MaxVolume_Ticker As String

'Determine the last non-empty record in the summary table.
Bonus_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Add bonus headers. Format to bold and autofit cells.
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Deccrease"
ws.Range("O4").Value = "Greatest Total"
ws.Range("O:O").Font.Bold = True
ws.Range("O:O").Columns.AutoFit

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("P1:Q1").Font.Bold = True
ws.Range("P1:Q1").Columns.AutoFit

'Determine the maximum percentage increase.
MaxPercentage = WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q2").Value = MaxPercentage
ws.Range("Q2").NumberFormat = "0.00%"

'Determine the minimum percentage increase.
MinPercentage = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q3").Value = MinPercentage
ws.Range("Q3").NumberFormat = "0.00%"

'Determine the maximum stock volume.
MaxVolume = WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("Q4").Value = MaxVolume
ws.Range("Q4").NumberFormat = "#,##0"
ws.Range("Q:Q").Columns.AutoFit

    
    'Loop through all summary table.
    For k = 2 To Bonus_LastRow
        
        'Return the ticker code of stock with maximum % increase.
        If MaxPercentage = ws.Cells(k, 11).Value Then
            Max_Ticker = ws.Cells(k, 9).Value
            ws.Range("P2").Value = Max_Ticker
      
        'Return the ticker code of stock with minimum % increase.
        ElseIf MinPercentage = ws.Cells(k, 11).Value Then
            Min_Ticker = ws.Cells(k, 9).Value
            ws.Range("P3").Value = Min_Ticker

        'Return the ticker code of stock with maximum volume.
        ElseIf MaxVolume = ws.Cells(k, 12).Value Then
            MaxVolume_Ticker = ws.Cells(k, 9).Value
            ws.Range("P4").Value = MaxVolume_Ticker
        
        End If
        
    Next k

Next ws

End Sub



