Attribute VB_Name = "Module1"
Sub alphabetical_testing()

'create a script that will loop through all the stocks for one year and output [ticker symbol, yearly change from opening price at beginning of given year to closing price at end of year, percent change from opening price, total stock volume of stock]
Dim ticker As String
Dim day As Date
Dim OpenValue As Double
Dim CloseValue As Double
Dim PerChange As Integer
Dim Volume As LongLong
Dim OutputRow As Integer
Dim YearChange As Double


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Value"


'create variable to hold file name, last row, last column, and year
Dim WorksheetName As String


For Each ws In Worksheets
    
'determine last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

OpenValue = Cells(2, 3).Value
OutputRow = 2

    'loop through all stocks
    For i = 2 To LastRow
    
              
        'define the step to switch from ticker symbols
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'define close value
                CloseValue = Cells(i, 6).Value
            
                    'Subtract open from close
                    YearChange = (CloseValue - OpenValue)
                    
                    'Calculate PercentChange
                    PerChange = ((YearChange) / OpenValue)
                    
                    'place values in output row cells
                    Cells(OutputRow, 10).Value = YearChange
                        If Cells(OutputRow, 10).Value > 0 Then
                            Cells(OutputRow, 10).Interior.ColorIndex = 4
                        ElseIf Cells(OutputRow, 10).Value < 0 Then
                            Cells(OutputRow, 10).Interior.ColorIndex = 3
                        ElseIf Cells(OutputRow, 10).Value = 0 Then
                            Cells(OutputRow, 10).Interior.ColorIndex = 2
                        End If
                        
                    Cells(OutputRow, 11).Value = PerChange
                    
                    Cells(OutputRow, 11).NumberFormat = "0.00%"
                    
                    Cells(OutputRow, 12).Value = Volume
                    
                'place ticker in cell value
                Cells(OutputRow, 9).Value = Cells(i, 1).Value
            
                'define output row's increase
                OutputRow = OutputRow + 1
                
                'define cell values when switchover occurs
                OpenValue = Cells(i + 1, 3).Value
                
                'define new volume var for next range
                Volume = 0
                
                
            Else: Cells(i + 1, 1).Value = Cells(i, 1).Value
            
                'Sum Volume
                Volume = Volume + (Cells(i, 7))
                    
                
                
                'add yearly change from OpenValue to CloseValue in year
                
            
        
        End If
        
    Next i
    
Next ws


End Sub

