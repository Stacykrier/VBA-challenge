Attribute VB_Name = "Module1"
'Title
Sub StockanalysisTest()

'Declare worksheets in workbook
Dim ws As Worksheet
Dim wb As Workbook
For Each ws In Worksheets

'Set up Headers for new data columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

        
'Set up Ticker
Dim Ticker As String
TickerRow = 1

'Set up Variables
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim BeginYear As Double
Dim CloseYear As Double
Dim StockVolume As Double
Dim PercentChange As Double
Dim YearChange As Double
Dim TickerCount As Long

'Set up Summary Table
Dim SummaryTable As Double
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest total volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
SummaryTable = 2

'Initialize variables
StockVolume = 0
OpenPrice = ws.Cells(2, 3).Value
PercentValue = 0

'Finding the last row to search
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Creating Loop
    For i = 2 To LastRow
    
'Looping through the tickers using next cell and grabbing StockVolume if condition is met
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            StockVolume = StockVolume + ws.Cells(i, 7).Value
             
'Setting ClosePrice and doing the math calculation of the OpenPrice
            ClosePrice = ws.Cells(i, 6).Value
            YearChange = ClosePrice - OpenPrice
            PercentChange = (YearChange / OpenPrice)
            ws.Range("I" & SummaryTable).Value = ws.Cells(i, 1).Value
            ws.Range("J" & SummaryTable).Value = YearChange
            ws.Range("K" & SummaryTable).Value = PercentChange
            ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
            ws.Range("L" & SummaryTable).Value = StockVolume
 
            
'Reset variables
            
            StockVolume = 0
            SummaryTable = SummaryTable + 1
            OpenPrice = ws.Cells(i + 1, 3).Value

'Grabbing StockVolume when condition is not met
        Else
                
           StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        End If
        Next i

'Setting up conditional formatting
    
    SummaryTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To SummaryTable
        
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
        Next i
       
'Setting up Greatest % Increase/Decrease and Volume by Ticker

        For i = 2 To SummaryTable
            
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & SummaryTable)) Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 16).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & SummaryTable)) Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 16).NumberFormat = "0.00%"
                
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & SummaryTable)) Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
                
            End If
            
        Next i
        Next ws
    
End Sub
