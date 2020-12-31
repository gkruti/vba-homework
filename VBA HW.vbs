Attribute VB_Name = "Module2"

Sub UpdateAll()

Attribute VB_Name = "Module1"

Dim ws As Worksheet
Dim Ticker As String
Dim OpenStock As Double
Dim CloseStock As Double
Dim TotalVolume As Double
Dim TotalStock As Double
Dim YearlyChange As Double
Dim OutputRow As Integer
Dim LastRow As Double
Dim PercentChange As Double




For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume”"


    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Ticker = ws.Cells(2, 1).Value
    OpenStock = ws.Cells(2, 3).Value
    OutputRow = 2
    TotalStock = 0
    YearlyChange = ws.Cells(2, 10).Value
    YearlyChange = 0
    

    For i = 2 To LastRow
        
        Ticker_total = 0
        
        If (Ticker <> ws.Cells(i, 1).Value) <> ws.Cells(i, 1).Value Then
        
        ws.Cells(OutputRow, 9).Value = Ticker
    
        CloseStock = ws.Cells(LastRow, 6).Value
        
        YearlyChange = CloseStock - OpenStock

        ws.Cells(OutputRow, 10).Value = YearlyChange
        
        If YearlyChange > 0 Then
            ws.Cells(OutputRow, 10).Interior.ColorIndex = 4
        ElseIf YearlyChange < 0 Then
        ws.Cells(OutputRow, 10).Interior.ColorIndex = 3

        End If

        If OpenStock = 0 Then
            ws.Cells(OutputRow, 11).Value = 0
            ws.Cells(OutputRow, 11).Value = YearlyChange / OpenStock
    
            PercentChange = 0
      
        End If
        

        ws.Cells(OutputRow, 11).Value = PercentChange
        ws.Cells(OutputRow, 11).NumberFormat = "0.00%"

        ws.Cells(OutputRow, 12).Value = TotalStock

        Ticker = ws.Cells(i, 1).Value
        OpenStock = ws.Cells(i, 3).Value
        OutputRow = OutputRow + 1
        
        TotalStock = ws.Cells(i, 7).Value
    
        Else
            TotalStock = TotalStock + ws.Cells(i, 7).Value
    End If
    
Next i
    
        Next ws

End Sub


End Sub


