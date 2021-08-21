Attribute VB_Name = "AnalyzeData"
Sub AnalazyData():
    
    'Current worksheet ws as a worksheet
    Dim ws As Worksheet
    'Loop through all the worksheets
    For Each ws In Worksheets
        'Set initial variables for holding the ticket name
        Dim ticket As String
        ticket = " "
        'Set initial variable for holding total volume
        Dim totalVolume As Double
        totalVolume = 0
        'Set variables for OpenPrice, ClosePrice
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        OpenPrice = 0
        ClosePrice = 0
        'Set variable for Yearly Change
        Dim YearlyChange As Double
        YearlyChange = 0
        'Set variable for PercentChange
        Dim PercentChange As Double
        PercentChange = 0
        
        '-----------------------------------------------------
        'Keep track of ticker in summary table
        Dim summaryTable As Long
        summaryTable = 2
        'Set initial row for current ws
        Dim lastRow As Long
        Dim i As Long
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Fill out headers for Summary Table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set initial OpenPrice and for loop to get the rest open prices
        OpenPrice = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            'Check if same ticket, if not, then add open price and start next ticket
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Set ticker name
                ticket = ws.Cells(i, 1).Value
                'Calculate Yearly Change and Percent Change
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                'Check for division by 0
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                'Add total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'Print values and ticker name into Summary Table
                ws.Range("I" & summaryTable).Value = ticket
                ws.Range("J" & summaryTable).Value = YearlyChange
                'Set color for YearlyChange where green is positive and red for negative values
                If (YearlyChange > 0) Then
                    ws.Range("J" & summaryTable).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                    ws.Range("J" & summaryTable).Interior.ColorIndex = 3
                End If
                
                ws.Range("K" & summaryTable).Value = (CStr(PercentChange) & "%")
                ws.Range("L" & summaryTable).Value = totalVolume
                summaryTable = summaryTable + 1
                'Reset values for new Ticket
                YearlyChange = 0
                ClosePrice = 0
                OpenPrice = ws.Cells(i + 1, 3)
                PercentChange = 0
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
                
    Next ws
    
End Sub
