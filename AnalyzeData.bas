Attribute VB_Name = "AnalyzeData"
Sub AnalazyData():
    
    'Current worksheet ws as a worksheet
    Dim ws As Worksheet
    
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
            If OpePrice <> 0 Then
                PercentChange = (PercentChange / OpenPrice) * 100
            Else
                MsgBox ("Ope price is 0 and cannot be devided to get Percent Change")
            End If
            
            'Add total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'Print values and ticker name into Summary Table
            ws.Range("I").Value = ticket
            ws.Range("J").Value = YearlyChange
            'Set color for YearlyChange where green is positive and red for negative values
            If (YearlyChange > 0) Then
                ws.Range("J").Interior.ColorIndex = 4
            ElseIf (YearlyChange <= 0) Then
                ws.Range("J").Interior.ColorIndex = 3
            End If
            
            ws.Range("K").Value = (CStr(PercentChange) & "%")
            ws.Range("L").Value = totalVolume
            ' If the cell ticker is same, add volume.
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
                
    
    
End Sub
