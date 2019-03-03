Sub calculateTotalVolume()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        Dim calculateResult As Integer
        calculateResult = calculate()
    Next
End Sub
Function calculate()
    Dim lastRow_ As Long
    lastRow_ = lastRow()
    
    Dim resultsIndex As Integer
    resultsIndex = 1
    Dim currentTicker As String
    Dim currentTotalVolume As Double
    Dim rangeLimits() As String
    Dim firstOpen As Double
    Dim lastClose As Double
    
    Dim greatePercentDecrease As Double
    Dim greatePercentIncrease As Double
    Dim greateTotalVolume As Double
    Dim greatePercentDecreaseTicker As String
    Dim greatePercentIncreaseTicker As String
    Dim greateTotalVolumeTicker As String


    For i = 2 To lastRow_ + 1
        If Cells(i, 1).Value = currentTicker Then
            currentTotalVolume = currentTotalVolume + Cells(i, 7)
            lastClose = Cells(i, 6).Value
        Else
        
            Dim yearlyChange As Double
            Dim percentChange As Double
            
            yearlyChange = (firstOpen - lastClose) * -1
            If firstOpen > 0 Then
                percentChange = ((lastClose * 100 / firstOpen) / 100) - 1
            End If
            
            If percentChange > greatePercentIncrease Then
               greatePercentIncrease = percentChange
               greatePercentIncreaseTicker = currentTicker
            ElseIf percentChange < greatePercentDecrease Then
                greatePercentDecrease = percentChange
                greatePercentDecreaseTicker = currentTicker
            End If
            
            If currentTotalVolume > greateTotalVolume Then
                greateTotalVolume = currentTotalVolume
                greateTotalVolumeTicker = currentTicker
            End If
            
        
            Cells(resultsIndex, 10).Value = currentTicker
            Cells(resultsIndex, 11).Value = currentTotalVolume
            Cells(resultsIndex, 12).Value = yearlyChange
            Cells(resultsIndex, 12).NumberFormat = "0.00"
            Cells(resultsIndex, 13).Value = percentChange
            Cells(resultsIndex, 13).NumberFormat = "0.00%"
            
            If (yearlyChange > 0) Then
                Cells(resultsIndex, 12).Interior.ColorIndex = 4
            Else
                Cells(resultsIndex, 12).Interior.ColorIndex = 3
            End If


            
            resultsIndex = resultsIndex + 1

            firstOpen = Cells(i, 3).Value
            currentTicker = Cells(i, 1).Value
            currentTotalVolume = Cells(i, 7).Value
        End If

    Next i
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = greatePercentIncreaseTicker
    Cells(2, 17).Value = greatePercentIncrease
    Cells(2, 17).NumberFormat = "0.00%"
    'Cells(2, 17).Style = "Percent"
    
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = greatePercentDecreaseTicker
    Cells(3, 17).Value = greatePercentDecrease
    Cells(3, 17).NumberFormat = "0.00%"
    'Cells(3, 17).Style = "Percent"
    
    Cells(4, 15).Value = "Greatest total volume"
    Cells(4, 16).Value = greateTotalVolumeTicker
    Cells(4, 17).Value = greateTotalVolume
    'Cells(4, 17).Style = "text"
    
    
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Total Stock Volume"
    Range("L1").Value = "Yearly Change"
    Range("L1").Interior.ColorIndex = 2
    Range("M1").Value = "Percent Change"
    calculate = 0
End Function
Function lastRow()
    
    Dim sht As Worksheet
    Set sht = ActiveSheet
    lastRow = sht.Range("A1").CurrentRegion.Rows.Count
    
End Function
