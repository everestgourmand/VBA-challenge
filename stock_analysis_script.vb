Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim totalVolume As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Long
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim newSheet As Worksheet
    Dim rowNum As Long

    ' Loop through each sheet (representing each quarter)
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then ' Avoid summary sheet if exists
            ws.Activate
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Create new sheet for the output
            Set newSheet = ThisWorkbook.Sheets.Add
            newSheet.Name = ws.Name & "_Analysis"
            rowNum = 2 ' Starting row in new sheet for results
            
            ' Add headers
            newSheet.Cells(1, 1).Value = "Ticker Symbol"
            newSheet.Cells(1, 2).Value = "Total Volume"
            newSheet.Cells(1, 3).Value = "Quarterly Change ($)"
            newSheet.Cells(1, 4).Value = "Percent Change"
            
            ' Loop through the rows of the stock data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value ' Ticker symbol
                openPrice = ws.Cells(i, 2).Value ' Opening price
                closePrice = ws.Cells(i, 5).Value ' Closing price
                totalVolume = ws.Cells(i, 6).Value ' Stock volume

                ' Calculate Quarterly Change and Percent Change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Output the data to the new sheet
                newSheet.Cells(rowNum, 1).Value = ticker
                newSheet.Cells(rowNum, 2).Value = totalVolume
                newSheet.Cells(rowNum, 3).Value = quarterlyChange
                newSheet.Cells(rowNum, 4).Value = percentChange
                
                ' Apply conditional formatting
                If quarterlyChange > 0 Then
                    newSheet.Cells(rowNum, 3).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf quarterlyChange < 0 Then
                    newSheet.Cells(rowNum, 3).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
                
                If percentChange > 0 Then
                    newSheet.Cells(rowNum, 4).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf percentChange < 0 Then
                    newSheet.Cells(rowNum, 4).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If

                ' Track greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                rowNum = rowNum + 1
            Next i
            
            ' Add summary row
            newSheet.Cells(rowNum + 1, 1).Value = "Greatest % Increase"
            newSheet.Cells(rowNum + 1, 2).Value = greatestIncreaseTicker
            newSheet.Cells(rowNum + 1, 3).Value = greatestIncrease
            newSheet.Cells(rowNum + 2, 1).Value = "Greatest % Decrease"
            newSheet.Cells(rowNum + 2, 2).Value = greatestDecreaseTicker
            newSheet.Cells(rowNum + 2, 3).Value = greatestDecrease
            newSheet.Cells(rowNum + 3, 1).Value = "Greatest Total Volume"
            newSheet.Cells(rowNum + 3, 2).Value = greatestVolumeTicker
            newSheet.Cells(rowNum + 3, 3).Value = greatestVolume

        End If
    Next ws
End Sub
