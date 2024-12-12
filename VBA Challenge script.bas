Attribute VB_Name = "M—dulo 1"
Sub CalculateQuarterlyStockData()
    ' We declare each variable
    Dim ticker As String
    Dim startPrice As Double
    Dim endPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    ' We declared LastRow as LONG because some users recommend this for large datasets
    Dim lastRow As Long
    Dim summaryRow As Long
    ' Variables to save greatest values
    Dim greatestIncrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolume As Double
    Dim greatestVolumeTicker As String
    
    
    ' We loop through each worksheet, we used LastRow formula
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
        ' For adding headers to the summary table on each worksheet
        ' We used "with" as a recommentation online, to avoid repetition
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With
        ' We start writing results from row 2
        summaryRow = 2
        
        ' To loop through each row in the worksheet
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            startPrice = ws.Cells(i, 3).Value
            totalVolume = 0
            
            ' Loop through each row for the CURRENT ticker
            For j = i To lastRow
                If ws.Cells(j, 1).Value <> ticker Then
                    Exit For
                End If
                totalVolume = totalVolume + ws.Cells(j, 7).Value
                endPrice = ws.Cells(j, 6).Value
            Next j
            
            ' Calculate the changes
            quarterlyChange = endPrice - startPrice
            If startPrice <> 0 Then
                percentChange = (quarterlyChange / startPrice)
            Else
                percentChange = 0
            End If
            
            ' Output the results to the summary table on the same worksheet
            With ws
            
                .Cells(summaryRow, 9).Value = ticker
                .Cells(summaryRow, 10).Value = quarterlyChange
                .Cells(summaryRow, 11).Value = percentChange
                .Cells(summaryRow, 12).Value = totalVolume
                
                ' To format the Percent Change as percentage
                .Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' For applying conditional formatting for Quarterly Change
                If quarterlyChange = 0 Then
                    .Cells(summaryRow, 10).Interior.ColorIndex = 0 ' None
                ElseIf quarterlyChange < 0 Then
                    .Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red
                ElseIf quarterlyChange > 0 Then
                    .Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green
                End If
            End With
            
             ' Check for greatest values
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
            
            summaryRow = summaryRow + 1
            ' Adjust the outer loop counter to skip processed rows
            i = j - 1
        Next i
        
         ' Output the greatest values to tall worksheets
        With ws
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = greatestIncreaseTicker
            .Cells(2, 17).Value = greatestIncrease
            .Cells(2, 17).NumberFormat = "0.00%"
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = greatestDecreaseTicker
            .Cells(3, 17).Value = greatestDecrease
            .Cells(3, 17).NumberFormat = "0.00%"
            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = greatestVolumeTicker
            .Cells(4, 17).Value = greatestVolume
        End With
    
    
    Next ws
    
End Sub

