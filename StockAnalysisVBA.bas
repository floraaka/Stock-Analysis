Sub ProcessStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    ' Looping through each worksheet (quarter)
    For Each ws In ThisWorkbook.Worksheets
        ' Finding the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initializing the variables responsible for tracking greatest increase, decrease, and volume
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        
        ' Adding headers for new created columns
        ws.Cells(1, 7).Value = "Quarterly Change ($)"
        ws.Cells(1, 8).Value = "Percent Change (%)"
        ws.Cells(1, 9).Value = "Total Stock Volume"
        
        ' Looping through rows to calculate required metrics
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value
            
            ' Calculating quarterly change
            quarterlyChange = closePrice - openPrice
            
            ' Calculating percent change
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Outputing calculated values to respective columns
            ws.Cells(i, 7).Value = quarterlyChange
            ws.Cells(i, 8).Value = percentChange
            ws.Cells(i, 9).Value = volume
            
            ' Updating greatest increase, decrease, and volume if applicable
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If volume > greatestVolume Then
                greatestVolume = volume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Highlighting cells for greatest increase, decrease, and volume
        HighlightGreatest ws, greatestIncreaseTicker, 8 ' Column 8 is percent change
        HighlightGreatest ws, greatestDecreaseTicker, 8 ' Column 8 is percent change
        HighlightGreatest ws, greatestVolumeTicker, 9 ' Column 9 is volume
    Next ws
End Sub

Sub HighlightGreatest(ws As Worksheet, ticker As String, col As Long)
    Dim lastRow As Long
    Dim cell As Range
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Looping through each cell in the specified column range
    For Each cell In ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
        ' Checking if the ticker symbol matches the one with greatest increase or decrease
        If ws.Cells(cell.Row, 1).Value = ticker Then
            ' Applying conditional formatting based on column and condition
            If col = 8 Then ' Column 8 is percent change
                If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0) ' Red for negative change
                Else
                    ' Optionally handle cases where percent change is zero
                    cell.Interior.Color = RGB(255, 255, 255) ' White color
                End If
            ElseIf col = 9 Then ' Column 9 is volume
                ' Apply different formatting for volume
                 cell.Interior.Color = RGB(0, 0, 255) ' Blue color
            End If
        End If
    Next cell
End Sub

