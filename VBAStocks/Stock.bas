Attribute VB_Name = "Module1"
Sub Stock()

Dim openstock As Double             ' Dim variables openstock at the beginning of the year, closestock at the end of the year
Dim closestock As Double
Dim difference As Double             ' Dim variables the difference between open and close stock, and total stock volume as the sum of all of the stock count
Dim total_stock_volume As LongLong
Dim lastrow As Long                 ' Dim lastrow as the total row count of the data
Dim newrow As Integer                 ' Dim newrow as the counter for the new summary table
Dim wsname As String

For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    newrow = 2
    total_stock_volume = 0
' --------------------------------------------
' Generate new table headers

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Changed"
    ws.Cells(1, 11).Value = "Percentage Changed"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
'-----------------------------------------------
    openstock = ws.Cells(2, 3).Value            ' Take the initial openstock value of the first year in a sheet

' Looping through the given data
    For i = 2 To lastrow - 1
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(newrow, 9) = ws.Cells(i, 1)                        ' Ticker column
            closestock = ws.Cells(i, 6).Value                             '  Close Stock column
            difference = closestock - openstock
            ws.Cells(newrow, 10) = difference                          ' Yearly change column
            '-----------------------------------------------
            ' Formatting the yearly change, green for positive and red for negative
            If difference > 0 Then
                ws.Cells(newrow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(newrow, 10).Interior.ColorIndex = 3
            End If
            '-----------------------------------------------
            ' Check if the open stock is zero
            If openstock = 0 Then
                ws.Cells(newrow, 11).Value = 0
            Else
            ws.Cells(newrow, 11) = Round((difference / openstock) * 100, 2)
            End If
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            ws.Cells(newrow, 12) = total_stock_volume
            newrow = newrow + 1                          ' Count to next row in the new table
            openstock = ws.Cells(i + 1, 3).Value     ' Take the open stock for the following year
            total_stock_volume = 0                         ' Reset the total stock volume for next year
        Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7)
        End If
    Next i

' -----------------------------------------------
' Bouns part
' Reset the values before starting a new sheet
' -----------------------------------------------
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim max_volume As LongLong

    greatest_increase = 0
    greatest_decrease = 0
    max_volume = 0

    For j = 2 To newrow                 ' newrow is the last row for the new table
        If ws.Cells(j, 11) > 0 And ws.Cells(j, 11) >= greatest_increase Then
            greatest_increase = ws.Cells(j, 11)
            greatest_increase_ticker = ws.Cells(j, 9)
        ElseIf ws.Cells(j, 11) <= 0 And Abs(ws.Cells(j, 11)) >= Abs(greatest_decrease) Then
            greatest_decrease = ws.Cells(j, 11)
            greatest_decrease_ticker = ws.Cells(j, 9)
        End If
        
        If ws.Cells(j, 12) > max_volume Then
            max_volume = ws.Cells(j, 12)
            max_volume_ticker = ws.Cells(j, 9)
        End If
    Next j

    '-----------------------------------------------
    ' generate the table with all of the values
    ws.Cells(2, 16) = greatest_increase_ticker
    ws.Cells(2, 17) = greatest_increase
    ws.Cells(3, 16) = greatest_decrease_ticker
    ws.Cells(3, 17) = greatest_decrease
    ws.Cells(4, 16) = max_volume_ticker
    ws.Cells(4, 17) = max_volume
Next
 
End Sub

