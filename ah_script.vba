Attribute VB_Name = "Module1"
Sub MYSD()

    ' loop through all

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
        ' last row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row

        Cells(1, 9).Value = "ticker"
        Cells(1, 10).Value = "yearly change"
        Cells(1, 11).Value = "percent change"
        Cells(1, 12).Value = "total stock volume"
        
        Dim openp As Double
        Dim closep As Double
        Dim yearly_change As Double
        Dim ticker_name As String
        Dim percent_change As Double
        Dim volume As Double
        volume = 0
        Dim row As Double
        row = 2
        Dim column As Double
        column = 1
        Dim i As Double
        
        'initial open price
        openp = Cells(2, column + 2).Value
        
        ' Loop through all ticker
        For i = 2 To LastRow
            ' Check info in same ticker, if it is not...
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                ' ticker name
                ticker_name = Cells(i, column).Value
                Cells(row, column + 8).Value = ticker_name
                ' close price
                closep = Cells(i, column + 5).Value
                ' yearly change
                yearly_change = closep - openp
                Cells(row, column + 9).Value = yearly_change
                ' percent change
                percent_change = closep / openp - 1
                Cells(row, column + 10).Value = percent_change
                ' Total Volume
                volume = Cells(i, column + 6).Value + volume
                Cells(row, column + 11).Value = volume
                ' add one to row
                row = row + 1
                ' reset the open price
                openp = Cells(i + 1, column + 2)
                ' reset the Volume
                volume = 0
            'if cells are the same ticker
            Else
                volume = Cells(i, column + 6).Value + volume
            End If
        Next i
        
        ' last row of yearly change
        YCLastRow = WS.Cells(Rows.Count, column + 8).End(xlUp).row
        ' set colors
        For j = 2 To YCLastRow
            If (Cells(j, column + 9).Value >= 0) Then
                Cells(j, column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, column + 9).Value < 0 Then
                Cells(j, column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        
    Next WS
        
End Sub
