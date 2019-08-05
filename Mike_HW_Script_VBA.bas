Sub StockData()


Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

        LRow = WS.Cells(Rows.Count, 1).End(xlUp).Row


        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Vol"

        Dim StartP As Double
        Dim EndP As Double
        Dim ChangeP As Double
        Dim ticker1 As String
        Dim percentC As Double
        Dim Vol As Double
        Vol = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        StartP = Cells(2, Column + 2).Value

        
        For i = 2 To LRow

            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ticker1 = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = ticker1
                EndP = Cells(i, Column + 5).Value
                ChangeP = EndP - StartP
                Cells(Row, Column + 9).Value = ChangeP
                If (StartP = 0 And EndP = 0) Then
                    percentC = 0
                ElseIf (StartP = 0 And EndP <> 0) Then
                    percentC = 1
                Else
                    percentC = ChangeP / StartP
                    Cells(Row, Column + 10).Value = percentC
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If


                Vol = Vol + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Vol
                Row = Row + 1
                StartP = Cells(i + 1, Column + 2)
                Vol = 0
            Else
                Vol = Vol + Cells(i, Column + 6).Value
            End If
        Next i
        
        YCLRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        For j = 2 To YCLRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Vol"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"

        For Z = 2 To YCLRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub
