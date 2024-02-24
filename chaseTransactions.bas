Attribute VB_Name = "chaseTransactions"
Public Sub FormatChaseTransactions()
    If ActiveSheet.Range("A1").Value = "Transaction Date" Then
        Dim totalRows As Long
        With ActiveSheet.Range("A1").CurrentRegion
            totalRows = .Rows.Count
            .Columns.AutoFit
            .EntireColumn.HorizontalAlignment = xlHAlignLeft
        End With
        ActiveSheet.Range("A1").EntireRow.Font.Bold = True
        ActiveSheet.Range("H1").Value = "Fee"
        Dim feeRange As Range
        Set feeRange = ActiveSheet.Range("H2", Cells(totalRows, 8))
        feeRange.FormulaR1C1 = "=if(LOWER(RC[-3]=""sale""),ABS(FIXED(RC[-2]*0.03,2)),"""")"
        feeRange.EntireColumn.AutoFit
        Dim transactionRange As Range
        Set transactionRange = ActiveSheet.Range("A2", _
            Cells(Range("A2").CurrentRegion.Rows.Count, Range("A2").CurrentRegion.Columns.Count))
        With transactionRange.FormatConditions
            .Delete
            .Add Type:=xlExpression, Formula1:="=IF($G2<>"""",TRUE,FALSE)"
        End With
        With transactionRange.FormatConditions(1)
            .Interior.Color = RGB(245, 245, 245)
            .Font.Color = RGB(192, 192, 192)
            .StopIfTrue = False
        End With
    End If
End Sub

