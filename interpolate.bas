Attribute VB_Name = "interpolate"
Public Function Interpolate_Multilinear(x As Range, x_range As Range, y_range As Range) As Double

    Dim xs() As Variant, ys() As Variant, x_value As Double
    Dim i As Long, j As Long, k As Long

    x_value = x.Value
    
    If x_range.Columns.Count > x_range.Rows.Count Then
        xs = WorksheetFunction.Transpose(x_range)
    Else
        xs = x_range
    End If
    
    If y_range.Columns.Count > y_range.Rows.Count Then
        ys = WorksheetFunction.Transpose(y_range)
    Else
        ys = y_range
    End If
    
    Select Case x_value
        Case Is <= xs(LBound(xs, 1), 1)
            Interpolate_Multilinear = ys(LBound(ys, 1), 1)
            Exit Function
        Case Is >= xs(UBound(xs, 1), 1)
            Interpolate_Multilinear = ys(UBound(ys, 1), 1)
            Exit Function
        Case Else
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            For i = LBound(xs, 1) To UBound(xs, 1)
                dict.Add xs(i, 1), ys(i, 1)
            Next
            If dict.Exists(x_value) Then
                Interpolate_Multilinear = dict(x_value)
                Set dict = Nothing
            Else
                For k = LBound(xs, 1) To UBound(xs, 1) - 1
                    If x_value > xs(k, 1) And x_value < xs(k + 1, 1) Then
                        i = k
                        j = k + 1
                        Exit For
                    End If
                Next
                Interpolate_Multilinear = (x_value - xs(i, 1)) * (ys(j, 1) - ys(i, 1)) / (xs(j, 1) - xs(i, 1)) + ys(i, 1)
            End If
    End Select
End Function

Public Function Interpolate_BilinearGrid(x As Range, y As Range, x_range As Range, y_range As Range, z_range As Range, Optional z_x_as_rows As Boolean = True) As Double

    Dim xs() As Variant, ys() As Variant, zs() As Variant
    Dim x_value As Double, y_value As Double
    Dim i As Long
    Dim index_xi As Long, index_xj As Long, index_yi As Long, index_yj As Long
    Dim xi As Double, xj As Double, yi As Double, yj As Double
    Dim zii As Double, zij As Double, zji As Double, zjj As Double
    Dim x_dict As Object, y_dict As Object
    
    x_value = x.Value
    y_value = y.Value
    
    If x_range.Columns.Count > x_range.Rows.Count Then
        xs = WorksheetFunction.Transpose(x_range)
    Else
        xs = x_range
    End If
    
    If y_range.Columns.Count > y_range.Rows.Count Then
        ys = WorksheetFunction.Transpose(y_range)
    Else
        ys = y_range
    End If
    
    If z_x_as_rows Then
        zs = z_range
    Else
        zs = WorksheetFunction.Transpose(z_range)
    End If
    
    If x_value <= xs(LBound(xs, 1), 1) Then
        index_xi = 1
        index_xj = 1
    ElseIf x_value >= xs(UBound(xs, 1), 1) Then
        index_xi = UBound(xs, 1)
        index_xj = UBound(xs, 1)
    Else
        Set x_dict = CreateObject("Scripting.Dictionary")
        For i = LBound(xs, 1) To UBound(xs, 1)
            x_dict.Add xs(i, 1), i
        Next
        If x_dict.Exists(x_value) Then
            index_xi = x_dict(x_value)
            index_xj = x_dict(x_value) + 1
        Else
            For i = LBound(xs, 1) To UBound(xs, 1) - 1
                If x_value > xs(i, 1) And x_value < xs(i + 1, 1) Then
                    index_xi = i
                    index_xj = i + 1
                    Exit For
                End If
            Next
        End If
        Set x_dict = Nothing
    End If

    If y_value <= ys(LBound(ys, 1), 1) Then
        index_yi = 1
        index_yj = 1
    ElseIf y_value >= ys(UBound(ys, 1), 1) Then
        index_yi = UBound(ys, 1)
        index_yj = UBound(ys, 1)
    Else
        Set y_dict = CreateObject("Scripting.Dictionary")
        For i = LBound(ys, 1) To UBound(ys, 1)
            y_dict.Add ys(i, 1), i
        Next
        If y_dict.Exists(y_value) Then
            index_yi = y_dict(y_value)
            index_yj = y_dict(y_value) + 1
        Else
            For i = LBound(ys, 1) To UBound(ys, 1) - 1
                If y_value > ys(i, 1) And y_value < ys(i + 1, 1) Then
                    index_yi = i
                    index_yj = i + 1
                    Exit For
                End If
            Next
        End If
        Set y_dict = Nothing
    End If
    
    xi = xs(index_xi, 1)
    xj = xs(index_xj, 1)
    
    yi = ys(index_yi, 1)
    yj = ys(index_yj, 1)
    
    zii = zs(index_xi, index_yi)
    zji = zs(index_xj, index_yi)
    zij = zs(index_xi, index_yj)
    zjj = zs(index_xj, index_yj)
    
    If xi <> xj And yi <> yj Then
        Interpolate_BilinearGrid = (yj - y_value) / (yj - yi) * ((xj - x_value) / (xj - xi) * zii + (x_value - xi) / (xj - xi) * zji) + _
                              (y_value - yi) / (yj - yi) * ((xj - x_value) / (xj - xi) * zij + (x_value - xi) / (xj - xi) * zjj)
    ElseIf xi = xj And yi <> yj Then
        Interpolate_BilinearGrid = (y_value - yi) * (zij - zii) / (yj - yi) + zii
    ElseIf yi = yj And xi <> xj Then
        Interpolate_BilinearGrid = (xj - x_value) / (xj - xi) * zii + (x_value - xi) / (xj - xi) * zji
    Else
        Interpolate_BilinearGrid = zs(index_xi, index_yi)
    End If
End Function









