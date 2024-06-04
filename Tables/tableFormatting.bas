Attribute VB_Name = "tableFormatting"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "tableFormatting"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "2.1.2"
Public Const module_date As Date = #6/4/2024#
Public Const module_notes As String = _
    "This module is necessary for setting table related styles that " & _
    "cannot be set in the Workbook TableStyle element. This means " & _
    "the formatting must be done after table creating and INITial " & _
    "styling, and is set on the Worksheet level, not at the Workbook."
Public Const module_license As String = "GNU General Public License, v3.0"

'***********************************************************************
'                          Referenced Libraries
'***********************************************************************
' NONE


'***********************************************************************
'                        Module Level Preferences
'***********************************************************************
Public Const TYPICAL_FONT_NAME As String = "Arial"
Public Const TYPICAL_FONT_SIZE As Long = 10

Public Const ROW_PADDING As Long = 6

'***********************************************************************
'                  Utility Functions & Class Factories
'***********************************************************************
Function TableHasData(aTable As ListObject) As Boolean
    If Not aTable.DataBodyRange Is Nothing Then TableHasData = True
End Function

Public Function CreateTempRow(ByRef aTable As ListObject) As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        'The following resets the formats that are copied down from the
        ' header by default when a row is inserted.
        With aTable.ListRows(1).Range
            .Interior.Color = xlNone
            .Font.Bold = False
            .Font.Color = vbBlack
            .Font.Size = TYPICAL_FONT_SIZE
            .Font.Name = TYPICAL_FONT_NAME
        End With
        CreateTempRow = True
    End If
End Function

Public Function CreateFormatConfig(InteriorColor As Long, _
                                   Optional fontName As String = TYPICAL_FONT_NAME, _
                                   Optional fontSize As Double = TYPICAL_FONT_SIZE, _
                                   Optional isBold As Boolean = False) As FormatConfig
    ' This is a class factory to speed up instantiation.
    Dim style As FormatConfig
    Set style = New FormatConfig
    style.INIT InteriorColor, fontName, fontSize, isBold
    Set CreateFormatConfig = style
End Function


'***********************************************************************
'                         Fast Format Utilities
'***********************************************************************
Private Sub ResetFonts()
    With ActiveSheet.Cells
        .Font.Name = TYPICAL_FONT_NAME
        .Font.Size = TYPICAL_FONT_SIZE
    End With
End Sub

Public Sub ResetAllColumnWidths()
    ActiveSheet.Cells.ColumnWidth = 8.38
End Sub

Public Sub SetColumnWidths(aTable As ListObject, _
                           columnRef As String, _
                           columnWidths As String)
    ' Assumes columnRef and columnWidths are parsable strings
    ' ColumnRef can be a string of column numbers or the list
    ' column names.
    Dim cols As Variant, widths As Variant
    cols = ParseToArray(columnRef)
    widths = RecastArray(ParseToArray(columnWidths), vbDouble)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If IsNumeric(cols(i)) Then
            aTable.DataBodyRange.Columns(CLng(cols(i))).ColumnWidth = widths(i)
        Else
            aTable.ListColumns(cols(i)).DataBodyRange.ColumnWidth = widths(i)
        End If
    Next
    
End Sub

Public Sub AutoFitColumnWidths(aTable As ListObject)
    aTable.DataBodyRange.Columns.EntireColumn.AutoFit
End Sub

Public Sub WrapTextInColumns(aTable As ListObject, _
                             Optional wrapColumns As String = "", _
                             Optional isWrapped As Boolean = True)
    ' Assumes wrapColumns is a parsable string of either column names
    ' or indicies of columns. If nothing provided, all columns are wrapped.
    ' This sub also UNWRAPS columns by setting the isWrapped to False
    
    If wrapColumns <> "" Then
        Dim cols As Variant
        cols = ParseToArray(wrapColumns)
        Dim i As Long
        For i = LBound(cols) To UBound(cols)
            If IsNumeric(cols(i)) Then
                aTable.DataBodyRange.Columns(CLng(cols(i))).WrapText = isWrapped
            Else
                aTable.ListColumns(cols(i)).DataBodyRange.WrapText = isWrapped
            End If
        Next
    Else
        aTable.DataBodyRange.WrapText = isWrapped
    End If
End Sub

Public Sub HorizontalAlignColumns(aTable As ListObject, _
                             Optional targetColumns As String = "", _
                             Optional hAlign As XlHAlign = xlHAlignLeft)
    ' Align columns by name or index horizontally. If no columns specified,
    ' will align the whole table to the provided value.
    
    If targetColumns <> "" Then
        Dim cols As Variant
        cols = ParseToArray(targetColumns)
        Dim i As Long
        For i = LBound(cols) To UBound(cols)
            If IsNumeric(cols(i)) Then
                aTable.Columns(CLng(cols(i))).DataBodyRange.HorizontalAlignment = hAlign
            Else
                aTable.ListColumns(cols(i)).DataBodyRange.HorizontalAlignment = hAlign
            End If
        Next
    Else
        aTable.DataBodyRange.HorizontalAlignment = hAlign
    End If
End Sub

Public Sub VerticalAlignColumns(aTable As ListObject, _
                             Optional targetColumns As String = "", _
                             Optional vAlign As XlVAlign = xlVAlignTop)
    ' Align columns by name or index horizontally. If no columns specified,
    ' will align the whole table to the provided value.
    
    If targetColumns <> "" Then
        Dim cols As Variant
        cols = ParseToArray(targetColumns)
        Dim i As Long
        For i = LBound(cols) To UBound(cols)
            If IsNumeric(cols(i)) Then
                aTable.Columns(CLng(cols(i))).DataBodyRange.VerticalAlignment = vAlign
            Else
                aTable.ListColumns(cols(i)).DataBodyRange.VerticalAlignment = vAlign
            End If
        Next
    Else
        aTable.DataBodyRange.VerticalAlignment = vAlign
    End If
End Sub

Public Sub ResetTableRowHeights(aTable As ListObject)
    Call WrapTextInColumns(aTable, , False) 'Flatten rows
    aTable.DataBodyRange.Rows.AutoFit
End Sub

Public Sub ApplyComfyRowsToTable(aTable As ListObject, _
                                 Optional rowPadding As Double = ROW_PADDING, _
                                 Optional wrapColumns As String = "", _
                                 Optional maxHeight As Double = 0)
    Call utilities.MemorySaver
    
    Call ResetTableRowHeights(aTable)
    Call WrapTextInColumns(aTable, wrapColumns)
    aTable.DataBodyRange.Rows.AutoFit
    
    Dim aRow As Range, newHeight As Double
    For Each aRow In aTable.DataBodyRange.Rows
        aRow.RowHeight = aRow.RowHeight + rowPadding
    Next

    Call utilities.MemoryRestore
End Sub


'***********************************************************************
'              Data Validation and Conditional Formatting
'***********************************************************************
Public Function ApplyValidationListToColumn(aTable As ListObject, _
                    targetColumn As Variant, _
                    fieldsString As String, _
                    Optional isErrorSuppressed As Boolean = False)
    ' This method inserts a validation list with values parsed from
    ' the fieldsString, into the cells of the targetColumn as dropdown
    ' lists. USE: combine with conditional formatting. fieldsString must
    ' be a single string with options seperated by a comma and space.
            
    ' Test if DataBodyRange is empty, add a dummy row as neeeded
    Dim tempRow As Boolean
    tempRow = CreateTempRow(aTable)
    
    ' This region is what does the actual creation of the dropdown list
    With aTable.ListColumns(targetColumn).DataBodyRange.Validation
        .Delete 'Reset previous data validation list for targetColumn
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=fieldsString
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Disallowed Input"
        .InputMessage = ""
        .ErrorMessage = "Please select from the options: " & fieldsString
        .ShowInput = True
        .ShowError = isErrorSuppressed
    End With
    
    ' Remove dummy row needed for adding to empty table
    If tempRow Then aTable.ListRows.Item(1).Delete
    
    ApplyValidationListToColumn = fieldsString
End Function

Sub ApplyConditionalFormattingToTable(aTable As ListObject, _
                        targetColumn As Variant, _
                        fieldsString As String, _
                        stylesCollection As Collection, _
                        Optional deletePrevious As Boolean = True, _
                        Optional secondaryColumnFields As String = "")
    ' This method inserts conditional formatting based on the validation
    ' (dropdown style) list values in the targetColumn. Values must match
    ' those in fieldsString (not case sensitive). hasSecondaryFormats
    ' highlights the whole row with accent scheme, while the main condition
    ' only highlights the targetColumn values. The user must manually update
    ' preferences here in this method.
    
    ' Test for empty table
    Dim tempRow As Boolean
    tempRow = CreateTempRow(aTable)
    
    Dim choices As Variant
    Dim i As Long
    choices = ParseToArray(fieldsString)
    For i = LBound(choices) To UBound(choices)
        choices(i) = """" & choices(i) & """"
    Next

    Dim targetRange As Range, firstCellAddress As String
    Set targetRange = aTable.ListColumns(targetColumn).DataBodyRange
    firstCellAddress = "$" & Replace(targetRange(1).Address, "$", "")
    
    Dim appliedToSecondaryColumns As Boolean, secondaryColumns As Variant, secondaryRange As Range
    If secondaryColumnFields <> "" Then
        appliedToSecondaryColumns = True
        secondaryColumns = utilities.ParseToArray(secondaryColumnFields)
        Dim ws As Worksheet
        Set ws = aTable.Parent
        Set secondaryRange = ws.Range(aTable.ListColumns(secondaryColumns(0)).DataBodyRange, _
            aTable.ListColumns(secondaryColumns(1)).DataBodyRange)
        Set ws = Nothing
    End If
    
    If deletePrevious Then targetRange.FormatConditions.Delete
    If deletePrevious And appliedToSecondaryColumns Then secondaryRange.FormatConditions.Delete
    
    For i = LBound(choices) To UBound(choices)
        If appliedToSecondaryColumns = False Then
            targetRange.FormatConditions.Add _
                Type:=xlExpression, _
                Formula1:="=IF(LOWER(" & firstCellAddress & ")=" & choices(i) & ",TRUE,FALSE)"
        Else
            secondaryRange.FormatConditions.Add _
                Type:=xlExpression, _
                Formula1:="=IF(LOWER(" & firstCellAddress & ")=" & choices(i) & ",TRUE,FALSE)"
        End If
    Next

    ' NOTE: stylesCollections is a Collection of FormatConfig objects
    For i = 1 To stylesCollection.Count
        If appliedToSecondaryColumns = False Then
            With targetRange.FormatConditions(i)
                .Interior.Color = stylesCollection(i).InteriorColor
                .Font.Color = stylesCollection(i).FontColor
                .Font.Bold = stylesCollection(i).Bold
            End With
        Else
            With aTable.DataBodyRange.FormatConditions(i)
                .Interior.Color = stylesCollection(i).InteriorColor
                .Font.Color = stylesCollection(i).FontColor
                .Font.Bold = stylesCollection(i).Bold
            End With
        End If
    Next

    ' Remove dummy row needed for adding to empty table
    If tempRow Then aTable.ListRows.Item(1).Delete
    
End Sub

