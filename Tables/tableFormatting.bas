Attribute VB_Name = "tableFormatting"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "tableFormatting"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "2.1.5"
Public Const module_date As Date = #6/6/2024#
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

Public Const DEFAULT_BORDER_COLOR As Long = webcolors.DODGERBLUE
Public Const DEFAULT_LADDER_LINE_COLOR As Long = webcolors.GAINSBORO

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

Public Sub ApplyBorderAroundTable(aTable As ListObject, _
                                  Optional borderColor As Long = DEFAULT_BORDER_COLOR, _
                                  Optional borderWeight As XlBorderWeight = xlMedium, _
                                  Optional hasBorder As Boolean = True)
    ' Applies a border around the entire table. NOTE: you can remove the border by setting
    ' the hasBorder value to FALSE
    
    Dim lineStyleValue As XlLineStyle
    If hasBorder Then lineStyleValue = xlContinuous Else lineStyleValue = xlLineStyleNone
    aTable.HeaderRowRange.BorderAround LineStyle:=lineStyleValue, Weight:=borderWeight, Color:=borderColor
    aTable.DataBodyRange.BorderAround LineStyle:=lineStyleValue, Weight:=borderWeight, Color:=borderColor
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


'***********************************************************************
'                      Ladder Lines & Continuations
'***********************************************************************
Private Sub GetAllHBreakCells()
    ActiveWindow.View = xlPageBreakPreview
    For i = 1 To ActiveSheet.HPageBreaks.Count
        Debug.Print ActiveSheet.HPageBreaks(i).Location.Address, ActiveSheet.HPageBreaks(i).Location.Row
    Next
End Sub

Public Sub SetContinuationText(aTable As ListObject, _
                               Optional targetColumn As Variant = "", _
                               Optional hideContinuations As Boolean = False)
    
    
    Dim ws As Worksheet
    Dim targetColumnNumber As Long
    Dim columnAHBreakCell As Range
    Dim currHBreakCell As Range
    Dim prevNonEmptyCell As Range
    Dim suffix As String

    suffix = " (cont.)"
    
    Call utilities.MemorySaver
    ' Clear out continuations of previous runs
    If targetColumn = "" Then targetColumn = 1
    For Each aCell In aTable.ListColumns(targetColumn).DataBodyRange
        If Right(aCell, Len(suffix)) = suffix Then aCell.Value = ""
    Next

    ' If you set hideContinuations = True then the continuations will be deleted
    ' based on the code in the previous statement above.
    If hideContinuations = False Then
        ' You must set to Page Break View for the HPageBreaks function to work.
        ActiveWindow.View = xlPageBreakPreview
        
        ' Set continuation in all horizontal page break cells of targetColumn
        Set ws = aTable.Parent
        targetColumnNumber = aTable.ListColumns(targetColumn).DataBodyRange.Column
        For i = 1 To ws.HPageBreaks.Count
            Set columnAHBreakCell = ws.HPageBreaks(i).Location
            Set currHBreakCell = ws.Cells(columnAHBreakCell.Row, targetColumnNumber)
            Set prevNonEmptyCell = currHBreakCell.End(xlUp)
            
            If (Right(currHBreakCell, Len(suffix)) = suffix) Or _
                (currHBreakCell.Value = "") Then
                currHBreakCell.Value = prevNonEmptyCell.Value & suffix
            End If
        Next
    
        Set ws = Nothing
        ActiveWindow.View = xlNormalView
    End If
    Call utilities.MemoryRestore
End Sub

Public Sub ApplyLadderLinesToTable(aTable As ListObject, _
                                   firstLadderColumn As Variant, _
                                   firstRestColumn As Variant, _
                                   Optional ladderLineColor As Long = DEFAULT_LADDER_LINE_COLOR, _
                                   Optional hasVerticalBorderInLadderRange As Boolean = True, _
                                   Optional preLadderRegionIsRuled As Boolean = True)
    ' This function creates ladder lines in a given table. The TableStyle of
    ' the table should not have any inside lines, as this cannot override those
    Dim ws As Worksheet
    Dim ladderRange As Range
    Dim restRange As Range
    Dim preLadderRange As Range
    
    Call utilities.MemorySaver
    
    Set ws = aTable.Parent
    Set ladderRange = ws.Range(aTable.ListColumns(firstLadderColumn).DataBodyRange, _
                        aTable.ListColumns(firstRestColumn).DataBodyRange.Offset(, -1))

    Set restRange = ws.Range(aTable.ListColumns(firstRestColumn).DataBodyRange, _
                        aTable.DataBodyRange.Columns(aTable.DataBodyRange.Columns.Count))
    If ladderRange.Columns(1).Column > aTable.ListColumns(1).DataBodyRange.Column Then
        Set preLadderRange = ws.Range(aTable.ListColumns(1).DataBodyRange, _
            aTable.ListColumns(firstLadderColumn).DataBodyRange.Offset(, -1))
    End If
    
    ' Reset all interior lines
    For i = xlInsideVertical To xlInsideHorizontal
        aTable.DataBodyRange.Borders(i).LineStyle = xlLineStyleNone
    Next
    
    Dim ladderRangeFields As Range
    Set ladderRangeFields = ws.Range(ladderRange(1).Offset(-1, 0), restRange(1).Offset(-1, -1))
    For Each aCell In ladderRangeFields
        Call tableFormatting.SetContinuationText(aTable:=aTable, targetColumn:=aCell.Value)
    Next
    
    Dim currCol As Long, currColVal As String, rmdCol As Long
    Dim currRow As Range
    For i = 1 To ladderRange.Columns.Count
        currCol = i
        rmdCol = ladderRange.Columns.Count - currCol
        For j = 1 To ladderRange.Columns(i).Rows.Count
            currColVal = ladderRange.Columns(i).Rows(j).Value
            If currColVal <> "" And Right(currColVal, Len("(cont.)")) <> "(cont.)" Then
                Set currRow = ws.Range(ladderRange.Columns(i).Rows(j), _
                    ladderRange.Columns(i).Rows(j).Offset(0, rmdCol))
                With currRow.Borders(xlEdgeTop)
                    .Color = DEFAULT_LADDER_LINE_COLOR
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
        Next
    Next
    If hasVerticalBorderInLadderRange Then
        With ladderRange.Borders(xlInsideVertical)
            .Color = DEFAULT_LADDER_LINE_COLOR
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
    
    ' Format the Rest range
    With restRange.Borders(xlInsideHorizontal)
        .Color = DEFAULT_LADDER_LINE_COLOR
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With aTable.ListColumns(firstRestColumn).DataBodyRange.Borders(xlEdgeLeft)
        .Color = DEFAULT_LADDER_LINE_COLOR
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' Format the Ladder Range
    With aTable.ListColumns(firstLadderColumn).DataBodyRange.Borders(xlEdgeLeft)
        .Color = DEFAULT_LADDER_LINE_COLOR
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Format preLadderRange if it's not nothing
    If Not preLadderRange Is Nothing And preLadderRegionIsRuled Then
        With preLadderRange.Borders(xlInsideHorizontal)
            .Color = DEFAULT_LADDER_LINE_COLOR
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
    
    Call utilities.MemoryRestore
    
    Set ladderRange = Nothing
    Set restRange = Nothing
    Set preLadderRange = Nothing
    Set ws = Nothing
End Sub

