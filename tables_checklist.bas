Attribute VB_Name = "tables_checklist"
Public Const mod_name As String = "tables_checklist"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.4"
Public Const module_date As String = "2024/01/18"

'==================================  GLOBALS  ==================================
' Note: GLOBAL constant names are in all-caps.

Public Const LOGOCELL = "F1"
Public Const METAINPUTCOLOR = webcolors.GAINSBORO

Public Const TABLEANCHORCELL = "A14"           'Coordinate with the MetaRegion
Public Const TABLENAME = "ReviewList"         'No spaces permitted
Public Const TABLESTYLENAME = "Simple Table"
Public Const TABLERULECOLOR = webcolors.LIGHTGRAY
Public Const ISTABLERULED = False

Public Const LASTSAVEDCELL = "B12"

Public Const HEADERCOLUMNNAMES = _
    "Category, Topic, ID, Item, Status, Comment"
Public Const HEADERCOLUMNWIDTHS = "15, 20, 5, 40, 9, 35"
Public Const HEADERFILLCOLOR = webcolors.DODGERBLUE

Public Const WRAPTEXTCOLUMNS = "Category, Topic, Item, Comment"

Public Const VALIDATIONCOLUMN = "Status"
Public Const VALIDATIONLIST = "Yes, No, Unknown, NA"

'===============================  MAIN METHODS  ================================

Public Sub RefreshMetaRegion()
    
    CreateMetaRegion

End Sub

Private Sub OverwriteWithNewChecklist()

    GlobalDelete ActiveSheet

    ActiveWindow.DisplayGridlines = False

    CreateMetaRegion
    CreateNewTable
    ApplyLadderLines
    CopyLogos ActiveSheet, , 9
    
End Sub

Public Sub CreateNewChecklistSheet()
    'WARNING: This method needs to copy the Worksheet_Change() from
    ' "POAM Log" to the newly created Worksheet. This can be
    ' problematic if a person's VBA is not set to trust the VBA object model.

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim sht As Worksheet
    Set sht = Sheets.Add(After:=Sheets(Sheets.Count))
    sht.Name = "Checklist" & IterateSheetName("Checklist")

    ActiveWindow.DisplayGridlines = False
    CreateMetaRegion
    CreateNewTable
    ApplyLadderLines
    CopyLogos ActiveSheet, , 9

    On Error GoTo dump
    CopyWorksheetChangeCode sht
    On Error GoTo 0

NormalFlow:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ActiveSheet.Range(TABLEANCHORCELL).Offset(1, 1).Select
    Exit Sub
dump:
    Resume NormalFlow
End Sub

'==============================  METADATA HEADER  ==============================

Private Sub CreateMetaRegion()
    
    Dim titleData As Variant
    Dim projData As Variant
    Dim reviewData As Variant
    
    Dim titleRange As Range
    Dim projRange As Range
    Dim reviewRange As Range
    Dim fullRange As Range
    
    With ActiveSheet.Range("F5")
        .Value = "Checklist v" & tables_checklist.module_version
        .HorizontalAlignment = xlHAlignRight
    End With
    
    titleData = Array("Checklist Title", "Checklist Subtitle")
    projData = Array("Project:", "P2:", "Location:", "Client:", "Phase:", "Doc. Date:")
    reviewData = Array("Reviewer:", "Saved:")
    
    Set titleRange = ActiveSheet.Range("A1").Resize(UBound(titleData) + 1)
    Set projRange = titleRange(titleRange.Rows.Count + 2).Resize(UBound(projData) + 1)
    Set reviewRange = projRange(projRange.Rows.Count + 2).Resize(UBound(reviewData) + 1)
    Set fullRange = Range(titleRange(1), reviewRange(reviewRange.Rows.Count))

    fullRange.Clear
    fullRange.ClearOutline

    ' Insert static metadata
    titleRange = WorksheetFunction.Transpose(titleData)
    projRange = WorksheetFunction.Transpose(projData)
    reviewRange = WorksheetFunction.Transpose(reviewData)
    
    With titleRange(1)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With projRange
        .Font.Bold = True
        For i = 1 To 3
            With .Offset(0, i)
                .Interior.Color = webcolors.GAINSBORO
                .Borders(xlInsideHorizontal).Color = vbWhite
                .Borders(xlInsideHorizontal).Weight = xlThin
            End With
        Next
    End With
    
    For i = 3 To 6
        projRange(i).Offset(0, 3).Interior.Pattern = xlNone
    Next
    
    With projRange(1).Offset(0, 1)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With projRange(2).Offset(0, 3)
        .Borders(xlEdgeLeft).Color = vbWhite
        .Borders(xlEdgeLeft).Weight = xlThin
        .NumberFormat = "$#,###"
    End With
    
    With projRange(2).Offset(0, 4)
        .Value = ChrW(9665) & " PA"
        .HorizontalAlignment = xlHAlignLeft
    End With

    With reviewRange(1)
        .Font.Bold = True
        For i = 1 To 3
            With .Offset(0, i)
                .Interior.Color = webcolors.GAINSBORO
                .Borders(xlInsideHorizontal).Color = vbWhite
                .Borders(xlInsideHorizontal).Weight = xlThin
            End With
        Next
        With .Offset(0, 3)
            .Borders(xlEdgeLeft).Color = vbWhite
            .Borders(xlEdgeLeft).Weight = xlThin
        End With
    End With
    
    With reviewRange(1).Offset(0, 4)
        .Value = ChrW(9665) & " Email"
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    reviewRange(2).Font.Bold = True
    WriteSaveDate ActiveSheet, LASTSAVEDCELL

End Sub

'==========================  CONDITIONAL FORMATTING  ===========================

Sub ApplyStatusFormats(aTable As ListObject, _
                            Optional targetColumn As String = VALIDATIONCOLUMN, _
                            Optional selectionSet As String = VALIDATIONLIST, _
                            Optional hasSecondaryFormats As Boolean = True)
    'NOTE: This method inserts conditional formatting based on the validation "dropdown" list
    ' values in the targetColumn. Values must match those in selectionSet (not case sensitive).
    ' hasSecondaryFormats highlights the whole row with accent scheme, while the main
    ' condition only highlights the targetColumn values. The user must manually update
    ' preferences here in this method.
    
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        temporaryRow = True
    End If
    
    Dim choices As Variant
    Dim i As Long

    choices = ParseToArray(VALIDATIONLIST)
    For i = LBound(choices) To UBound(choices)
        choices(i) = """" & choices(i) & """"
    Next

    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange


    Dim firstCell As String
    firstCell = "$" & Replace(statusColumn(1).Address, "$", "")

    '-------------  SECONDARY CONDITIONS  -------------
    If hasSecondaryFormats Then
    
        Set otherColumns = Union(aTable.ListColumns("ID").DataBodyRange, aTable.ListColumns("Item").DataBodyRange, aTable.ListColumns("Comment").DataBodyRange)
        otherColumns.FormatConditions.Delete
        
        For i = LBound(choices) To UBound(choices)
            otherColumns.FormatConditions.Add Type:=xlExpression, _
                Formula1:="=IF(LOWER(" & firstCell & ")=" & choices(i) & ",TRUE,FALSE)"
        Next
    
        With otherColumns.FormatConditions(1)
            '<YES>
            .Interior.Color = webcolors.MINTCREAM
            .Font.Color = webcolors.MEDIUMSEAGREEN
            .Font.Bold = False
        End With
    
        With otherColumns.FormatConditions(2)
            '<NO>
            .Interior.Color = webcolors.LAVENDERBLUSH
            '.Font.Color = RGB(230, 0, 0)
            .Font.Color = webcolors.FIREBRICK
            .Font.Bold = False
        End With
    
        With otherColumns.FormatConditions(3)
            '<Unknown>
            .Interior.Color = webcolors.LIGHTYELLOW
            .Font.Color = ContrastText(.Interior.Color)
            .Font.Bold = False
        End With
        
        With otherColumns.FormatConditions(4)
            '<NA>
            .Interior.Color = RGB(240, 240, 240)
            .Font.Color = RGB(150, 150, 150)
            .Font.Bold = False
        End With

    End If

    '---------------  MAIN CONDITIONS  ----------------
    statusColumn.FormatConditions.Delete

    For i = LBound(choices) To UBound(choices)
        statusColumn.FormatConditions.Add Type:=xlExpression, _
            Formula1:="=IF(LOWER(" & firstCell & ")=" & choices(i) & ",TRUE,FALSE)"
    Next

    With statusColumn.FormatConditions(1)
        '<YES>
        .Interior.Color = RGB(0, 176, 80)   ' Green
        '.Interior.Color = webcolors.LIMEGREEN
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = False
    End With

    With statusColumn.FormatConditions(2)
        '<NO>
        .Interior.Color = webcolors.ORANGERED
        '.Interior.Pattern = xlUp
        '.Interior.PatternColor = webcolors.FIREBRICK
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = False
    End With

    With statusColumn.FormatConditions(3)
        '<Unknown>
        .Interior.Color = webcolors.YELLOW
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = False
    End With

    With statusColumn.FormatConditions(4)
        '<NA>
        .Interior.Color = RGB(240, 240, 240)
        .Font.Color = RGB(150, 150, 150)
        .Font.Bold = False
    End With

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub

'============================  LADDER LINE FORMATS  ============================

Public Sub ApplyLadderLines()

    Dim aTable As ListObject
    Set aTable = ActiveSheet.ListObjects(1)
    
    Dim hasDummyRow As Boolean
    Dim tempRow As ListRow
    
    If aTable.ListRows.Count = 0 Then
        Set tempRow = aTable.ListRows.Add
        hasDummyRow = True
    End If
    
    ' Reboot all lines in table, except the top
    For Each i In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                        xlInsideVertical, xlInsideHorizontal)
        aTable.DataBodyRange.Borders(i).LineStyle = xlNone
    Next
    
    ' Apply a border around the whole DataBodyRange
    aTable.DataBodyRange.BorderAround Weight:=xlMedium, Color:=HEADERFILLCOLOR
    
    Dim categories As Range
    Dim topics As Range
    Dim otherRegion As Range
    Dim statuses As Range
    
    Set categories = aTable.ListColumns("Category").DataBodyRange
    Set topics = aTable.ListColumns("Topic").DataBodyRange
    Set otherRegion = ActiveSheet.Range(aTable.ListColumns("ID").DataBodyRange(1), _
        aTable.ListColumns(aTable.ListColumns.Count).DataBodyRange(aTable.DataBodyRange.Rows.Count))
    Set statuses = aTable.ListColumns("Status").DataBodyRange
    
    For Each aBorder In Array(xlEdgeLeft, xlInsideHorizontal)
        With otherRegion.Borders(aBorder)
            .Color = webcolors.SILVER
            .Weight = xlThin
        End With
    Next
    
    For Each aBorder In Array(xlEdgeLeft, xlEdgeRight)
        With statuses.Borders(aBorder)
            .Color = webcolors.SILVER
            .Weight = xlThin
        End With
    Next
    
    With topics.Borders(xlEdgeLeft)
        .Color = webcolors.SILVER
        .Weight = xlThin
    End With
    
    For Each aCell In topics
        If aCell <> topics(1) And aCell <> "" Then
            With aCell.Borders(xlEdgeTop)
                .Color = webcolors.SILVER
                .Weight = xlThin
            End With
        End If
    Next
    
    For Each aCell In categories
        If aCell <> categories(1) And aCell <> "" Then
            With aCell.Borders(xlEdgeTop)
                .Color = webcolors.SILVER
                .Weight = xlThin
            End With
        End If
    Next
    
    
    With categories
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    With topics
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    With otherRegion
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    
    aTable.DataBodyRange.Font.Size = 9

    If hasDummyRow Then
        tempRow.Delete
        hasDummyRow = False
    End If

End Sub

