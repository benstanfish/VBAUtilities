Attribute VB_Name = "tables_status"
Public Const mod_name As String = "tables_status"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.7"
Public Const module_date As String = "2024/01/18"

'==================================  GLOBALS  ==================================
' Note: GLOBAL constant names are in all-caps.

Public Const LOGOCELL = "F1"
Public Const METAINPUTCOLOR = 15461355

Public Const TABLEANCHORCELL = "A37"           'Coordinate with the MetaRegion
Public Const TABLENAME = "StatusTable"         'No spaces permitted
Public Const TABLESTYLENAME = "Status Table"
Public Const TABLERULECOLOR = webcolors.LIGHTGRAY
Public Const ISTABLERULED = True

Public Const LASTSAVEDCELL = "B6"

Public Const HEADERCOLUMNNAMES = _
    "ID, Discipline, Description, Action, Resolution, Date, Days Open, Status"
Public Const HEADERCOLUMNWIDTHS = "10, 20, 45, 35, 35, 12, 12, 12"
Public Const HEADERFILLCOLOR = webcolors.DODGERBLUE

Public Const WRAPTEXTCOLUMNS = "Discipline, Description, Action, Resolution"

Public Const VALIDATIONCOLUMN = "Status"
Public Const VALIDATIONLIST = "Urgent, Pending, Resolved"

'===============================  MAIN METHODS  ================================

Public Sub RefreshMetaRegion()
    
    InsertPOAMMetadata

End Sub

Private Sub OverwriteWithNewPOAM()

    GlobalDelete ActiveSheet
    
    ActiveWindow.DisplayGridlines = False
    
    InsertPOAMMetadata
    CreateNewTable
    CopyLogos ActiveSheet, , 27
    CreateEmailTeamButton ActiveSheet
    
End Sub

Public Sub CreateNewPOAMSheet()
    'WARNING: This method needs to copy the Worksheet_Change() from
    ' "POAM Log" to the newly created Worksheet. This can be
    ' problematic if a person's VBA is not set to trust the VBA object model.

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim sht As Worksheet
    Set sht = Sheets.Add(After:=Sheets(Sheets.Count))
    sht.Name = "POAM" & IterateSheetName("POAM")

    ActiveWindow.DisplayGridlines = False
    InsertPOAMMetadata
    CreateNewTable
    CopyLogos ActiveSheet, , 27
    CreateEmailTeamButton ActiveSheet

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

Private Sub InsertPOAMMetadata()

    Dim sht As Worksheet
    Dim projData As Variant
    Dim pdtData As Variant
    Dim schedData As Variant
    
    Dim inputColor As Long
    inputColor = METAINPUTCOLOR
    
    Dim projRange As Range
    Dim pdtRange As Range
    Dim schedRange As Range
    Dim fullRange As Range
    
    projData = Array("Project:", "P2:", "A/E:")
    pdtData = Array(ChrW(9660) & " Project Delivery Team", "PM:", "DM:", "TL:", _
                "Arch:", "Civ:", "Str:", "Elec:", "Mech:", "Fire:", "TSvc:", _
                "Cost:", "VE:", "Env:", "Sust:", "COS/CX:")
    schedData = Array(ChrW(9660) & " Schedule Overview", "", _
                "Kickoff", "Planning", "Concept", "Intermed", "Final", _
                "Backcheck", "RTA", "Advertise", "Award")
    
    With Range("A1")
        .Value = "Plan of Action and Milestones (POAM)"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    Set projRange = ActiveSheet.Range("A3").Resize(UBound(projData) + 1)
    Set pdtRange = projRange(projRange.Rows.Count + 3).Resize(UBound(pdtData) + 1)
    Set schedRange = pdtRange(pdtRange.Rows.Count + 2).Resize(UBound(schedData) + 1)
    Set fullRange = Range(projRange(1), schedRange(schedRange.Rows.Count))
    
    fullRange.Clear
    fullRange.ClearOutline
    
    ' Insert static metadata
    projRange = WorksheetFunction.Transpose(projData)
    pdtRange = WorksheetFunction.Transpose(pdtData)
    schedRange = WorksheetFunction.Transpose(schedData)
    
    Dim arr As Variant
    arr = Array("Start", "End", ChrW(9661) & " Duration:")
    For i = LBound(arr) To UBound(arr)
        schedRange(2).Offset(0, i + 1) = arr(i)
    Next
    
    Range(LASTSAVEDCELL).Offset(0, -1).Value = "Saved:"

    ' Format Project Region
    projRange.Font.Bold = True
    projRange(projRange.Rows.Count).Offset(1, 0).Font.Bold = True
    For i = 1 To 2
        With projRange.Offset(0, i)
            .Interior.Color = inputColor
            With .Borders(xlInsideHorizontal)
                .Color = vbWhite
                .Weight = xlThin
            End With
            .HorizontalAlignment = xlHAlignLeft
        End With
    Next
    With projRange(1).Offset(0, 1)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' Format PDT Region
    pdtRange(1).Font.Bold = True
    Dim teamRange As Range
    Set teamRange = Range(pdtRange(2), pdtRange(pdtRange.Rows.Count)).Offset(0, 1)
    For i = 0 To 1
        With teamRange.Offset(0, i)
            .Interior.Color = inputColor
            With .Borders(xlInsideHorizontal)
                .Color = vbWhite
                .Weight = xlThin
            End With
            .HorizontalAlignment = xlHAlignLeft
        End With
    Next
    With teamRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=INDIRECT(""contacts[Name]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = False
    End With
    teamRange.Offset(0, 1).FormulaR1C1 = _
        "=IFERROR(HYPERLINK(XLOOKUP(RC2,contacts[Name],contacts[Email])),"""")"
    Range(pdtRange(2), pdtRange(pdtRange.Rows.Count + 1)).Rows.Group
    
    
    ' Format Schedule Region
    schedRange(1).Font.Bold = True
    With Range(schedRange(3), schedRange(schedRange.Rows.Count).Offset(0, 1)).Offset(0, 1)
        .Interior.Color = inputColor
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "mm/dd/yy;@"
        With .Borders(xlInsideHorizontal)
            .Color = vbWhite
            .Weight = xlThin
        End With
    End With
    With Range(schedRange(3).Offset(0, 3), schedRange(schedRange.Rows.Count).Offset(0, 3))
        .Formula2R1C1 = _
            "=IF(AND(RC[-2]<>"""",RC[-1]<>"""",DAYS(RC[-1],RC[-2])>0),DAYS(RC[-1],RC[-2]),"""")"
        .HorizontalAlignment = xlHAlignLeft
    End With
    Range(schedRange(2), schedRange(schedRange.Rows.Count)).Rows.Group

    ' Project Amount (PA) Cell
    Dim paRange As Range
    Set paRange = projRange(2).Offset(0, 3)
    paRange = ChrW(9665) & " PA:"
    With paRange.Offset(0, -1)
        .NumberFormat = "$#,###"
        .HorizontalAlignment = xlHAlignRight
        .Interior.Color = inputColor
        For i = xlEdgeLeft To xlEdgeBottom
            With .Borders(i)
                .Color = vbWhite
                .Weight = xlThin
            End With
        Next
    End With
    
    ' Tech Lead LCC Cell
    Dim lccRange As Range
    Set lccRange = pdtRange(3).Offset(0, 3)
    lccRange = ChrW(9661) & " LCC:"
    With lccRange.Offset(1, 0)
        .HorizontalAlignment = xlHAlignLeft
        .Interior.Color = inputColor
        For i = xlEdgeLeft To xlEdgeBottom
            With .Borders(i)
                .Color = vbWhite
                .Weight = xlThin
            End With
        Next
    End With

    'Insert my module information
    With ActiveSheet.Range("H5")
        .Value = "Status v" & tables_status.module_version
        .HorizontalAlignment = xlHAlignRight
        .Font.Size = 9.5
    End With
    
    ' Set initial column widths
    columnWidths = Array(10, 25, 45)
    For i = LBound(columnWidths) To UBound(columnWidths)
        ActiveSheet.Columns(i + 1).EntireColumn.ColumnWidth = columnWidths(i)
    Next

    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

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
    
        Set otherColumns = aTable.DataBodyRange
        otherColumns.FormatConditions.Delete
        
        For i = LBound(choices) To UBound(choices)
            otherColumns.FormatConditions.Add Type:=xlExpression, _
                Formula1:="=IF(LOWER(" & firstCell & ")=" & choices(i) & ",TRUE,FALSE)"
        Next
    
        With otherColumns.FormatConditions(1)
            '<URGENT>
            .Interior.Color = webcolors.LAVENDERBLUSH
            '.Font.Color = RGB(230, 0, 0)
            .Font.Color = webcolors.FIREBRICK
            .Font.Bold = False
        End With
    
        With otherColumns.FormatConditions(2)
            '<PENDING>
            .Interior.Color = webcolors.LIGHTYELLOW
            .Font.Color = ContrastText(.Interior.Color)
            .Font.Bold = False
        End With
        
        With otherColumns.FormatConditions(3)
            '<RESOLVED>
            .Interior.Color = webcolors.MINTCREAM
            .Font.Color = webcolors.MEDIUMSEAGREEN
            '.Font.Color = RGB(0, 176, 80)
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
        '<URGENT>
        .Interior.Color = webcolors.ORANGERED
        '.Interior.Pattern = xlUp
        '.Interior.PatternColor = webcolors.FIREBRICK
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = True
    End With

    With statusColumn.FormatConditions(2)
        '<PENDING>
        .Interior.Color = webcolors.YELLOW
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = True
    End With

    With statusColumn.FormatConditions(3)
        '<RESOLVED>
        .Interior.Color = RGB(0, 176, 80)   ' Green
        '.Interior.Color = webcolors.LIMEGREEN
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = True
    End With

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub

'===========================  INSERT DEFAULT VALUES  ===========================

Public Sub InsertDefaultValues(aTable As ListObject)

    If aTable.ListRows.Count <> 0 Then  ' Prevents error with empty DataBodyRange
        
        On Error Resume Next
        
        AutoincrementIDs aTable, "ID"

        Dim statusColumn As Range
        Set statusColumn = aTable.ListColumns("Status").DataBodyRange
        If Not statusColumn Is Nothing Then
            For Each aCell In statusColumn
                If aCell.Value = "" Then aCell.Value = "Pending"
            Next
        End If

        Dim dateColumn As Range
        Set dateColumn = aTable.ListColumns("Date").DataBodyRange
        If Not dateColumn Is Nothing Then
            For Each aCell In dateColumn
                If aCell.Value = "" Then aCell.Value = Now
            Next
            dateColumn.NumberFormat = "mm/dd/yyyy"
        End If
        
        Dim resolutionColumn As Range
        Set resolutionColumn = aTable.ListColumns("Resolution").DataBodyRange
        If Not resolutionColumn Is Nothing And Not statusColumn Is Nothing Then
            For i = 1 To aTable.ListRows.Count
                If resolutionColumn(i) <> "" Then statusColumn(i).Value = "Resolved"
            Next
        End If
        
        WriteDaysOpen aTable
        
        On Error GoTo 0
        
    End If
End Sub
