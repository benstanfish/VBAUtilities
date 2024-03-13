Attribute VB_Name = "tables_status"
Public Const mod_name As String = "Status Table"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.9.1"
Public Const module_date As Date = #3/11/2024#

'==================================  GLOBALS  ==================================
' Note: GLOBAL constant names are in all-caps.

Public Const LOGOCELL = "F1"
Public Const METAINPUTCOLOR = webcolors.ALICEBLUE

Public Const TABLEANCHORCELL = "A30"           'Coordinate with the MetaRegion
Public Const TABLENAME = "StatusTable"         'No spaces permitted
Public Const TABLESTYLENAME = "Status Table"
Public Const TABLERULECOLOR = webcolors.LIGHTGRAY
Public Const ISTABLERULED = True

Public Const LASTSAVEDCELL = "B6"

Public Const HEADERCOLUMNNAMES = _
    "ID, Topic, Description, Action, Resolution, Date, Days Open, Status"
Public Const HEADERCOLUMNWIDTHS = "10, 20, 45, 25, 35, 12, 12, 12"
Public Const HEADERFILLCOLOR = webcolors.DODGERBLUE

Public Const WRAPTEXTCOLUMNS = "Discipline, Description, Action, Resolution"

Public Const VALIDATIONCOLUMN = "Status"
Public Const VALIDATIONLIST = "Urgent, Pending, Resolved"

'===============================  MAIN METHODS  ================================

Public Sub RefreshMetaRegion()
    
    CreateMetaRegion

End Sub

Private Sub OverwriteWithNewPOAM()

    GlobalDelete ActiveSheet
    
    ActiveWindow.DisplayGridlines = False
    
    CreateMetaRegion
    CreateNewTable
    CopyLogos ActiveSheet, , 27

    ActiveSheet.Range(TABLEANCHORCELL).Select
    
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

    CreateMetaRegion
    CreateNewTable
    CopyLogos ActiveSheet, , 27

    On Error GoTo dump
    CopyWorksheetChangeCode sht
    On Error GoTo 0

NormalFlow:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ActiveSheet.Range(TABLEANCHORCELL).offset(1, 1).Select
    Exit Sub
dump:
    Resume NormalFlow
End Sub

'==============================  METADATA HEADER  ==============================

Private Sub CreateMetaRegion()

    ActiveWindow.DisplayGridlines = False
    Range("A1").ClearOutline
    
    Dim rowOffset As Long, colOffset As Long
    rowOffset = 2
    colOffset = 1
    
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    Dim interiorColor As Long
    interiorColor = METAINPUTCOLOR

    ' SHEET TITLE REGION
    Dim titleRange As Range
    Set titleRange = sht.Range("A1")
    With titleRange
        .Value = "Plan of Action and Milestones (POAM)"
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' PROJECT DATA REGION
    Dim projectFields As Variant
    Dim projectFieldRange As Range
    Dim projectInputRange As Range
    projectFields = Array("Project", "P2", "Type", "Saved")
    Set projectFieldRange = sht.Range(titleRange.offset(rowOffset, 0), _
        titleRange.offset(rowOffset + UBound(projectFields), 0))
    With projectFieldRange
        .Value = WorksheetFunction.Transpose(projectFields)
        .Font.Bold = True
    End With
    Set projectInputRange = Range(projectFieldRange.offset(0, 1), projectFieldRange.offset(0, 2))
    With projectInputRange
        .Interior.Color = interiorColor
        .Font.Color = ContrastText(.Interior.Color)
        .Borders(xlInsideHorizontal).Color = webcolors.WHITE
        With .Item(1)
            .Font.Size = 12
            .Font.Bold = True
        End With
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    ' PROJECT TEAM REGION
    Dim pdtFields As Variant
    Dim pdtFieldRange As Range
    Dim pdtInputRange As Range
    Dim pdtInputHeaders As Variant
    pdtFields = Array("PM", "DM", "TL", _
        "Geotech", "Civil", "Struct", "Arch", "FP", "Mech", "Elec", "Comm", "Cyber", _
        "Env", "Sus", "Cost", "TS", "VE", "RO", "MCX")
    pdtInputHeaders = Array("JED", "A/E")
    ' Add two additional rows as placeholders for the region title and subtitle, added later
    Set pdtFieldRange = sht.Range(projectFieldRange(projectFieldRange.Rows.Count).offset(rowOffset + 2, 0), _
        projectFieldRange(projectFieldRange.Rows.Count).offset(rowOffset + 2 + UBound(pdtFields), 0))
    With pdtFieldRange
        .Value = WorksheetFunction.Transpose(pdtFields)
        .Font.Bold = False
    End With
    With pdtFieldRange(1).offset(-2, 0)
        .Value = ChrW(9660) & " Project Delivery Team"
        .Font.Bold = True
    End With
    Set pdtInputRange = Range(pdtFieldRange.offset(0, 1), pdtFieldRange.offset(0, UBound(pdtInputHeaders) + 1))
    With pdtInputRange
        .Interior.Color = interiorColor
        .Font.Color = ContrastText(.Interior.Color)
        .Borders(xlInsideHorizontal).Color = webcolors.WHITE
        .Borders(xlInsideVertical).Color = webcolors.WHITE
    End With
    With pdtInputRange(0, 1).Resize(1, UBound(pdtInputHeaders) + 1)
        .Value = pdtInputHeaders
        .Font.Bold = False
    End With
    
    ' PROJECT SCHEDULE REGION
    Dim schedFields As Variant
    Dim schedFieldRange As Range
    Dim schedInputRange As Range
    Dim schedInputHeaders As Variant
    schedFields = Array("Kickoff", "Charrette", "  " & ChrW(9500) & " " & "Draft PDCR", "  " & ChrW(9500) & " " & "Final PDCR", _
        "  " & ChrW(9492) & " " & "Backcheck", "Concept", "  " & ChrW(9492) & " " & "OBR", "Intermediate", "  " & ChrW(9492) & " " & "OBR", _
        "Final", "  " & ChrW(9492) & " " & "OBR", "Backcheck", "BCOES Cert", "RTA", "Advertise")
    schedInputHeaders = Array("Submittal", "Start", "End", ChrW(9661) & " Duration")
    Set schedFieldRange = pdtInputRange(1, pdtInputRange.Columns.Count).offset(0, colOffset).Resize(UBound(schedFields, 1) + 1, 1)
    With schedFieldRange
        .Value = WorksheetFunction.Transpose(schedFields)
    End With
        With schedFieldRange(1).offset(-2, 0)
        .Value = ChrW(9660) & " Project Schedule"
        .Font.Bold = True
    End With
    Set schedInputRange = Range(schedFieldRange.offset(0, 1), schedFieldRange.offset(0, UBound(schedInputHeaders) + 1))
    With schedInputRange(0, 1).Resize(1, UBound(schedInputHeaders) + 1)
        .Value = schedInputHeaders
        .Font.Bold = False
    End With
    With Range(schedInputRange(1), schedInputRange(schedInputRange.Rows.Count, 3))
        .NumberFormat = "d mmm yy"
        .Interior.Color = interiorColor
        .Font.Color = ContrastText(.Interior.Color)
        .Borders(xlInsideHorizontal).Color = webcolors.WHITE
        .Borders(xlInsideVertical).Color = webcolors.WHITE
    End With
    With Range(schedInputRange(1).offset(0, 3), schedInputRange(schedInputRange.Rows.Count, 4))
        .Formula2R1C1 = "=DAYS(RC[-1],RC[-2])"
        .NumberFormat = "0;-0;;@"
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    ' GROUP PDT AND SCHEDULE REGION
    Dim pdtRegion As Range
    Dim schedRegion As Range
    Set pdtRegion = sht.Range(pdtFieldRange(1).offset(-1, 0), _
        pdtInputRange(pdtInputRange.Cells.Count))
    Set schedRegion = sht.Range(schedFieldRange(1).offset(-1, 0), _
        schedInputRange(schedInputRange.Cells.Count))
    If pdtRegion.Rows.Count >= schedRegion.Rows.Count Then pdtRegion.Rows.Group Else schedRegion.Rows.Group
    sht.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    
    ' WRITE VERSION INFO
    With sht.Range("H" & Range(TABLEANCHORCELL).Row - 1)
        .HorizontalAlignment = xlHAlignRight
        .Font.Size = 9
        .Value = mod_name & " v" & module_version
    End With
    
    ' PRINT OUT CALCULATED "TABLEANCHORCELL"
    'Debug.Print pdtRegion(1).offset(pdtRegion.Rows.Count - 1 + rowOffset, 0).Address

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
