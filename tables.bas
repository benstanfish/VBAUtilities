Attribute VB_Name = "tables"
Public Const mod_name As String = "tables"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.9.2"
Public Const module_date As Date = #3/13/2024#

'==================================  GLOBALS  ==================================
' Note: GLOBAL constant names are in all-caps.

Public Const VERTICALPADDING = 6

'==============================  HELPER METHODS  ===============================

Sub ResetRangeIndent(rng As Range)
    If rng.IndentLevel > 0 Then rng.InsertIndent -1 * (rng.IndentLevel)
End Sub
Sub IndentRange(rng As Range, Optional indentAmount As Long = 1)
    ResetRangeIndent rng
    rng.InsertIndent indentAmount
End Sub

Function ContrastText(interiorColor As Long) As Long
    'Based on W3.org visibility recommendations:
    'https://www.w3.org/TR/AERT/#color-contrast
    
    Dim color_brightness As Double
    Dim r As Long, g As Long, b As Long
    
    b = interiorColor \ 65536
    g = (interiorColor - b * 65536) \ 256
    r = interiorColor - b * 65536 - g * 256
    
    color_brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    If color_brightness > 0.55 Then ContrastText = vbBlack Else ContrastText = vbWhite
End Function

Function ParseToArray(namedConstant As String) As Variant
    ParseToArray = Split(namedConstant, ", ")
End Function

Function ParseToLongArray(namedConstant As String) As Variant
    Dim arr As Variant
    Dim arr2 As Variant
    Dim i As Long
    
    arr = ParseToArray(namedConstant)
    ReDim arr2(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        arr2(i) = CLng(arr(i))
    Next

    ParseToLongArray = arr2
End Function

Sub CopyLogos(sht As Worksheet, Optional logoRange As String = LOGOCELL, _
                Optional moveLeft As Long = 0, Optional moveTop As Long = 0)
    Sheets("logos").Shapes(1).Copy
    sht.Paste Destination:=Range(logoRange)
    sht.Shapes(1).IncrementLeft moveLeft
    sht.Shapes(1).IncrementTop moveTop
End Sub


Sub CreateEmailTeamButton(sht As Worksheet)
    Dim shp As Shape

    Set shp = sht.Shapes.AddShape(msoShapeRoundedRectangle, 675, 10, 100, 40)
    shp.Name = "EmailButton"

    shp.ShapeStyle = msoShapeStylePreset13
    shp.Line.Visible = msoFalse
    shp.Shadow.Type = msoShadow21

    With shp.TextFrame2.TextRange
        .Characters.Text = "Copy PDT Email List to Clipboard"
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    With shp.TextFrame2.TextRange.Font
        .NameComplexScript = "Aptos"
        .NameFarEast = "Aptos"
        .Name = "Aptos"
        .Size = 10
    End With
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 3.6
        .MarginRight = 3.6
    End With
    
    shp.Select
    With Selection
        .PrintObject = msoFalse
        .Placement = xlFreeFloating
        .OnAction = "EmailTeam"
    End With

End Sub

'============================  INITIAL TABLE SETUP  ============================

Sub GlobalDelete(sht As Worksheet)
    Cells.Delete
    For i = 1 To sht.Shapes.Count
        sht.Shapes(i).Delete
    Next
End Sub

Function AutoincrementTable(baseName As String)
    Dim maxIndex As Long
    Dim sht As Worksheet
    Dim lstObj As ListObject
    For Each sht In ActiveWorkbook.Sheets
        For Each lstObj In sht.ListObjects
            If Left(lstObj.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
        Next
    Next
    If maxIndex = 0 Then AutoincrementTable = "" Else AutoincrementTable = maxIndex
End Function

Function IntersectsTable(proposedRange As Range) As Boolean
    Dim sht As Worksheet
    Dim lstObj As ListObject
    Dim doesIntersect As Boolean    'False by default
    For Each lstObj In ActiveSheet.ListObjects
        If Not Intersect(proposedRange, lstObj.Range) Is Nothing Then doesIntersect = True
    Next
    IntersectsTable = doesIntersect
End Function

Function WriteHeaders(sht As Worksheet, _
                    Optional tblAnchorCell As String = TABLEANCHORCELL, _
                    Optional headerNames As String = HEADERCOLUMNNAMES, _
                    Optional overwriteMode As Boolean = False) As Variant
    
    'NOTE: Returns the header range if successful (for next step of creating a table)
    
    Dim headerFields As Variant
    Dim headerRange As Range
    Dim i As Long
    Dim newName As String
    
    headerFields = ParseToArray(headerNames)
    Set headerRange = sht.Range(tblAnchorCell, _
        sht.Range(tblAnchorCell).offset(0, UBound(headerFields)))

    If IntersectsTable(headerRange) = False Or overwriteMode <> False Then
        For i = 0 To UBound(headerFields)
            headerRange(i + 1) = headerFields(i)
        Next
    End If
    Set WriteHeaders = headerRange
End Function

Function MakeIntoTable(headerRange As Range, sht As Worksheet, _
    Optional tblName As String = TABLENAME) As ListObject

    Dim newName As String
    newName = tblName & AutoincrementTable(tblName)
    
    If IntersectsTable(headerRange) = False Then
        sht.ListObjects.Add(xlSrcRange, headerRange, , xlYes).Name = newName
        Set MakeIntoTable = sht.ListObjects(newName)
    End If
    
End Function

Sub SetHeaderWidths(aTable As ListObject, _
                        Optional headerNameString As String = HEADERCOLUMNNAMES, _
                        Optional headerWidthString As String = HEADERCOLUMNWIDTHS)
    
    Dim headerNames As Variant
    Dim headerWidths As Variant
    
    headerNames = ParseToArray(headerNameString)
    headerWidths = ParseToLongArray(headerWidthString)
    
    On Error Resume Next
    For i = LBound(headerNames) To UBound(headerNames)
        aTable.ListColumns(headerNames(i)).Range.EntireColumn.ColumnWidth = headerWidths(i)
    Next
    On Error GoTo 0
    
End Sub

Sub CreateNewTable()

    Dim headerRange As Range
    Dim aTable As ListObject
    
    Set headerRange = WriteHeaders(ActiveSheet)
    Set aTable = MakeIntoTable(headerRange, ActiveSheet, TABLENAME)

    ' If a new table was not created (ERROR STATE), the following is not run.
    If Not aTable Is Nothing Then
        ApplyTableStyle aTable, TABLESTYLENAME
        SetHeaderWidths aTable
        InsertDropdown aTable
        InsertDropdown aTable:=aTable, targetColumn:="POAM", selectionSet:="Yes, No"
        ApplyStatusFormats aTable
    End If

End Sub

Public Sub ResetExistingTable()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo dump
    Dim aTable As ListObject
    Set aTable = ActiveSheet.ListObjects(1)
    
    WriteHeaders ActiveSheet, overwriteMode:=True
    
    ApplyTableStyle aTable
    SetHeaderWidths aTable
    ApplyComfyRowHeights aTable
    
    On Error GoTo 0
dump:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Function TableHasData(aTable As ListObject) As Boolean
    If Not aTable.DataBodyRange Is Nothing Then TableHasData = True
End Function

Sub RenumberIDs(aTable As ListObject, Optional idColumnName As String = "ID")
    ' Turn off screen updating and set enable events to false if used OnChange.
    Dim i As Long: i = 1
    On Error GoTo dump
    If TableHasData(aTable) Then
        For Each aCell In aTable.ListColumns(idColumnName).DataBodyRange
            aCell.Value = i
            i = i + 1
        Next
    End If
    'aTable.ListColumns(idColumnName).DataBodyRange.HorizontalAlignment = xlHAlighLeft
    On Error GoTo 0
dump:
End Sub

Sub AutoincrementIDs(aTable As ListObject, Optional idColumnName As String = "ID")
    ' Turn off screen updating and set enable events to false if used OnChange.
    Dim currentMax As Long
    On Error GoTo dump
    If TableHasData(aTable) Then
        For Each aCell In aTable.ListColumns(idColumnName).DataBodyRange
            If aCell = "" Then
                currentMax = WorksheetFunction.Max(aTable.ListColumns(idColumnName).DataBodyRange)
                aCell.Value = currentMax + 1
            End If
        Next
        'aTable.ListColumns(idColumnName).DataBodyRange.HorizontalAlignment = xlHAlighLeft
    End If
    On Error GoTo 0
dump:
End Sub

Sub ResetRowHeights(aTable As ListObject)
    'NOTE: This sets to default Excel row height
    If Not aTable.DataBodyRange Is Nothing Then
        aTable.DataBodyRange.WrapText = False
        For Each aRow In aTable.DataBodyRange.Rows
            aRow.EntireRow.AutoFit
        Next
    End If
End Sub

Sub ApplyComfyRowHeights(aTable As ListObject, _
    Optional columnsToWrap As String = WRAPTEXTCOLUMNS)
    
    If Not aTable.DataBodyRange Is Nothing Then
        Dim arr As Variant
        aTable.DataBodyRange.WrapText = False
        With aTable.DataBodyRange
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignLeft
        End With
        For Each aRow In aTable.DataBodyRange.Rows
            aRow.EntireRow.AutoFit
        Next
        arr = ParseToArray(columnsToWrap)
        'On Error Resume Next
        For Each thing In arr
            With aTable.ListColumns(thing).DataBodyRange
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignLeft
                .WrapText = True
            End With
        Next
        For Each aRow In aTable.DataBodyRange.Rows
            aRow.RowHeight = aRow.RowHeight + VERTICALPADDING
        Next
    End If
    On Error GoTo 0
End Sub

'===================  VALIDATION AND CONDITIONAL FORMATTING  ===================

Sub InsertDropdown(aTable As ListObject, _
                            Optional targetColumn As String = VALIDATIONCOLUMN, _
                            Optional selectionSet As String = VALIDATIONLIST, _
                            Optional suppressError As Boolean = False)
    'NOTE: This method inserts a validation list with values parsed from the selectionSet, into
    ' the cells of the targetColumn as dropdown lists. USE: combine with conditional formatting.
    ' selectionSet must be a single string with options seperated by a comma and space.
            
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        temporaryRow = True
    End If
    
    With aTable.ListColumns(targetColumn).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=selectionSet
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Disallowed Input"
        .InputMessage = ""
        .ErrorMessage = "Please select from the options: " & selectionSet
        .ShowInput = True
        .ShowError = suppressError
    End With
    
    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete
 
End Sub

Sub InsertPOAMDropdown()

    

End Sub


'===============================  TABLE STYLES  ================================

Function StyleExists(Optional styleName As String = TABLESTYLENAME) As Boolean
    Dim dummyBool As Boolean
    On Error GoTo dump
        If ActiveWorkbook.TableStyles(styleName) = styleName Then dummyBool = True
        ' Simply need to evaluate an expression to force potential error
    On Error GoTo 0
    StyleExists = True
    Exit Function
dump:
End Function

Sub CreateTableStyle(Optional styleName As String = TABLESTYLENAME, _
    Optional isRuled As Boolean = ISTABLERULED)
    'NOTE: This method can both create and reset (and existing) style
    ' to the settings below.
    
    Dim i As Long
    
    If Not StyleExists(styleName) Then
        ActiveWorkbook.TableStyles.Add styleName
        ActiveWorkbook.TableStyles(styleName).ShowAsAvailableTableStyle = True
    End If
    
    With ActiveWorkbook.TableStyles(styleName)
        With .TableStyleElements(xlHeaderRow)
            'Interior must preceed Font colors for Table Styles
            .Interior.Color = HEADERFILLCOLOR
            .Font.Color = ContrastText(.Interior.Color)
            .Font.FontStyle = "Bold"
            'NOTE: You cannot edit the header row height or font size in the
            ' table style definition. For this reason, those parameters are
            ' edited in the method called ApplyTableStyle().
        End With
        If isRuled Then
            .TableStyleElements(xlRowStripe2).Clear
            With .TableStyleElements(xlRowStripe2)
                For i = xlEdgeTop To xlEdgeBottom
                    With .Borders(i)
                        .Color = TABLERULECOLOR
                        .Weight = xlThin
                    End With
                Next
            End With
        End If
    End With
End Sub

Sub ApplyTableStyle(aTable As ListObject, Optional styleName As String = TABLESTYLENAME)
    
    aTable.TableStyle = ""
    CreateTableStyle styleName
    aTable.TableStyle = styleName
    
    With aTable.HeaderRowRange
        'Cannot format size and row height in table style creation
        .WrapText = False
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
        .Font.Size = 10
        
        .EntireRow.AutoFit
        .WrapText = True
        .RowHeight = .RowHeight + VERTICALPADDING
    End With
    
    aTable.HeaderRowRange.BorderAround Weight:=xlMedium, Color:=HEADERFILLCOLOR
    
End Sub

'==============================  DATE FUNCTIONS  ===============================

Sub WriteSaveDate(sht As Worksheet, Target As String)
    ' Called by AfterSave event in ThisWorkbook module
    With sht.Range(Target)
        .offset(0, -1).Value = "Saved:"
        .offset(0, -1).Font.Bold = True
        .Value = Now
        .NumberFormat = "m/d/yyyy h:mm AM/PM"
        .HorizontalAlignment = xlHAlignLeft
    End With
End Sub

Sub WriteDaysOpen(aTable As ListObject)
    
    If Not aTable.ListRows.Count = 0 Then
        
        Dim dateColumn As Range
        Dim statusColumn As Range
        Dim openColumn As Range
        
        Set dateColumn = aTable.ListColumns("Date").DataBodyRange
        Set statusColumn = aTable.ListColumns("Status").DataBodyRange
        Set daysOpenColumn = aTable.ListColumns("Days").DataBodyRange
        
        For i = 1 To aTable.ListRows.Count
            If LCase(statusColumn(i)) <> "resolved" Then _
                daysOpenColumn(i).Value = CLng(Now) - CLng(dateColumn(i).Value)
        Next
    End If
    
End Sub

'=============================  PRINTER SETTINGS  ==============================

Public Sub ApplyPageLayoutSettings(Optional orient As Long = xlLandscape, _
                                    Optional papSize As Long = xlPaperLetter, _
                                    Optional zoomPct As Long = 100)

    Dim sht As Worksheet
    Dim rowNo As Long
    
    Set sht = ActiveSheet
    rowNo = Split(ActiveSheet.Range(TABLEANCHORCELL).Address, "$")(2)

    If MsgBox("Continue with checklist print styles?", _
        vbCritical + vbYesNo, "Print Settings Dialog") = vbYes Then
        
        ResetExistingTable
        
        Application.PrintCommunication = False
        With sht.PageSetup
        
            .LeftFooter = "&A"
            .CenterFooter = "Page &P of &N"
            .RightFooter = "Printed: &D"
        
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.7)
            .BottomMargin = Application.InchesToPoints(0.7)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
        
            .CenterHorizontally = True
            .Orientation = orient
            .PaperSize = papSize
            .BlackAndWhite = False
            
            .Zoom = zoomPct
            
            .PrintTitleRows = "$" & rowNo & ":$" & rowNo
        End With
        Application.PrintCommunication = True   'Send all cached settings to printer
    End If
End Sub


'============================  GENERATIVE METHODS  =============================

Function IterateSheetName(baseName As String)

    Dim maxIndex As Long
    For Each sht In ActiveWorkbook.Sheets
        If Left(sht.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
    Next
    If maxIndex = 0 Then IterateSheetName = "" Else IterateSheetName = maxIndex

End Function

Sub CopyWorksheetChangeCode(sht As Worksheet)
    'Requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3"
    'and you must check YES to "Trust Access to VBA Object Model" in Macro Security Settings
    
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComp2 As VBIDE.VBComponent

    Set VBAEditor = Application.VBE
    Set VBProj = VBAEditor.ActiveVBProject
    Set VBComp = VBProj.VBComponents("Sheet2")  ' This should be the initial "POAM Log" sheet
                                                ' even if renamed or shifted by the user.
    Set VBComp2 = VBProj.VBComponents(sht.CodeName)
    
    codeString = VBComp.CodeModule.Lines(1, VBComp.CodeModule.CountOfLines)
    
    VBComp2.CodeModule.DeleteLines 1, VBComp2.CodeModule.CountOfLines
    VBComp2.CodeModule.InsertLines 1, codeString
        
End Sub
