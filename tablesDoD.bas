Attribute VB_Name = "tablesDOD"
Public Const mod_name As String = "tablesDoD"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.3"
Public Const module_date As Date = #4/2/2024#

Public Const Army = "Army"
Public Const USN = "Navy"
Public Const AF = "AF"
Public Const USMC = "USMC"

Public Const TARGET_INFO = "A2"
Public Const TARGET_PDT = "A14"
Public Const TARGET_COMMENTS = "A34"
Public Const TARGET_SCHED = "D2"

Public Enum TableInfo
    [_First]
    RegionTitle = -1
    HeaderRow = 1
    IDColumn = 1
    ProjectNameLine = 2
    P2Line = 3
    PALine = 4
    CWELine = 5
    JESLine = 6
    FundingLine = 7
    ClientLine = 8
    ContractLine = 9
    WatermarkLine = 10
    [_Last]
End Enum

Private Sub CreateDoDTableStyles()
    Dim ArmyArr As Variant, AFArr As Variant, USNArr As Variant, USMCArr As Variant
    ArmyArr = Array(3357492, 6648690, 10728899, 3331582)
    AFArr = Array(4662044, 15263977, 13816532, 16677632)
    USNArr = Array(5192448, 15263977, 13683910, 1028328)
    USMCArr = Array(2185797, 7646395, 7377836, 3150532)
    
    CreateTableStyle Army, ArmyArr(0), ArmyArr(1), ArmyArr(2)
    CreateTableStyle AF, AFArr(0), AFArr(1), AFArr(2)
    CreateTableStyle USN, USNArr(0), USNArr(1), USNArr(2)
    CreateTableStyle USMC, USMCArr(0), USMCArr(1), USMCArr(2)
End Sub

Public Function DoesTableStyleExist(styleName As String) As Boolean
    Dim dummyBool As Boolean
    On Error GoTo dump
        If ActiveWorkbook.TableStyles(styleName) = styleName Then dummyBool = True
        ' Simply need to evaluate an expression to force potential error
    On Error GoTo 0
    DoesTableStyleExist = True
    Exit Function
dump:
End Function

Private Sub DeleteTableStyle(styleName As String)
    If DoesTableStyleExist(styleName) Then ActiveWorkbook.TableStyles(styleName).Delete
End Sub

Private Sub DeleteAllTableStyles()
    DeleteTableStyle Army
    DeleteTableStyle AF
    DeleteTableStyle USN
    DeleteTableStyle USMC
End Sub

Private Sub CreateTableStyle(ByVal styleName As String, ByVal headerColor As Long, _
                            ByVal rowStripe1Color As Long, ByVal rowStripe2Color As Long)
    If DoesTableStyleExist(styleName) = False Then
        ActiveWorkbook.TableStyles.Add styleName
        ActiveWorkbook.TableStyles(styleName).ShowAsAvailableTableStyle = True
        With ActiveWorkbook.TableStyles(styleName)
            'Note: you can't set the default font name or size
            With .TableStyleElements(xlHeaderRow)
                .Interior.Color = headerColor
                .Borders(xlEdgeBottom).Color = headerColor
                .Font.Color = ContrastText(.Interior.Color)
            End With
            With .TableStyleElements(xlRowStripe1)
                .Interior.Color = rowStripe1Color
                .Font.Color = ContrastText(.Interior.Color)
            End With
            With .TableStyleElements(xlRowStripe2)
                .Interior.Color = rowStripe2Color
                .Font.Color = ContrastText(.Interior.Color)
            End With
        End With
    End If
End Sub

Public Sub HighlightRow()
    Application.ScreenUpdating = False
    
    Dim arr As Variant
    arr = Array(3331582, 16677632, 1028328, 3150532)
    
    Dim aTable As ListObject
    Set aTable = ActiveCell.ListObject
    
    Dim rowIndex As Long
    rowIndex = ActiveCell.Row - aTable.HeaderRowRange.Row
    
    Dim currentStyle As String
    currentStyle = aTable.TableStyle.Name
    
    With aTable.DataBodyRange.Rows(rowIndex)
        Select Case currentStyle
            Case Is = Army
                .Interior.Color = arr(0)
            Case Is = AF
                .Interior.Color = arr(1)
            Case Is = USN
                .Interior.Color = arr(2)
            Case Is = USMC
                .Interior.Color = arr(3)
            Case Else
        End Select
        .Font.Color = ContrastText(.Interior.Color)
        .Font.Bold = True
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub ApplyTableSecondaryFormats(aTable As ListObject)
    
    With aTable
        .ShowAutoFilterDropDown = False
        .Range.Font.Name = "Arial"
        .Range.Font.Size = 10.5
    End With
    
    With aTable.HeaderRowRange
        .Font.Bold = True
        .WrapText = True
    End With
    
    For i = 3 To aTable.HeaderRowRange.Columns.Count
        aTable.HeaderRowRange(, i).HorizontalAlignment = xlHAlignRight
        With aTable.DataBodyRange.EntireColumn(i)
            .HorizontalAlignment = xlHAlignRight
            .NumberFormat = "dd MMM yy"
        End With
    Next
    
    Dim colWidths As Variant
    colWidths = Array(4, 20, 10, 10, 10)
    For i = LBound(colWidths) To UBound(colWidths)
        aTable.HeaderRowRange.Columns(i + 1).ColumnWidth = colWidths(i)
    Next
    
    For Each aRow In aTable.DataBodyRange.Rows
        aRow.EntireRow.AutoFit
        aRow.RowHeight = aRow.RowHeight + 2
    Next
    aTable.DataBodyRange.VerticalAlignment = xlVAlignCenter
    
    Range(TARGET_SCHED).EntireRow.AutoFit
    aTable.Range.Columns(TableInfo.IDColumn).HorizontalAlignment = xlHAlignCenter
    
End Sub

Public Sub UnhighlightRow()
    Application.ScreenUpdating = False
    
    Dim aRng As Range
    Set aRng = Selection.ListObject.Range
    
    Dim currentStyle As String
    currentStyle = Selection.ListObject.TableStyle.Name
    
    With Selection.ListObject
        .Range.ClearFormats
        .ShowAutoFilterDropDown = False
        .TableStyle = currentStyle
    End With
    
    ApplyTableSecondaryFormats aRng.ListObject
    Application.ScreenUpdating = True
End Sub

Function AutoincrementTable(baseName As String) As String
    Dim maxIndex As Long
    Dim sht As Worksheet
    Dim lstObj As ListObject
    For Each sht In ActiveWorkbook.Sheets
        For Each lstObj In sht.ListObjects
            If Left(lstObj.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
        Next
    Next
    If maxIndex = 0 Then AutoincrementTable = baseName Else AutoincrementTable = baseName & maxIndex
End Function

Public Sub CreateProjectInfoTable(sht As Worksheet)

    Dim target As Range
    Set target = sht.Range(TARGET_INFO)

    Dim arr As Variant
    arr = Array("Parameter", "Project Name", "P2", "PA", "CWE/ECC", "JES?", "Funding", "Client", "Contract", "Watermark")
    
    target.Resize(UBound(arr) + 1) = WorksheetFunction.Transpose(arr)
    target.Offset(0, 1).Value = "Value"
    
    Dim pITable As ListObject
    Set pITable = ActiveSheet.ListObjects.Add(xlSrcRange, target.CurrentRegion, , xlYes)
    pITable.Name = AutoincrementTable("ProjectInfo")
    With ActiveSheet.ListObjects(pITable.Name)
        .TableStyle = "TableStyleLight9"
        .ShowAutoFilterDropDown = False
        .DataBodyRange.HorizontalAlignment = xlHAlignLeft
    End With

    target.EntireColumn.AutoFit
    target.Offset(0, 1).ColumnWidth = 50
    
    With target.Offset(TableInfo.RegionTitle, 0)
        .Value = "Project Info"
        .Font.Bold = True
    End With
    
    With target.Offset(TableInfo.CWELine - 1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="CWE " & ChrW(8804) & " ECC, CWE " & ChrW(8805) & " ECC, CWE ? ECC"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With target.Offset(TableInfo.JESLine - 1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Yes, No, Unknown"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    
    With target.Offset(TableInfo.FundingLine - 1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="MILCON, SRM, O&M, Host Nation, Other"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    
    With target.Offset(TableInfo.ClientLine - 1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Army, Air Force, Navy, Marines, DPW, DLA, DoDEA"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    
    With target.Offset(TableInfo.ContractLine - 1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="DBB, DB"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    
    
    
End Sub

Public Sub CreatePDTTable(sht As Worksheet)

    Dim target As Range
    Set target = sht.Range(TARGET_PDT)

    Dim arr As Variant
    arr = Array("Role", "TL", "PM", "DM", "A/E", "Civ", "Str", "Arch", _
                "Mech", "Elec", "FPE", "Cyber", "Env", "Sust", "Cost", _
                "VE", "TS", "MCX")
    
    target.Resize(UBound(arr) + 1) = WorksheetFunction.Transpose(arr)
    target.Offset(0, 1).Value = "Person"

    Dim pDTTable As ListObject
    Set pDTTable = ActiveSheet.ListObjects.Add(xlSrcRange, target.CurrentRegion, , xlYes)
    pDTTable.Name = AutoincrementTable("PDT")
    With ActiveSheet.ListObjects(pDTTable.Name)
        .TableStyle = "TableStyleDark11"
        .ShowAutoFilterDropDown = False
    End With
    
     With target.Offset(TableInfo.RegionTitle, 0)
        .Value = "Project Team"
        .Font.Bold = True
    End With

End Sub

Public Sub CreateButtons(sht As Worksheet)
    
    For Each aShape In sht.Shapes
        aShape.Delete
    Next
    
    Dim schedButton As Shape
    Set schedButton = sht.Shapes.AddShape(msoShapeRoundedRectangle, 800, 30, 150, 30)
    With schedButton.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = "Overwrite Schedule"
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With schedButton
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Shadow.Type = msoShadow21
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 10
            .OffsetX = 2.25
            .OffsetY = 2.25
            .RotateWithShape = msoFalse
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.6
            .Size = 100
        End With
        .Placement = xlFreeFloating
        .OnAction = "OverwriteSchedule"
    End With
    
    Dim slideButton As Shape
    Set slideButton = sht.Shapes.AddShape(msoShapeRoundedRectangle, 800, 70, 150, 30)
    With slideButton.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = "Generate Slide"
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With slideButton
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Shadow.Type = msoShadow21
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 10
            .OffsetX = 2.25
            .OffsetY = 2.25
            .RotateWithShape = msoFalse
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.6
            .Size = 100
        End With
        .Placement = xlFreeFloating
        .OnAction = "GenerateThisSlide"
    End With
    
    Dim slidesButton As Shape
    Set slidesButton = sht.Shapes.AddShape(msoShapeRoundedRectangle, 1000, 70, 150, 30)
    With slidesButton.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = "Generate ALL Slides"
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With slidesButton
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Shadow.Type = msoShadow21
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 10
            .OffsetX = 2.25
            .OffsetY = 2.25
            .RotateWithShape = msoFalse
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.6
            .Size = 100
        End With
        .Placement = xlFreeFloating
        .OnAction = "GenerateAllSlides"
    End With
    
    Dim newSheetButton As Shape
    Set newSheetButton = sht.Shapes.AddShape(msoShapeRoundedRectangle, 800, 110, 150, 30)
    With newSheetButton.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = "Create New Sheet"
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With newSheetButton
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Shadow.Type = msoShadow21
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 10
            .OffsetX = 2.25
            .OffsetY = 2.25
            .RotateWithShape = msoFalse
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.6
            .Size = 100
        End With
        .Placement = xlFreeFloating
        .OnAction = "CreateNewSheet"
    End With
       
End Sub

Public Sub RefreshButtons()
    CreateButtons ActiveSheet
End Sub

Public Sub CreateCommentsTable(sht As Worksheet)

    Dim target As Range
    Set target = sht.Range(TARGET_COMMENTS)

    Dim arr As Variant
    arr = Array("Show", "Comment")
    
    target.Resize(, UBound(arr) + 1) = arr
    target.Offset(1, 0).Value = "X"
    
    Dim commentTable As ListObject
    Set commentTable = ActiveSheet.ListObjects.Add(xlSrcRange, target.CurrentRegion, , xlYes)
    commentTable.Name = AutoincrementTable("Comments")
    With ActiveSheet.ListObjects(commentTable.Name)
        .TableStyle = "TableStyleLight9"
        .ShowAutoFilterDropDown = False
        With .DataBodyRange
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .Columns(2).WrapText = True
        End With
    End With
    
    With target.Offset(TableInfo.RegionTitle, 0)
        .Value = "Critical Issues/Updates Table"
        .Font.Bold = True
    End With


End Sub

Public Sub CreateDBBSchedule(sht As Worksheet)

    Dim target As Range
    Set target = sht.Range(TARGET_SCHED)

    Dim arr As Variant
    arr = Array("ID", "Task", "Mtg or Submittal", "Review Start", "End")
    target.Resize(, UBound(arr) + 1) = arr
    
    arr = Array("Kickoff", "Charette", "  " & ChrW(9500) & " Draft PDCR", "  " & ChrW(9500) & " Final", "  " & ChrW(9492) & " Backcheck", _
        "Concept", "  " & ChrW(9492) & " OBR", "Intermediate", "  " & ChrW(9492) & " OBR", "Final", "  " & ChrW(9492) & " OBR", "Backcheck", _
        "BCOES Cert", "RTA", "Advertise")

    target.Offset(1, 1).Resize(UBound(arr) + 1) = WorksheetFunction.Transpose(arr)
    
    Dim schedTable As ListObject
    Dim projType As String
    Set schedTable = ActiveSheet.ListObjects.Add(xlSrcRange, target.CurrentRegion, , xlYes)
    schedTable.Name = AutoincrementTable("Schedule")
    With ActiveSheet.ListObjects(schedTable.Name)
        projType = ActiveSheet.ListObjects(1).DataBodyRange.Cells(7, 2)
        Select Case projType
            Case Is = "Air Force"
                .TableStyle = AF
            Case Is = "Navy"
                .TableStyle = USN
            Case Is = "Marines"
                .TableStyle = USMC
            Case Else
                .TableStyle = Army
        End Select
        .ShowAutoFilterDropDown = False
    End With
    ApplyTableSecondaryFormats schedTable
    
    With target.Offset(TableInfo.RegionTitle, 0)
        .Value = "Project Scheduled"
        .Font.Bold = True
    End With
    
    For i = 1 To schedTable.DataBodyRange.Rows.Count
        schedTable.DataBodyRange.Columns(1).Rows(i) = i
    Next
    
End Sub

Public Sub CreateDBSchedule(sht As Worksheet)

    Dim target As Range
    Set target = sht.Range(TARGET_SCHED)

    Dim arr As Variant
    arr = Array("ID", "Task", "Mtg or Submittal", "Review Start", "End")
    target.Resize(, UBound(arr) + 1) = arr
    
    arr = Array("Kickoff", "Scope Valid.", "  " & ChrW(9492) & " Site Visit", _
        "Draft RFP", "  " & ChrW(9492) & " OBR", "Final RFP", "  " & ChrW(9492) & " OBR", "Backcheck", _
        "RTA", "Advert")

    target.Offset(1, 1).Resize(UBound(arr) + 1) = WorksheetFunction.Transpose(arr)
    
    Dim schedTable As ListObject
    Dim projType As String
    Set schedTable = ActiveSheet.ListObjects.Add(xlSrcRange, target.CurrentRegion, , xlYes)
    schedTable.Name = AutoincrementTable("Schedule")
    With ActiveSheet.ListObjects(schedTable.Name)
        projType = ActiveSheet.ListObjects(1).DataBodyRange.Cells(TableInfo.ClientLine, 2)
        Select Case projType
            Case Is = "Air Force"
                .TableStyle = AF
            Case Is = "Navy"
                .TableStyle = USN
            Case Is = "Marines"
                .TableStyle = USMC
            Case Else
                .TableStyle = Army
        End Select
        .ShowAutoFilterDropDown = False
    End With
    ApplyTableSecondaryFormats schedTable
    
    With target.Offset(TableInfo.RegionTitle, 0)
        .Value = "Project Scheduled"
        .Font.Bold = True
    End With
    
    For i = 1 To schedTable.DataBodyRange.Rows.Count
        schedTable.DataBodyRange.Columns(TableInfo.IDColumn).Rows(i) = i
    Next

End Sub

Public Sub CreateProjectTablesExceptSchedule(sht As Worksheet)
    sht.Cells.Clear
    CreateProjectInfoTable sht
    CreatePDTTable sht
    CreateCommentsTable sht
End Sub

Public Sub OverwriteSchedule()
    CreateScheduleTable ActiveSheet
End Sub

Public Sub CreateScheduleTable(sht As Worksheet)
    If vbYes = MsgBox("Do you want to overwrite existing schedule (if exists)?", vbYesNo + vbCritical, "Overwrite Warning") Then
        Range(TARGET_SCHED).Offset(TableInfo.RegionTitle, 0).Clear
        On Error Resume Next
        sht.ListObjects(TableType.sched).Delete
        Dim contract As String
        contract = sht.ListObjects(TableType.info).DataBodyRange.Cells(TableInfo.ContractLine, 2)
        Select Case LCase(contract)
            Case Is = "db"
                CreateDBSchedule sht
            Case Else
                CreateDBBSchedule sht
        End Select
    End If
End Sub

Public Sub CreateNewSheet()
    Dim aSht As Worksheet
    Set aSht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    With aSht
        CreateProjectTablesExceptSchedule aSht
        CreateButtons aSht
        ActiveWindow.DisplayGridlines = False
    End With
End Sub
