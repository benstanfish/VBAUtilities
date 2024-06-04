Attribute VB_Name = "dev_testing"

Private Sub Test_ParseToArray()
    Dim test_string As String
    test_string = "hello, here, there, by the way"
    arr = ParseToArray(test_string, "")
    For i = LBound(arr, 1) To UBound(arr, 1)
        Debug.Print "-" & arr(i) & "-"
    Next
End Sub

Private Sub Test_RecastArray()
    Dim test_string As String
    test_string = "123, 324, 6.3223234, 8233"
    arr = ParseToArray(test_string, ",")
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i), VarType(arr(i))
    Next
    arr2 = RecastArray(arr, vbLong)
    For i = LBound(arr2) To UBound(arr2)
        Debug.Print arr2(i), VarType(arr2(i))
    Next
End Sub

Private Sub Test_DoesTableStyleExist()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim styleName As String: styleName = "ZZZ"
    Debug.Print DoesTableStyleExist(styleName, wb)
End Sub

Private Sub Test_ListAllStyles()
    For Each aStyle In ActiveWorkbook.TableStyles
        Debug.Print aStyle.Name
    Next
End Sub

Private Sub Test_DeleteTableStyle()
    DeleteTableStyle "BaseStyle", ActiveWorkbook
End Sub

Private Sub Test_CreateTableStyle()
    CreateTableStyle "BaseStyle", ActiveWorkbook
End Sub

Private Sub Test_GetWSFromCell()
    Debug.Print Range("A1").Parent.Name
End Sub

Private Sub Test_CreateTableHeaderFromString()
    Dim fieldString As String, rng As Range
    fieldString = "ID, Status, Date, Discipline, Comment, Action, Response, Hello, World"
    Set rng = CreateTableHeaderFromString(Range("A1"), fieldString)
    Debug.Print rng.Address
End Sub

Private Sub Test_GetIntersectRange()
    Dim rng As Range, iRng As Range
    Set rng = Range("G1:K10")
    Set iRng = GetIntersectRange(rng)
    If Not iRng Is Nothing Then Debug.Print iRng.Address Else Debug.Print "No intersection"
End Sub

Private Sub Test_ApplyValidationListToColumn()
    ApplyValidationListToColumn aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
        fieldsString:="Yes,No,Maybe", isErrorSuppressed:=True
End Sub

'Private Sub Test_ApplyConditionalFormattingToTable()
'
'    Dim redStyle As New Dictionary
'    Dim yellowStyle As New Dictionary
'    Dim greenStyle As New Dictionary
'    Dim styleColl As New Collection
'
'    redStyle.Add Key:="interior_color", Item:=webcolors.MISTYROSE
'    redStyle.Add Key:="font_color", Item:=ContrastText(redStyle("interior_color"))
'    redStyle.Add Key:="bold_font", Item:=False
'
'    yellowStyle.Add Key:="interior_color", Item:=webcolors.LEMONCHIFFON
'    yellowStyle.Add Key:="font_color", Item:=ContrastText(yellowStyle("interior_color"))
'    yellowStyle.Add Key:="bold_font", Item:=False
'
'    greenStyle.Add Key:="interior_color", Item:=webcolors.MINTCREAM
'    greenStyle.Add Key:="font_color", Item:=ContrastText(greenStyle("interior_color"))
'    greenStyle.Add Key:="bold_font", Item:=False
'
'    styleColl.Add greenStyle
'    styleColl.Add redStyle
'    styleColl.Add yellowStyle
'
'    ApplyConditionalFormattingToTable aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
'        fieldsString:="Yes,No,Maybe", formatDictionaries:=styleColl, deletePrevious:=True, _
'        secondaryColumnFields:="Discipline,Hello"
'
'    Set styleColl = New Collection
'    Set redStyle = New Dictionary
'    Set yellowStyle = New Dictionary
'    Set greenStyle = New Dictionary
'
'    redStyle.Add Key:="interior_color", Item:=webcolors.TOMATO
'    redStyle.Add Key:="font_color", Item:=ContrastText(redStyle("interior_color"))
'    redStyle.Add Key:="bold_font", Item:=False
'
'    yellowStyle.Add Key:="interior_color", Item:=webcolors.YELLOW
'    yellowStyle.Add Key:="font_color", Item:=ContrastText(yellowStyle("interior_color"))
'    yellowStyle.Add Key:="bold_font", Item:=False
'
'    greenStyle.Add Key:="interior_color", Item:=webcolors.LIMEGREEN
'    greenStyle.Add Key:="font_color", Item:=ContrastText(greenStyle("interior_color"))
'    greenStyle.Add Key:="bold_font", Item:=False
'
'    styleColl.Add greenStyle
'    styleColl.Add redStyle
'    styleColl.Add yellowStyle
'
'    ApplyConditionalFormattingToTable aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
'        fieldsString:="Yes,No,Maybe", formatDictionaries:=styleColl, deletePrevious:=True
'
'    Set styleColl = Nothing
'    Set redStyle = Nothing
'    Set yellowStyle = Nothing
'    Set greenStyle = Nothing
'
'End Sub

Private Sub Test_CurrentInterior()
    Debug.Print "Color: " & Selection.Interior.Color
End Sub

Private Sub Test_SetToWhite()
    Selection.Interior.Color = xlNone
End Sub

Private Sub Test_FormatConfig_Init()

    Dim style As FormatConfig
    Set style = New FormatConfig
    Debug.Print style.fontName, style.fontSize, style.Bold
    
    style.InteriorColor = vbBlack
    Debug.Print style.FontColor = vbWhite
    
End Sub

Private Sub Test_CreateFormatConfig()
    Dim style As FormatConfig
    Set style = CreateFormatConfig(vbRed, "Arial Bold", 36, True)
    
    Debug.Print style.fontName, style.InteriorColor
End Sub

Private Sub Test_ApplyConditionalFormattingToTable_WithClass()
    
    Dim styles As New Collection
    
    Dim alert As FormatConfig, warn As FormatConfig, good As FormatConfig
    Set alert = CreateFormatConfig(webcolors.MISTYROSE)
    Set warn = CreateFormatConfig(webcolors.LEMONCHIFFON)
    Set good = CreateFormatConfig(webcolors.MINTCREAM)

    styles.Add alert
    styles.Add warn
    styles.Add good
    
    ActiveSheet.ListObjects(1).DataBodyRange.FormatConditions.Delete

    ApplyConditionalFormattingToTable aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
        fieldsString:="Yes,No,Maybe", stylesCollection:=styles, deletePrevious:=True, _
        secondaryColumnFields:="ID, World"

    alert.InteriorColor = webcolors.ORANGERED
    warn.InteriorColor = webcolors.YELLOW
    good.InteriorColor = webcolors.LIMEGREEN

    styles.Add alert
    styles.Add warn
    styles.Add good

    ApplyConditionalFormattingToTable aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
        fieldsString:="Yes,No,Maybe", stylesCollection:=styles, deletePrevious:=True

    Set alert = Nothing
    Set warn = Nothing
    Set good = Nothing
    Set styles = Nothing

End Sub

Private Sub Test_ApplyConditionalFormattingToTable_WithClass2()
    
    Dim styles As New Collection
    
    Dim alert As FormatConfig, warn As FormatConfig, good As FormatConfig
    Set alert = CreateFormatConfig(webcolors.ORANGERED)
    Set warn = CreateFormatConfig(webcolors.YELLOW)
    Set good = CreateFormatConfig(webcolors.LIMEGREEN)
    
    ActiveSheet.ListObjects(1).DataBodyRange.FormatConditions.Delete

    styles.Add alert
    styles.Add warn
    styles.Add good

    ApplyConditionalFormattingToTable aTable:=ActiveSheet.ListObjects(1), targetColumn:="Status", _
        fieldsString:="Yes,No,Maybe", stylesCollection:=styles, deletePrevious:=True

    Set alert = Nothing
    Set warn = Nothing
    Set good = Nothing
    Set styles = Nothing

End Sub

Private Sub Test_SetColumnWidths()
    SetColumnWidths ActiveSheet.ListObjects(1), "ID, Status, Hello", "20, 20, 8.38"
End Sub

Private Sub Test_AutoFitColumnWidths()
    AutoFitColumnWidths ActiveSheet.ListObjects(1)
End Sub
