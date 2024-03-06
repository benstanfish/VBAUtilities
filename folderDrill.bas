Attribute VB_Name = "folderDrill"
Private Const mod_name As String = "folderDrill"
Private Const module_author As String = "Ben Fisher"
Private Const module_version As String = "2.0"
Private Const module_update_date As Date = #3/6/2024#

' REFERENCE: Microsoft Scripting Runtime

Const firstDataRow = 2

Enum ColumnInfo
    [_First]
    idColumn = 1
    levelColumn = 2
    typeColumn = 3
    pathColumn = 4
    sizeColumn = 5
    linkColumn = 5
    extColumn = 6
    [_Last]
End Enum

Enum FileType
    [_First]
    Image = 0
    Drawing = 1
    Media = 2
    Data = 3
    Script = 4
    [_Last]
End Enum

Enum ColorSetting
    [_First]
    Stroke = 0
    Fill = 1
    [_Last]
End Enum

Public Sub RunDrill()
    Dim rootPath As String
    rootPath = PickAFolder()
    Application.ScreenUpdating = False
    If rootPath <> vbEmptyString Then
        Sheet1.Cells.Clear
        
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject

        Call WriteHeader
        
        Cells(firstDataRow, ColumnInfo.idColumn) = 1
        Cells(firstDataRow, ColumnInfo.levelColumn) = 0
        Cells(firstDataRow, ColumnInfo.typeColumn) = "dir"
        
        Cells(firstDataRow, ColumnInfo.pathColumn) = "../" & fso.GetBaseName(rootPath) & "/"
        Cells(firstDataRow, ColumnInfo.linkColumn).Hyperlinks.Add _
            Anchor:=Cells(firstDataRow, ColumnInfo.linkColumn), _
            Address:=rootPath, TextToDisplay:="Goto Folder"
        
        ActiveSheet.Name = Left(fso.GetFolder(rootPath).Name, 17)
        
        Set fso = Nothing
        
        Drill rootPath, 1
        
        With Cells(firstDataRow, ColumnInfo.idColumn).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = .ColumnWidth + 1
        End With
        
            With Cells(firstDataRow, ColumnInfo.levelColumn).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = 4
        End With
        
        With Cells(firstDataRow, ColumnInfo.typeColumn).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = .ColumnWidth + 1
        End With
        
        With Cells(firstDataRow, ColumnInfo.pathColumn).EntireColumn
            .Font.Name = "Consolas"
            .AutoFit
            If .ColumnWidth > 200 Then .ColumnWidth = 200
        End With
        
        With Cells(firstDataRow, ColumnInfo.sizeColumn).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignRight
            .AutoFit
        End With
        
        With Cells(firstDataRow, ColumnInfo.extColumn).EntireColumn
            .Font.Name = "Consolas"
            .AutoFit
        End With
        
        Call ColorFileTypes
        Call SummarizeData
        
    End If
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub

Private Function PickAFolder() As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        PickAFolder = fd.SelectedItems(1)
    End If
    Set fd = Nothing
End Function
Private Sub WriteHeader()
    Dim arr As Variant
    arr = Array("ID", "Lvl", "Type", "Path", "Size", "Ext")
    
    Dim headerRange As Range
    Set headerRange = ActiveSheet.Range("A1").Resize(1, UBound(arr) + 1)
    headerRange = arr
    headerRange.AutoFilter
    headerRange.Font.Bold = True
End Sub

Private Function CalcSize(aSize As Long) As String
    Select Case Log(aSize) / Log(10#)
        Case Is <= 3
            CalcSize = Round(aSize, 0) & " kB"
        Case Is <= 6
            CalcSize = Round(aSize / 1000#, 0) & " kB"
        Case Is <= 9
            CalcSize = Round(aSize / 1000000#, 1) & " MB"
        Case Is <= 12
            CalcSize = Round(aSize / 1000000000#, 1) & " GB"
        Case Else
            CalcSize = Round(aSize / 1000000000000#, 1) & " TB"
    End Select
End Function


Private Sub Drill(ByVal aPath As String, tierNumber As Long)

    Dim fso As FileSystemObject
    Dim aFolder As Folder, thisFolder As Folder
    Dim aFile As File, thisFile As File
    
    Set fso = New FileSystemObject
    Set aFolder = fso.GetFolder(aPath)
    
    For Each thisFile In aFolder.files
        NextRow
        ActiveCell = ActiveCell.Offset(-1, 0) + 1
        ActiveCell.Offset(0, ColumnInfo.levelColumn - 1) = tierNumber
        If fso.GetExtensionName(thisFile) = "lnk" Then
            ActiveCell.Offset(0, ColumnInfo.typeColumn - 1) = "lnk"
        Else
            ActiveCell.Offset(0, ColumnInfo.typeColumn - 1) = "f"
        End If
        ActiveCell.Offset(0, ColumnInfo.pathColumn - 1) = String(tierNumber * 4, " ") & thisFile.Name
        ActiveCell.Offset(0, ColumnInfo.linkColumn - 1) = CalcSize(thisFile.Size)
        ActiveCell.Offset(0, ColumnInfo.extColumn - 1) = fso.GetExtensionName(thisFile.Path)
    Next
    
    
    For Each thisFolder In aFolder.SubFolders
        NextRow
        ActiveCell = ActiveCell.Offset(-1, 0) + 1
        ActiveCell.Offset(0, ColumnInfo.levelColumn - 1) = tierNumber
        ActiveCell.Offset(0, ColumnInfo.typeColumn - 1) = "dir"
        ActiveCell.Offset(0, ColumnInfo.pathColumn - 1) = String(tierNumber * 4, " ") & "./" & fso.GetBaseName(thisFolder.Path) & "/"
        ActiveCell.Offset(0, ColumnInfo.linkColumn - 1).Hyperlinks.Add Anchor:=ActiveCell.Offset(0, ColumnInfo.linkColumn - 1), Address:=thisFolder.Path, TextToDisplay:="Goto Folder"
        Drill thisFolder.Path, tierNumber + 1
    Next

    Set fso = Nothing
    Set aFolder = Nothing
    Set aFile = Nothing

End Sub

Private Sub NextRow(Optional myColumn As String = "A")

    Range(myColumn & Rows.Count).End(xlUp).Offset(1, 0).Select

End Sub

Private Sub ColorFileTypes()

    Dim aDict As Dictionary
    Set aDict = New Dictionary
    
    aDict.Add "jpg", FileType.Image
    aDict.Add "jpeg", FileType.Image
    aDict.Add "bmp", FileType.Image
    aDict.Add "tif", FileType.Image
    aDict.Add "tiff", FileType.Image
    aDict.Add "png", FileType.Image
    aDict.Add "webp", FileType.Image
    aDict.Add "xcf", FileType.Image
    aDict.Add "svg", FileType.Image
    aDict.Add "dwg", FileType.Drawing
    aDict.Add "dxf", FileType.Drawing
    aDict.Add "rvt", FileType.Drawing
    aDict.Add "mp3", FileType.Media
    aDict.Add "mp4", FileType.Media
    aDict.Add "mpa", FileType.Media
    aDict.Add "avi", FileType.Media
    aDict.Add "wav", FileType.Media
    aDict.Add "mov", FileType.Media
    aDict.Add "xlsx", FileType.Data
    aDict.Add "xlsm", FileType.Data
    aDict.Add "csv", FileType.Data
    aDict.Add "doc", FileType.Data
    aDict.Add "xml", FileType.Data
    aDict.Add "json", FileType.Data
    aDict.Add "md", FileType.Data
    aDict.Add "txt", FileType.Data
    aDict.Add "log", FileType.Data
    aDict.Add "ini", FileType.Data
    aDict.Add "bas", FileType.Script
    aDict.Add "py", FileType.Script
    aDict.Add "cs", FileType.Script
    aDict.Add "ipynb", FileType.Script
    aDict.Add "bin", FileType.Script
    aDict.Add "exe", FileType.Script
    aDict.Add "msi", FileType.Script
    
    Dim arr(4, 1) As Long
    arr(FileType.Image, ColorSetting.Stroke) = webcolors.ORANGERED
    arr(FileType.Image, ColorSetting.Fill) = webcolors.MISTYROSE
    
    arr(FileType.Drawing, ColorSetting.Stroke) = webcolors.DODGERBLUE
    arr(FileType.Drawing, ColorSetting.Fill) = webcolors.ALICEBLUE
    
    arr(FileType.Media, ColorSetting.Stroke) = webcolors.KHAKI
    arr(FileType.Media, ColorSetting.Fill) = webcolors.LEMONCHIFFON
    
    arr(FileType.Data, ColorSetting.Stroke) = webcolors.SEAGREEN
    arr(FileType.Data, ColorSetting.Fill) = webcolors.MINTCREAM
    
    arr(FileType.Script, ColorSetting.Stroke) = webcolors.DARKMAGENTA
    arr(FileType.Script, ColorSetting.Fill) = webcolors.LAVENDERBLUSH
    
    
    Dim fileTypeRange As Range
    Set fileTypeRange = ActiveSheet.Range("E2", Cells(ActiveSheet.UsedRange.Rows.Count, 4))
    
    For Each aCell In fileTypeRange
        If aDict.Exists(LCase(aCell.Value)) Then
            With aCell
                .Font.Color = arr(aDict(aCell.Value), ColorSetting.Stroke)
                .Interior.Color = arr(aDict(aCell.Value), ColorSetting.Fill)
            End With
        End If
    Next

End Sub


Private Sub SummarizeData()

    Dim typeRange As Range
    Dim extRange As Range
    
    Dim typeColumnLetter As String: typeColumnLetter = "B2"
    Dim extColumnLetter As String: extColumnLetter = "E2"
       
    Set typeRange = Range(typeColumnLetter, _
        Cells(Range(typeColumnLetter).CurrentRegion.Rows.Count, _
            Range(typeColumnLetter).Column))

    Set extRange = Range(extColumnLetter, _
        Cells(Range(extColumnLetter).CurrentRegion.Rows.Count, _
            Range(extColumnLetter).Column))
    
    Dim anchorCell As Range
    Set anchorCell = Range("H1")
    With anchorCell
        .Value = "Summary"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    anchorCell.Offset(2, 0) = "Total subfolders:"
    anchorCell.Offset(3, 0) = "Total files:"
    anchorCell.Offset(4, 0) = "Total shortcuts:"
    
    anchorCell.Offset(2, 1).Formula2 = "=COUNTIF(" & Replace(typeRange.Address, "$", "") & "," & """dir""" & ") - 1"
    anchorCell.Offset(3, 1).Formula2 = "=COUNTIF(" & Replace(typeRange.Address, "$", "") & "," & """f""" & ")"
    anchorCell.Offset(4, 1).Formula2 = "=COUNTIF(" & Replace(typeRange.Address, "$", "") & "," & """lnk""" & ")"
    
    With anchorCell.Offset(6, 0)
        .Value = "File Types:"
        .Font.Bold = True
    End With
    
    anchorCell.Offset(7, 0).Formula2 = _
        "=UNIQUE(FILTER(" & Replace(extRange.Address, "$", "") & "," & Replace(extRange.Address, "$", "") & "<>""""))"
    
    Dim uniqueTypes As Long
    uniqueTypes = anchorCell.Offset(7, 0).CurrentRegion.Rows.Count - 2

    For i = 0 To uniqueTypes
        anchorCell.Offset(7 + i, 1).Formula2 = "=COUNTIF(" & Replace(extRange.Address, "$", "") & "," & anchorCell.Offset(7 + i, 0).Address & ")"
    Next

    anchorCell.EntireColumn.AutoFit

End Sub
