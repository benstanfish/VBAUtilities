Attribute VB_Name = "folderDrill"
Private Const mod_name As String = "folderDrill"
Private Const module_author As String = "Ben Fisher"
Private Const module_version As String = "1.3"
Private Const module_update_date As Date = #3/6/2024#

' REFERENCE: Microsoft Scripting Runtime

Enum ColumnInfo
    [_First]
    ID = 1
    dirFile = 2
    Tree = 3
    fileSize = 4
    dirLink = 4
    extType = 5
    [_Last]
End Enum

Enum FileType
    [_First]
    Image = 0
    Drawing = 1
    Media = 2
    Data = 3
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

        Cells(1, ColumnInfo.ID) = 1
    
        Cells(1, ColumnInfo.dirFile) = "dir"
        
        Cells(1, ColumnInfo.Tree) = "../" & fso.GetBaseName(rootPath) & "/"
        Cells(1, ColumnInfo.dirLink).Hyperlinks.Add Anchor:=Cells(1, ColumnInfo.dirLink), Address:=rootPath, TextToDisplay:="Goto Folder"
        
        ActiveSheet.Name = Left(fso.GetFolder(rootPath).Name, 17)
        
        Set fso = Nothing
        
        Drill rootPath, 1
        
        With Cells(1, ColumnInfo.ID).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = .ColumnWidth + 1
        End With
        
        With Cells(1, ColumnInfo.dirFile).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = .ColumnWidth + 1
        End With
        
        With Cells(1, ColumnInfo.Tree).EntireColumn
            .Font.Name = "Consolas"
            .AutoFit
            If .ColumnWidth > 200 Then .ColumnWidth = 200
        End With
        
        With Cells(1, ColumnInfo.fileSize).EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignRight
            .AutoFit
        End With
        
        With Cells(1, ColumnInfo.extType).EntireColumn
            .Font.Name = "Consolas"
            .AutoFit
        End With
        
        Call ColorFileTypes
        
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
        ActiveCell.Offset(0, ColumnInfo.dirFile - 1) = "f"
        ActiveCell.Offset(0, ColumnInfo.Tree - 1) = String(tierNumber * 4, " ") & thisFile.Name
        ActiveCell.Offset(0, ColumnInfo.dirLink - 1) = CalcSize(thisFile.Size)
        ActiveCell.Offset(0, ColumnInfo.extType - 1) = fso.GetExtensionName(thisFile.Path)
    Next
    
    
    For Each thisFolder In aFolder.SubFolders
        NextRow
        ActiveCell = ActiveCell.Offset(-1, 0) + 1
        ActiveCell.Offset(0, ColumnInfo.dirFile - 1) = "dir"
        ActiveCell.Offset(0, ColumnInfo.Tree - 1) = String(tierNumber * 4, " ") & "./" & fso.GetBaseName(thisFolder.Path) & "/"
        ActiveCell.Offset(0, ColumnInfo.dirLink - 1).Hyperlinks.Add Anchor:=ActiveCell.Offset(0, ColumnInfo.dirLink - 1), Address:=thisFolder.Path, TextToDisplay:="Goto Folder"
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
    
    Dim arr(3, 1) As Long
    arr(FileType.Image, ColorSetting.Stroke) = webcolors.ORANGERED
    arr(FileType.Image, ColorSetting.Fill) = webcolors.MISTYROSE
    
    arr(FileType.Drawing, ColorSetting.Stroke) = webcolors.DODGERBLUE
    arr(FileType.Drawing, ColorSetting.Fill) = webcolors.ALICEBLUE
    
    arr(FileType.Media, ColorSetting.Stroke) = webcolors.KHAKI
    arr(FileType.Media, ColorSetting.Fill) = webcolors.LEMONCHIFFON
    
    arr(FileType.Data, ColorSetting.Stroke) = webcolors.SEAGREEN
    arr(FileType.Data, ColorSetting.Fill) = webcolors.MINTCREAM
    
    
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

