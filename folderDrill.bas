Attribute VB_Name = "folderDrill"
Private Const mod_name As String = "folderDrill"
Private Const module_author As String = "Ben Fisher"
Private Const module_version As String = "1.0.0"
Private Const module_update_date As Date = #3/5/2024#

Public Sub RunDrill()
    Dim rootPath As String
    rootPath = PickAFolder()
    Application.ScreenUpdating = False
    If rootPath <> vbEmptyString Then
        Sheet1.Cells.Clear
        
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject

        Range("A1") = 1

        Range("B1") = "../" & fso.GetBaseName(rootPath) & "/"
        Range("B1").Offset(0, 1).Hyperlinks.Add Anchor:=Range("A1").Offset(0, 1), Address:=rootPath, TextToDisplay:="Goto Folder"
        
        Set fso = Nothing
        
        Drill rootPath, 1
        
        With Range("A1").EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignLeft
            .AutoFit
            .ColumnWidth = .ColumnWidth + 1
        End With
        
        With Range("B1").EntireColumn
            .Font.Name = "Consolas"
            .AutoFit
            If .ColumnWidth > 200 Then .ColumnWidth = 200
        End With
        
        With Range("C1").EntireColumn
            .Font.Name = "Consolas"
            .HorizontalAlignment = xlHAlignRight
            .AutoFit
        End With
        
    End If
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
        ActiveCell.Offset(0, 1) = String(tierNumber * 4, " ") & thisFile.Name
        ActiveCell.Offset(0, 2) = CalcSize(thisFile.Size)
    Next
    
    
    For Each thisFolder In aFolder.SubFolders
        NextRow
        ActiveCell = ActiveCell.Offset(-1, 0) + 1
        ActiveCell.Offset(0, 1) = String(tierNumber * 4, " ") & "./" & fso.GetBaseName(thisFolder.Path) & "/"
        ActiveCell.Offset(0, 2).Hyperlinks.Add Anchor:=ActiveCell.Offset(0, 2), Address:=thisFolder.Path, TextToDisplay:="Goto Folder"
        
        Drill thisFolder.Path, tierNumber + 1
    Next

    Set fso = Nothing
    Set aFolder = Nothing
    Set aFile = Nothing

End Sub



Private Sub NextRow(Optional myColumn As String = "A")

    Range(myColumn & Rows.Count).End(xlUp).Offset(1, 0).Select

End Sub
