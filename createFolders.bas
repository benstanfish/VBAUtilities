Attribute VB_Name = "createFolders"

Public Sub CreateProjectFolders()

    Dim rootPath As String: rootPath = "C:\Users\j2ee9bsf\Documents\00 Projects\"
    Dim fso As FileSystemObject
    Dim aTable As ListObject
    
    Set aTable = ActiveWorkbook.Sheets("Projects").ListObjects(1)
    Set fso = New FileSystemObject
    
    Dim cRow As Long, cProject As String
    For Each aProject In aTable.ListColumns("Name").DataBodyRange
        cProject = ScrubIllegalChars(aProject.Value)
        cRow = aProject.Row - aTable.HeaderRowRange.Row
        If aTable.ListColumns("Create Folder").DataBodyRange.Rows(cRow) = True Then
            If Not fso.FolderExists(fso.BuildPath(rootPath, cProject)) Then
                fso.CreateFolder (fso.BuildPath(rootPath, cProject))
            End If
        End If
    Next

    Set aTable = Nothing
    Set fso = Nothing

End Sub

Public Function ScrubIllegalChars(aString As String) As String
    Dim illegalChars As Variant
    illegalChars = Array("<", ">", ":", Chr(34), "\", "/", "|", "?", "*")
    For Each aChar In illegalChars
        aString = Replace(aString, aChar, "")
    Next
    ScrubIllegalChars = aString
End Function
