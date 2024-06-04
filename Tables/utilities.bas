Attribute VB_Name = "utilities"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "utilities"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.0.0"
Public Const module_date As Date = #6/4/2024#
Public Const module_notes As String = ""
Public Const module_license As String = "GNU General Public License, v3.0"

'***********************************************************************
'                          Referenced Libraries
'***********************************************************************
' NONE


'***********************************************************************
'                        Module Level Preferences
'***********************************************************************



'***********************************************************************
'                           Utility Functions
'***********************************************************************
Public Sub MemorySaver(Optional isEngaged As Boolean = True)
    Application.ScreenUpdating = isEngaged
    Application.EnableEvents = isEngaged
    Application.DisplayAlerts = isEngaged
End Sub

Public Sub MemoryRestore()
    ' Alias method
    MemorySaver False
End Sub

Public Function ParseToArray(parseString As String, _
                              Optional delimiter As String = ",") As Variant
    ' Parses a string to an array by splitting at the delimiter, which
    ' can be more than one character. It also trims off the delimiter if
    ' it's the first character in the string. Finally, it trims white spaces
    ' from the left and right of each entry in the array. Finally, the
    ' castAs parameter can be used to recast the values.
    
    If Left(parseString, Len(delimiter)) = delimiter Then
        parseString = Right(parseString, Len(parseString) - Len(delimiter))
    End If
    Dim i As Long, arr As Variant
    arr = Split(parseString, delimiter)
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim(arr(i))
    Next
    ParseToArray = arr
End Function

Public Function RecastArray(arr As Variant, _
                            Optional castTo As VbVarType = vbLong) As Variant
    ' Recast the values of an array to another value type.
    ' By default will recast values as long
    
    Dim Cnt As Long, i As Long
    Cnt = UBound(arr) - LBound(arr) ' 0-Based
    ReDim arr2(0 To Cnt)
    Select Case castTo
        Case Is = vbInteger
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CInt(arr(i))
            Next
        Case Is = vbLong
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CLng(arr(i))
            Next
        Case Is = vbSingle
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CSng(arr(i))
            Next
        Case Is = vbDouble
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CDbl(arr(i))
            Next
        Case Is = vbCurrency
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CCur(arr(i))
            Next
        Case Is = vbDate
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CDate(arr(i))
            Next
        Case Is = vbBoolean
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CBool(arr(i))
            Next
        Case Else
            For i = LBound(arr2) To UBound(arr2)
                arr2(i) = CLng(arr(i))
            Next
    End Select
    RecastArray = arr2
End Function

Public Function RemoveIllegalChars(aString As String) As String
    Dim illegalChars As Variant, aChar As Variant
    illegalChars = Array("<", ">", ":", Chr(34), "\", "/", "|", "?", "*", ";")
    For Each aChar In illegalChars
        aString = Replace(aString, aChar, "")
    Next
    RemoveIllegalChars = aString
End Function

Public Function IterateTableName(baseName As String, wb As Workbook) As String
    Dim maxIndex As Long, aTable As ListObject, ws As Worksheet
    For Each ws In wb.Sheets
        For Each aTable In ws.ListObjects
            If Left(aTable.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
        Next
    Next
    If maxIndex = 0 Then
        IterateTableName = baseName
    Else
        IterateTableName = baseName & maxIndex
    End If
End Function

Public Function IterateSheetName(baseName As String, wb As Workbook) As String
    Dim maxIndex As Long, ws As Worksheet
    For Each ws In wb.Sheets
        If Left(ws.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
    Next
    If maxIndex = 0 Then
        IterateSheetName = baseName
    Else
        IterateSheetName = baseName & maxIndex
    End If
End Function

Function RenameSheet(proposedName As String, ws As Worksheet) As String
    ' Renames a Worksheet (Spreadsheet tab)
    ' It limits the name to the maximum permitted character count (31 - 4 = 27) and removes
    ' illegal characters from the name
    
    Dim newName As String
    newName = RemoveIllegalChars(proposedName)
    newName = Left(newName, 27)

    On Error GoTo dump
    ws.Name = IterateSheetName(newName, ws.Parent)
dump:
    RenameSheet = ws.Name
End Function

Private Sub PrintSectionTitleToDebug()
    ' This sub takes the user's input and writes to the debug window
    ' a pretty Section Header with the title centered.

    Dim totalChars As Long, paddingCount As Long, reptChar As String
    totalChars = 72    ' Preferred code width limit (' is the first character)
    reptChar = "*"
    paddingCount = 2
    
    Dim delimiter As String
    delimiter = String(totalChars - 1, reptChar)
    delimiter = "'" & delimiter
    
    Dim sectionTitle As String
    sectionTitle = String(paddingCount, " ") & _
                Application.InputBox("Write a string", "Input Title") & _
                String(paddingCount, " ")

    Dim paddingString As String
    paddingString = String((totalChars - Len(sectionTitle) - 1) / 2, " ")
        
    Debug.Print delimiter
    Debug.Print "'" & paddingString & sectionTitle
    Debug.Print delimiter
    
End Sub


'=======================================================================
'Function CreateWorkbook(save_path As String, _
'    Optional workbook_name As String = "Bidder RFI Summary Report", _
'    Optional include_timestamp As Boolean = TIMESTAMPFILE) As Workbook
'    ' Return a  new Workbook object with the provided name
'    ' and appends with a timestamp as noted.
'    Dim combined_workbook As Workbook
'    Dim file_name As String
'    Set combined_workbook = Workbooks.Add
'    Application.DisplayAlerts = False
'    With combined_workbook
'        .Title = workbook_name
'        If include_timestamp Then
'            file_name = save_path & "\" & workbook_name & " " _
'                & Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
'        Else
'            file_name = save_path & "\" & workbook_name & ".xlsx"
'        End If
'        .SaveAs Filename:=file_name, FileFormat:=xlOpenXMLWorkbook
'    End With
'    Application.DisplayAlerts = True
'    Set CreateWorkbook = combined_workbook
'End Function


