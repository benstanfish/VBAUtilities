Attribute VB_Name = "tableDefaults"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "tableDefaults"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "2.1.3"
Public Const module_date As Date = #6/4/2024#
Public Const module_notes As String = _
    "This module provides functions for default or auto cell values"
Public Const module_license As String = "GNU General Public License, v3.0"

'***********************************************************************
'                          Referenced Libraries
'***********************************************************************
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (id As Any) As Long


'***********************************************************************
'                        Module Level Preferences
'***********************************************************************



'***********************************************************************
'                           Utility Functions
'***********************************************************************
Public Sub AutoNumberColumn(aTable As ListObject, _
                            Optional numberColumn As Variant = "ID", _
                            Optional overwriteExisting As Boolean = False, _
                            Optional useGUIDs As Boolean = False)
    Call utilities.MemorySaver
    
    If useGUIDs = False Then
        Dim maxIndex As Long
        If overwriteExisting = False Then
            maxIndex = WorksheetFunction.Max(aTable.ListColumns(numberColumn).DataBodyRange) + 1
            For Each aCell In aTable.ListColumns(numberColumn).DataBodyRange
                If aCell.Value = "" Then aCell.Value = maxIndex
                maxIndex = maxIndex + 1
            Next
        Else
            maxIndex = 1
            For Each aCell In aTable.ListColumns(numberColumn).DataBodyRange
                aCell.Value = maxIndex
                maxIndex = maxIndex + 1
            Next
        End If
    Else
        If overwriteExisting = False Then
            For Each aCell In aTable.ListColumns(numberColumn).DataBodyRange
                aCell.Value = CreateGUID
            Next
        Else
            For Each aCell In aTable.ListColumns(numberColumn).DataBodyRange
                aCell.Value = CreateGUID
            Next
        End If
    End If
    Call utilities.MemoryRestore
End Sub

Public Function CreateGUID() As String
    ' Adaptation by Mike Wolfe
    Const S_OK As Long = 0
    Dim id(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    If CoCreateGuid(id(0)) = S_OK Then
        For Cnt = 0 To 15
            CreateGUID = LCase(CreateGUID & IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt)))
        Next Cnt
        CreateGUID = LCase(Left$(CreateGUID, 8) & "-" & _
                     Mid$(CreateGUID, 9, 4) & "-" & _
                     Mid$(CreateGUID, 13, 4) & "-" & _
                     Mid$(CreateGUID, 17, 4) & "-" & _
                     Right$(CreateGUID, 12))
    End If
End Function

Public Function TimestampColumn(aTable As ListObject, _
                                stampColumn As String, _
                                Optional overwriteExisting As Boolean = False)

    Call utilities.MemorySaver
    With aTable.ListColumns(stampColumn).DataBodyRange
        .NumberFormat = Format("YYYY-MM-DD")
    End With
    If overwriteExisting = False Then
        For Each aCell In aTable.ListColumns(stampColumn).DataBodyRange
            If aCell.Value = "" Then aCell.Value = Now
        Next
    Else
        For Each aCell In aTable.ListColumns(stampColumn).DataBodyRange
            aCell.Value = Now
        Next
    End If
    Call utilities.MemoryRestore
End Function

Public Function ApplyDefaultValue(defaultValue As Variant, _
                                aTable As ListObject, _
                                targetColumn As String, _
                                Optional overwriteExisting As Boolean = False)

    Call utilities.MemorySaver
    If overwriteExisting = False Then
        For Each aCell In aTable.ListColumns(targetColumn).DataBodyRange
            If aCell.Value = "" Then aCell.Value = defaultValue
        Next
    Else
        For Each aCell In aTable.ListColumns(targetColumn).DataBodyRange
            aCell.Value = defaultValue
        Next
    End If
    Call utilities.MemoryRestore
End Function


