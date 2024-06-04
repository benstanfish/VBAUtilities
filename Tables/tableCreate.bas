Attribute VB_Name = "tableCreate"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "tableCreate"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "2.1.2"
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
Public Const DEFAULT_TABLESTYLE_NAME As String = "BaseStyle"
Public Const DEFAULT_TABLE_NAME As String = "MyTable"


'***********************************************************************
'                           Utility Functions
'***********************************************************************
Public Function ItersectsWithExistingTable(proposedTableRange As Range) As Boolean
    ' This function tests to see if the proposed table range would
    ' intersect with an existing table, which is disallowed by Excel.
    ' If you want to know the intersection range, use GetIntersectRange()
    Dim ws As Worksheet, aTable As ListObject, temp As Boolean
    Set ws = proposedTableRange.Parent
    For Each aTable In ws.ListObjects
        If Not Intersect(proposedTableRange, aTable.Range) Is Nothing Then
            temp = True
            Exit For
        End If
    Next
    Set aTable = Nothing
    Set ws = Nothing
    ItersectsWithExistingTable = temp
End Function

Public Function GetIntersectRange(proposedTableRange As Range) As Range
    ' This function returns the intersection range of a proposed range
    ' with the overlap of an existing table. Note that if there is no
    ' intersection this function returns a Nothing object.
    Dim ws As Worksheet, aTable As ListObject, temp As Range
    Set ws = proposedTableRange.Parent
    For Each aTable In ws.ListObjects
        If Not Intersect(proposedTableRange, aTable.Range) Is Nothing Then
            Set temp = Intersect(proposedTableRange, aTable.Range)
            Exit For
        End If
    Next
    Set aTable = Nothing
    Set ws = Nothing
    Set GetIntersectRange = temp
End Function


'***********************************************************************
'                      Create and Delete Functions
'***********************************************************************
Public Function CreateTableHeaderFromString(anchorCell As Range, _
                                       fieldsString As String, _
                                       Optional styleName As String = DEFAULT_TABLESTYLE_NAME, _
                                       Optional tableName As String = DEFAULT_TABLE_NAME) As Range
    ' Function that accepts a target cell, and parsable string of column
    ' field names, then pastes and turns the range into a ListObject header.
    
    Dim arr As Variant, fieldCount As Long, headerRange As Range
    Dim ws As Worksheet, wb As Workbook
    Set ws = anchorCell.Parent
    Set wb = ws.Parent
    
    arr = utilities.ParseToArray(fieldsString)
    fieldCount = UBound(arr) - LBound(arr) + 1          'Add 1 to get into Base 1
    Set headerRange = anchorCell.Resize(1, fieldCount)
    headerRange.Value = arr
    
    CreateTableStyle styleName, wb
    
    If ItersectsWithExistingTable(headerRange) = False Then
        ws.ListObjects.Add SourceType:=xlSrcRange, _
                           Source:=headerRange, _
                           xlListObjectHasHeaders:=xlYes, _
                           TableStyleName:=styleName
    End If
    
    Set CreateTableHeaderFromString = headerRange
    Set ws = Nothing
    Set headerRange = Nothing
    Set arr = Nothing
End Function

Public Sub DeleteTable(aTable As ListObject)
    aTable.Delete
End Sub

Public Sub ConvertTableToRange(aTable As ListObject)
    ' NOTE: You have to wrap this function in a test
    ' to see if the aTable ListObject is NOT NOTHING
    Dim headerRng As Range
    Set headerRng = aTable.HeaderRowRange
    
    aTable.Unlist   'This converts the Table to a Simple Range
    
    headerRng.AutoFilter
End Sub




