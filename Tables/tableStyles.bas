Attribute VB_Name = "TableStyles"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "tableStyles"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "2.1.2"
Public Const module_date As Date = #6/4/2024#
Public Const module_notes As String = _
    "This module creates TableStyles which are configurations set on " & _
    "the Workbook level, and are applied equally over all Worksheets."
Public Const module_license As String = "GNU General Public License, v3.0"

'***********************************************************************
'                          Referenced Libraries
'***********************************************************************
' NONE


'***********************************************************************
'                        Module Level Preferences
'***********************************************************************
Public Const HEADER_ROW_COLOR As Long = webcolors.DODGERBLUE
Public Const TABLE_HRULE_COLOR As Long = webcolors.GAINSBORO
Public Const TABLE_VRULE_COLOR As Long = webcolors.GAINSBORO

Public Const HEADER_ROW_VRULE As Boolean = False
Public Const TABLE_VRULE As Boolean = False


'***********************************************************************
'                           Utility Functions
'***********************************************************************
Public Function DoesTableStyleExist(styleName As String, wb As Workbook) As Boolean
    ' Checks to see if the TableStyle exists within a workbook WB.
    ' It removes all spaces as TableStyles do not permit spaces.
    ' This is an improvement over TableStyles.Exists() as it does not
    ' raise an error if the TableStyle doesn't exist
    
    Dim temp As Boolean
    For Each aStyle In ActiveWorkbook.TableStyles
        If aStyle.Name = styleName Then temp = True
    Next
    DoesTableStyleExist = temp
End Function


'***********************************************************************
'                            Main Functions
'***********************************************************************
Public Sub CreateTableStyle(styleName As String, wb As Workbook)
    ' This function will create the TableStyle of styleName, or if it
    ' exists it will reset the formats to the values in this function.

    If Not DoesTableStyleExist(styleName, wb) Then
        wb.TableStyles.Add (styleName)
        wb.TableStyles(styleName).ShowAsAvailableTableStyle = True
    End If
    With wb.TableStyles(styleName)
        'Use the xlTableStyleElementType enum to select table element
        With .TableStyleElements(xlHeaderRow)
            .Interior.Color = HEADER_ROW_COLOR
            .Font.Color = ContrastText(.Interior.Color)
            .Font.Bold = True
            'NOTE: You cannot edit the header row height or font size in the
            ' table style defINITion.
            If HEADER_ROW_VRULE Then
                With .Borders(xlInsideVertical)
                    .Color = ContrastText(wb.TableStyles(styleName).TableStyleElements(xlHeaderRow).Interior.Color)
                    .Weight = xlThin
                End With
            Else
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
        End With
        .TableStyleElements(xlRowStripe1).Clear
        With .TableStyleElements(xlRowStripe1)
            For i = xlEdgeTop To xlEdgeTop
                With .Borders(i)
                    .Color = TABLE_HRULE_COLOR
                    .Weight = xlThin
                End With
            Next
            If TABLE_VRULE Then
                With .Borders(xlInsideVertical)
                    .Color = TABLE_VRULE_COLOR
                    .Weight = xlThin
                End With
            Else
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
        End With
        .TableStyleElements(xlRowStripe2).Clear
        With .TableStyleElements(xlRowStripe2)
            For i = xlEdgeTop To xlEdgeTop
                With .Borders(i)
                    .Color = TABLE_HRULE_COLOR
                    .Weight = xlThin
                End With
            Next
            If TABLE_VRULE Then
                With .Borders(xlInsideVertical)
                    .Color = TABLE_VRULE_COLOR
                    .Weight = xlThin
                End With
            Else
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
        End With
    End With
End Sub

Public Sub ResetTableStyle(styleName As String, wb As Workbook)
    ' Simply a wrapper function for a different name
    CreateTableStyle styleName, wb
End Sub

Public Sub DeleteTableStyle(styleName As String, wb As Workbook)
    If DoesTableStyleExist(styleName, wb) Then wb.TableStyles(styleName).Delete
End Sub

