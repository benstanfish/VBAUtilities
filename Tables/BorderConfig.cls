VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BorderConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pLineStyles() As Long
Private pWeights() As Long
Private pColors() As Long

Public Sub SetStyle(borderIndex As XlBordersIndex, _
                    borderColor As Long, _
                    Optional borderLineStyle As XlLineStyle = xlContinuous, _
                    Optional borderWeight As XlBorderWeight = xlThin)
    ' Cannot use a property method, so use a regular method instead.
    ' Class initializes with no borders, but if at least a color is
    ' set then it will "turn on" the style
    pColors(borderIndex) = borderColor
    pLineStyles(borderIndex) = borderLineStyle
    pWeights(borderIndex) = borderWeight
End Sub

Public Function LineStyles() As Long()
    LineStyles = pLineStyles
End Function

Public Function Weights() As Long()
    Weights = pWeights
End Function

Public Function Colors() As Long()
    Colors = pColors
End Function

Public Function GetBorder(borderIndex As XlBordersIndex) As Variant
    Dim arr(2) As Long
    arr(0) = pColor(borderIndex)
    arr(1) = pLineStyles(borderIndex)
    arr(2) = pWeight(borderIndex)
End Function

Private Sub Class_Initialize()
    Dim i As Long
    ReDim pLineStyles(5 To 12), pWeights(5 To 12), pColors(5 To 12)
    For i = 5 To 12
        pColors(i) = webcolors.GAINSBORO
        pLineStyles(i) = xlLineStyleNone
        pWeights(i) = xlThin
    Next
End Sub

Private Sub PrintStylesToDebug()
    For i = 5 To 12
        Debug.Print i, pLineStyles(i), pWeights(i), pColors(i)
    Next
End Sub
