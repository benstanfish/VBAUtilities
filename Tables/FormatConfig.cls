VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pInteriorColor As Long

Private pFontName As String
Private pFontColor As Long
Private pFontBold As Boolean
Private pFontSize As Double


Property Let InteriorColor(userColor As Long)
    pInteriorColor = userColor
    pFontColor = ContrastText(userColor)
End Property
Property Get InteriorColor() As Long
    InteriorColor = pInteriorColor
End Property

Property Let fontName(fontFamily As String)
    pFontName = fontFamily
End Property
Property Get fontName() As String
    fontName = pFontName
End Property

Property Let FontColor(userColor As Long)
    pFontColor = userColor
End Property
Property Get FontColor() As Long
    FontColor = pFontColor
End Property

Property Let fontSize(userSize As Double)
    pFontSize = userSize
End Property
Property Get fontSize() As Double
    fontSize = pFontSize
End Property

Property Let Bold(isFontBold As Boolean)
    pFontBold = userIsFontBold
End Property
Property Get Bold() As Boolean
    Bold = pFontBold
End Property



Private Function ContrastText(bgColor As Long, _
    Optional darkColor As Long = vbBlack, _
    Optional lightColor As Long = vbWhite) As Long
    'Based on W3.org visibility recommendations:
    'https://www.w3.org/TR/AERT/#color-contrast
    
    Dim color_brightness As Double
    Dim r As Long, g As Long, b As Long
    
    b = bgColor \ 65536
    g = (bgColor - b * 65536) \ 256
    r = bgColor - b * 65536 - g * 256
    
    color_brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    If color_brightness > 0.55 Then ContrastText = darkColor Else ContrastText = lightColor
End Function

Private Sub Class_Initialize()
    Me.InteriorColor = vbWhite
    Me.fontName = "Arial"
    Me.fontSize = 10
End Sub

Public Sub INIT(intColor As Long, Optional fName As String = "Arial", _
                Optional fSize As Double = 10, Optional fBold As Boolean = False)

    Me.InteriorColor = intColor
    Me.fontName = fName
    Me.fontSize = fSize
    Me.Bold = fBold

End Sub
