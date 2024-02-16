Attribute VB_Name = "colorThemes"


Private Sub Test()

    WriteXMLString "#4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"

End Sub


Sub WriteXMLString(ParamArray accents() As Variant)

    Dim xmlStringA As String, xmlStringB As String, xmlStringC As String
    Dim schemeName As String
    Dim dk1 As String, dk2 As String
    Dim lt1 As String, lt2 As String
    Dim arr As Variant
    
    dk1 = "000000"
    lt1 = "FFFFFF"
    dk2 = "44546A"
    lt2 = "E7E6E6"
    hlink = "0563C1"
    folHLink = "954F72"
    
    schemeName = "MyScheme"
    
    xmlStringA = _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & Chr(10) & _
        "<a:clrScheme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""" & schemeName & """>" & Chr(10) & _
        "<a:dk1><a:sysClr val=""windowText"" lastClr=""" & dk1 & """/></a:dk1>" & Chr(10) & _
        "<a:lt1><a:sysClr val=""window"" lastClr=""" & lt1 & """/></a:lt1>" & Chr(10) & _
        "<a:dk2><a:srgbClr val=""" & dk2 & """/></a:dk2>" & Chr(10) & _
        "<a:lt2><a:srgbClr val=""" & lt2 & """/></a:lt2>" & Chr(10)
    
    
    For i = 1 To 6
        xmlStringB = xmlStringB & "<a:accent" & i & "><a:srgbClr val=""" & Replace(accents(i - 1), "#", "") & """/></a:accent" & i & ">" & Chr(10)
    Next
        
    xmlStringC = _
        "<a:hlink><a:srgbClr val=""" & hlink & """/></a:hlink>" & Chr(10) & _
        "<a:folHlink><a:srgbClr val=""" & folHLink & """/></a:folHlink>" & Chr(10) & _
        "</a:clrScheme>"
    
    Debug.Print xmlStringA & xmlStringB & xmlStringC

End Sub


Sub WriteXMLStringFromSelection()

    Dim xmlStringA As String, xmlStringB As String, xmlStringC As String
    Dim schemeName As String
    Dim dk1 As String, dk2 As String
    Dim lt1 As String, lt2 As String
    Dim arr As Variant
    
    dk1 = "000000"
    lt1 = "FFFFFF"
    dk2 = "44546A"
    lt2 = "E7E6E6"
    hlink = "0563C1"
    folHLink = "954F72"
    
    schemeName = "MyScheme"
    
    xmlStringA = _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & Chr(10) & _
        "<a:clrScheme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""" & schemeName & """>" & Chr(10) & _
        "<a:dk1><a:sysClr val=""windowText"" lastClr=""" & dk1 & """/></a:dk1>" & Chr(10) & _
        "<a:lt1><a:sysClr val=""window"" lastClr=""" & lt1 & """/></a:lt1>" & Chr(10) & _
        "<a:dk2><a:srgbClr val=""" & dk2 & """/></a:dk2>" & Chr(10) & _
        "<a:lt2><a:srgbClr val=""" & lt2 & """/></a:lt2>" & Chr(10)
    
    For i = 1 To 6
        xmlStringB = xmlStringB & "<a:accent" & i & "><a:srgbClr val=""" & Replace(Selection(i), "#", "") & """/></a:accent" & i & ">" & Chr(10)
    Next
        
    xmlStringC = _
        "<a:hlink><a:srgbClr val=""" & hlink & """/></a:hlink>" & Chr(10) & _
        "<a:folHlink><a:srgbClr val=""" & folHLink & """/></a:folHlink>" & Chr(10) & _
        "</a:clrScheme>"
    
    Debug.Print xmlStringA & xmlStringB & xmlStringC

End Sub

