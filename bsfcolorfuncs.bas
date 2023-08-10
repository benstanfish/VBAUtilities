Attribute VB_Name = "bsfcolorfuncs"
Private Const mod_name as String = "bsfcolorfuncs"
Private Const module_author as String = "Ben Fisher"
Private Const module_version as String = "0.0.1"

Public Function rgb_to_hsb(rgb_arr As Variant) as Variant
    Dim color_scale As Long: color_scale = 255

    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim c_max As Double
    Dim c_min As Double
    Dim c_delta As Double
    Dim arr(2) As Double

    Dim hue As Double
    Dim sat As Double
    Dim bright As Double

    r = rgb_arr(0) / color_scale
    g = rgb_arr(1) / color_scale
    b = rgb_arr(2) / color_scale
    
    c_max = WorksheetFunction.Max(r, g, b)
    c_min = WorksheetFunction.Min(r, g, b)
    
    c_delta = c_max - c_min
    
    If c_max = r And g >= b Then
        hue = 60 * (g - b) / c_delta
    ElseIf c_max = r And g < b Then
        hue = 60 * (g - b) / c_delta + 360
    ElseIf c_max = g Then
        hue = 60 * (b - r) / c_delta + 120
    ElseIf c_max = b Then
        hue = 60 * (r - g) / c_delta + 240
    Else
        hue = 0
    End If
    
    If c_max <> 0 Then sat = c_delta / c_max * 100
    
    bright = c_max * 100
    
    arr(0) = CLng(hue)
    arr(1) = CLng(sat)
    arr(2) = CLng(bright)

    rgb_to_hsb = arr

End Function

Public Function hsb_to_rgb(hsb_arr As Variant) as Variant
    'Note H is 360 scale, S and V or B on 100 scale
    
    Dim color_scale As Long: color_scale = 255
    
    Dim chroma As Double
    Dim x As Double
    Dim m As Double
        
    Dim arr(2) As Double

    hue = hsb_arr(0)
    sat = hsb_arr(1) / 100
    bright = hsb_arr(2) / 100

    chroma = bright * sat
    x = chroma * (1 - Abs(hue / 60 - Int(hue / 60) - 1))
    m = bright - chroma
    
    If hue >= 0 And hue < 60 Then
        arr = Array(chroma, x, 0)
    ElseIf hue >= 60 And hue < 120 Then
        arr = Array(x, chroma, 0)
    ElseIf hue >= 120 And hue < 180 Then
        arr = Array(0, chroma, x)
    ElseIf hue >= 180 And hue < 240 Then
        arr = Array(0, x, chroma)
    ElseIf hue >= 240 And hue < 300 Then
        arr = Array(x, 0, chroma)
    ElseIf hue >= 300 And hue < 360 Then
        arr = Array(chroma, 0, x)
    Else
        arr = Array(0, 0, 0)
    End If
    
    arr(0) = CLong((arr(0) + m) * color_scale)
    arr(1) = CLong((arr(1) + m) * color_scale)
    arr(2) = CLong((arr(2) + m) * color_scale)
    
    hsb_to_rgb = arr
End Function

Public Function long_to_rgb(a_long As Long) as Variant
    Dim r, g, b as Double
    Dim arr(0 To 2) as Long
    b = a_long \ 65536
    g = (a_long - b * 65536) \ 256
    r = a_long - b * 65536 - g * 256
    arr(0) = r: arr(1) = g: arr(2) = b
    long_to_rgb = arr
End Function

Public Function rgb_to_long(rgb_arr As Variant) As Long
    rgb_to_long = RGB(rgb_arr(0), rgb_arr(1), rgb_arr(2))
End Function

Public Function hex_to_rgb(hex_color As String, Optional as_string as Boolean = True) As Variant
    'Returns a hex color value as an RGB array
    'Several prefixes are used to identify hex numbers: "&H" is used
    'in VBA, however, "#" is used for webcolors, also "0h" or "0x"
    'are sometimes used as well.
    Dim remove_characters, rgb_arr As Variant
    Dim i As Long
    Dim r, g, b As String
    remove_characters = Array("&H", "&h", "#", "0H", "0h", "0X", "0x")
    For i = LBound(remove_characters) To UBound(remove_characters)
        hex_color = Replace(hex_color, remove_characters(i), "")
    Next i
    r = Mid(hex_color, 1, 2)
    g = Mid(hex_color, 3, 2)
    b = Mid(hex_color, 5, 2)
    rgb_arr = Array(WorksheetFunction.Hex2Dec(r), _
                        WorksheetFunction.Hex2Dec(g), _
                        WorksheetFunction.Hex2Dec(b))
    If as_string = True Then rgb_arr =  JOIN(rgb_arr, ", ")
    hex_to_rgb = rgb_arr
End Function

Public Function rgb_string_to_array(rgb_string as Variant)
    'This is a helper-function that cleans up RGB strings 
    'and converts them to an RGB array.
    Dim rgb_arr As Variant
    Dim clean_arr As Variant
    Dim clean_str As String
    clean_arr = Array("(", ")", " ", "r", "g", "b", "=")
    For i = 0 To UBound(clean_arr)
        rgb_string = Replace(rgb_string, clean_arr(i), "")
    Next
    rgb_arr = Split(rgb_string, ",")
    For i = LBound(rgb_arr) To UBound(rgb_arr)
        rgb_arr(i) = CLng(rgb_arr(i))
    Next
    rgb_string_to_array = rgb_arr
End Function

Public Function is_rgb_string(rgb_test as Variant) as Boolean
    On Error Goto dump
    If Len(Replace(rgb_test, ",", "")) < Len(rgb_test) Then rgb_test = True
dump:
End Function

Public Function rgb_to_hex(rgb_arr As Variant) As String
    'Can accept triplet-like strings, VBA arrays or Excel ranges of 3 values
    Dim arr As Variant
    If is_rgb_string(rgb_arr) = True Then
        arr = rgb_string_to_array(rgb_arr)
        For i = LBound(arr) To UBound(arr)
            arr(i) = WorksheetFunction.Dec2Hex(arr(i))
            If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
        Next
    Else
        ReDim arr(2)
        If TypeName(rgb_arr) = "Range" Then
            'Assumes Excel range vector (row or column) of 3 values
            For i = 0 To 2
                arr(i) = CLng(rgb_arr(i + 1))
                arr(i) = WorksheetFunction.Dec2Hex(arr(i))
                If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
            Next
        Else
            'Assumes VBA array
            For i = 0 To 2
                arr(i) = CLng(rgb_arr(i))
                arr(i) = WorksheetFunction.Dec2Hex(arr(i))
                If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
            Next
        End If
    End If
    rgb_to_hex = "#" & Join(arr, "")
End Function

Public Function rgb_to_hex2(rgb_arr As Variant) As String
    'Can accept triplet-like strings, VBA arrays or Excel ranges of 3 values
    Dim arr As Variant
    arr = rgb_to_array(rgb_arr)
    For i = LBound(arr) To UBound(arr)
        arr(i) = WorksheetFunction.Dec2Hex(arr(i))
        If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
    Next
    rgb_to_hex2 = "#" & Join(arr, "")
End Function


Public Function apply_contrasting_font_color(background_color As Long)
    'Based on W3.org visibility recommendations:
    'https://www.w3.org/TR/AERT/#color-contrast
    
    Dim arr As Variant
    Dim color_constant As Long
    Dim color_brightness As Double
    
    arr = long_to_rgb(background_color)
    color_brightness = (0.299 * arr(0) + 0.587 * arr(1) + 0.114 * arr(2)) / 255
    If color_brightness > 0.55 Then color_constant = vbBlack Else color_constant = vbWhite

    apply_contrasting_font_color = color_constant    
End Function

Public Function rgb_to_array(rgb_like as Variant)
    'Converts triplet-like strings, VBA arrays or Excel ranges of 3 values
    'and returns a VBA array of 3 values
    Dim arr As Variant
    If is_rgb_string(rgb_like) = True Then
        arr = rgb_string_to_array(rgb_like)
    Else
        ReDim arr(2)
        If TypeName(rgb_like) = "Range" Then
            'Assumes Excel range vector (row or column) of 3 values
            For i = 0 To 2
                arr(i) = CLng(rgb_like(i + 1))
            Next
        Else
            'Assumes VBA array
            For i = 0 To 2
                arr(i) = CLng(rgb_like(i))
            Next
        End If
    End If
    rgb_to_array = arr
End Function


Public Function relative_luminance(rgb_arr as Variant) as Double

    

End Function