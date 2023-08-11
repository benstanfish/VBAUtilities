Attribute VB_Name = "colorfuncs"
Private Const mod_name as String = "colorfuncs"
Private Const module_author as String = "Ben Fisher"
Private Const module_version as String = "0.0.3"

Public Function rgb_to_hsb(rgb_string As String) as Variant
    Dim color_scale As Long: color_scale = 255
    Dim rgb_arr as Variant
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

    rgb_arr = split_rgb_string(rgb_string)

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
    arr(1) = sat
    arr(2) = bright

    rgb_to_hsb = arr

End Function

Public Function clean_hsb_string(hsb_string As String) As String
    Dim clean_arr As Variant
    Dim arr As Variant
    Dim i As Long
    clean_arr = Array("(", ")", " ", "h", "s", "b", "v", "=", "deg", "°", "%")
    For i = LBound(clean_arr) To UBound(clean_arr)
        hsb_string = Replace(hsb_string, clean_arr(i), "")
    Next
    clean_hsb_string = hsb_string
End Function

Public Function split_hsb_string(hsb_string As String) As Variant
    Dim hsb_arr As Variant
    hsb_arr = Split(clean_hsb_string(hsb_string), ",")
    split_hsb_string = hsb_arr
End Function

Public Function hsb_to_rgb(hsb_string As String) as String
    'Note H is 360 scale, S and V or B on 100 scale
    
    Dim color_scale As Double
    Dim chroma As Double
    Dim x As Double
    Dim m As Double
    Dim hsb_arr as Variant
    Dim arr as Variant

    hsb_arr = split_hsb_string(hsb_string)

    color_scale = 255
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
    
    arr(0) = CLng((arr(0) + m) * color_scale)
    arr(1) = CLng((arr(1) + m) * color_scale)
    arr(2) = CLng((arr(2) + m) * color_scale)
    
    hsb_to_rgb = Join(arr, ", ")
End Function

Public Function long_to_rgb(a_long As Long) as String
    Dim r, g, b as Double
    Dim arr(0 To 2) as Long
    b = a_long \ 65536
    g = (a_long - b * 65536) \ 256
    r = a_long - b * 65536 - g * 256
    arr(0) = r: arr(1) = g: arr(2) = b
    long_to_rgb = Join(arr, ", ")
End Function

Public Function rgb_to_long(rgb_string as String) As Long
    Dim rgb_arr as Variant
    rgb_arr = split_rgb_string(rgb_string)
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


Public Function clean_color_string(color_string as String) as Variant
    'Accepts rgb, hex and hsb color strings, cleans the unnecessary
    'characters out of the string and returns a triplet array.

End Function


Public Function rgb_to_hex(rgb_string As String) As String
    arr = split_rgb_string(rgb_string)
    For i = LBound(arr) To UBound(arr)
        arr(i) = WorksheetFunction.Dec2Hex(arr(i))
        If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
    Next
    rgb_to_hex = "#" & Join(arr, "")
End Function

Public Function clean_rgb_string(rgb_string As String) As String
    Dim clean_arr As Variant
    Dim arr As Variant
    Dim i As Long
    clean_arr = Array("(", ")", " ", "r", "g", "b", "=")
    For i = LBound(clean_arr) To UBound(clean_arr)
        rgb_string = Replace(rgb_string, clean_arr(i), "")
    Next
    clean_rgb_string = rgb_string
End Function



Public Function split_rgb_string(rgb_string As String) As Variant
    Dim rgb_arr As Variant
    rgb_arr = Split(clean_rgb_string(rgb_string), ",")
    split_rgb_string = rgb_arr
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

Public Function relative_luminance(rgb_string As String)
    Dim arr(2)
    rbg_arr = split_rgb_string(rgb_string)
    For i = 0 To 2
        arr(i) = rbg_arr(i) / 255
        If arr(i) <= 0.03928 Then arr(i) = arr(i) / 12.92 Else arr(i) = ((arr(i) + 0.055) / 1.055) ^ 2.4
    Next
    relative_luminance = 0.2126 * arr(0) + 0.7152 * arr(1) + 0.0722 * arr(2)
End Function

Public Function contrast_ratio(rgb_color_1 As String, rbg_color_2 As String) As Double
    Dim lum_1 As Double
    Dim lum_2 As Double
    Dim lum_min As Double
    Dim lum_max As Double
    lum_1 = relative_luminance(rgb_color_1)
    lum_2 = relative_luminance(rbg_color_2)
    If lum_1 > lum_2 Then
        lum_max = lum_1
        lum_min = lum_2
    Else
        lum_max = lum_2
        lum_min = lum_1
    End If
    contrast_ratio = (lum_max + 0.05) / (lum_min + 0.05)
End Function