Attribute VB_Name = "colorfuncs"
Private Const mod_name as String = "colorfuncs"
Private Const module_author as String = "Ben Fisher"
Private Const module_version as String = "0.0.3"

Public Function rgb_to_hsl(ByVal rgb_string As String) As Variant
    'Note that pure white and black return errors

    Dim color_scale As Long: color_scale = 255
    Dim rgb_arr As Variant
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim c_max As Double
    Dim c_min As Double
    Dim c_delta As Double
    Dim c_avg As Double
    Dim arr(2) As Double

    Dim hue As Double
    Dim sat As Double
    Dim light As Double

    rgb_arr = split_rgb_string(rgb_string)

    r = rgb_arr(0) / color_scale
    g = rgb_arr(1) / color_scale
    b = rgb_arr(2) / color_scale
    
    c_max = WorksheetFunction.Max(r, g, b)
    c_min = WorksheetFunction.Min(r, g, b)
    c_delta = c_max - c_min
    
    If c_delta = 0 Then
        arr(0) = 0
        arr(1) = 0
        arr(2) = c_max * 100
        rgb_to_hsl = arr
    Else
        light = (c_max + c_min) / 2
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
        If c_max = 0 Then sat = 0 Else sat = c_delta / c_max
        arr(0) = hue
        arr(1) = sat * 100
        arr(2) = light * 100
        rgb_to_hsl = arr
    End If
End Function

Public Function hsl_to_rgb(ByVal hsl_string As String) As String
    'Note H is 360 scale, S and V or B on 100 scale
    'Note that pure white and black return errors

    Dim color_scale As Double
    Dim chroma As Double
    'The VBA mod function does not really work as expected
    Dim mod_term As Double
    Dim x As Double
    Dim m As Double
    Dim hsb_arr As Variant
    Dim arr As Variant

    hsl_arr = split_hsb_string(hsl_string)

    color_scale = 255
    hue = hsl_arr(0)
    sat = hsl_arr(1) / 100
    light = hsl_arr(2) / 100

    chroma = (1 - Abs(2 * light - 1)) * sat
    mod_term = ((hue / 60) / 2 - Int((hue / 60) / 2)) * 2
    x = chroma * (1 - Abs(mod_term - 1))
    m = light - chroma / 2
    
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
    
    hsl_to_rgb = Join(arr, ", ")
End Function

Public Function hsb_to_hsl(ByVal hsb_string As String)
    'TODO: Still needs a little work
    
    Dim arr(2) As Variant
    hsb_arr = split_hsb_string(hsb_string)
    
    hue = hsb_arr(0)
    sat_hsb = hsb_arr(1) / 100
    bright = hsb_arr(2) / 100
    
    light = bright * (1 - sat_hsb / 2)
    
    If light = 0 Or light = 1 Then
        sat_hsl = 0
    Else
        sat_hsl = (bright - light) / WorksheetFunction.Min(light, 1 - light)
    End If
    arr(0) = Round(hue, 3)
    arr(1) = Round(sat_hsl * 100, 3)
    arr(2) = Round(light * 100, 3)
    
    hsb_to_hsl = Join(arr, ", ")
    
End Function

Public Function hsl_to_hsb(ByVal hsl_string As String)
    'TODO: Still needs a little work
    Dim arr(2) As Variant
    hsl_arr = split_hsb_string(hsl_string)
    
    hue = hsl_arr(0)
    sat_hsl = hsl_arr(1) / 100
    lum = hsl_arr(2) / 100
    bright = lum + hsl_arr(1) * WorksheetFunction.Min(lum, 1 - lum)
    
    If bright = 0 Then
        sat_hsb = 0
    Else
        sat_hsb = 2 * (1 - lum / bright)
    End If
    arr(0) = Round(hue, 3)
    arr(1) = Round(sat_hsb * 100, 3)
    arr(2) = Round(lum * 100, 3)
    
    hsl_to_hsb = Join(arr, ", ")
End Function
    
Public Function rgb_to_hsb(ByVal rgb_string As String) As Variant
    'Note that pure white and black return errors

    Dim color_scale As Long: color_scale = 255
    Dim rgb_arr As Variant
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
    
    If c_delta <> 0 Then
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
        arr(0) = Round(hue, 3)
        arr(1) = Round(sat, 3)
        arr(2) = Round(bright, 3)
    Else
        arr(0) = 0
        arr(1) = 0
        arr(2) = c_max * 100
    End If

    rgb_to_hsb = arr
End Function

Public Function hsb_to_rgb(ByVal hsb_string As String) as String
    'Note H is 360 scale, S and V or B on 100 scale
    'Note that pure white and black return errors

    Dim color_scale As Double
    Dim chroma As Double
    'The VBA mod function does not really work as expected
    Dim mod_term as Double
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
    mod_term = ((hue / 60) / 2 - Int((hue / 60) / 2)) * 2
    x = chroma * (1 - Abs(mod_term - 1))
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

Public Function clean_hsb_string(ByVal hsb_string As String) As String
    Dim clean_arr As Variant
    Dim arr As Variant
    Dim i As Long
    clean_arr = Array("(", ")", " ", "h", "s", "b", "v", "=", "deg", "°", "%")
    For i = LBound(clean_arr) To UBound(clean_arr)
        hsb_string = Replace(hsb_string, clean_arr(i), "")
    Next
    clean_hsb_string = hsb_string
End Function

Public Function split_hsb_string(ByVal hsb_string As String) As Variant
    Dim hsb_arr As Variant
    hsb_arr = Split(clean_hsb_string(hsb_string), ",")
    split_hsb_string = hsb_arr
End Function



Public Function long_to_rgb(ByVal a_long As Long) as String
    Dim r, g, b as Double
    Dim arr(2) as Variant
    b = a_long \ 65536
    g = (a_long - b * 65536) \ 256
    r = a_long - b * 65536 - g * 256
    arr(0) = r: arr(1) = g: arr(2) = b
    long_to_rgb = Join(arr, ", ")
End Function

Public Function rgb_to_long(ByVal rgb_string as String) As Long
    Dim rgb_arr as Variant
    rgb_arr = split_rgb_string(rgb_string)
    rgb_to_long = RGB(rgb_arr(0), rgb_arr(1), rgb_arr(2))
End Function

Public Function hex_to_rgb(ByVal hex_color As String, Optional as_string As Boolean = True) As Variant
    'Returns a hex color value as an RGB array
    'Several prefixes are used to identify hex numbers: "&H" is used
    'in VBA, however, "#" is used for webcolors, also "0h" or "0x"
    'are sometimes used as well.
    Dim remove_characters As Variant
    Dim rgb_arr As Variant
    Dim i As Long
    Dim r As Variant
    Dim g As Variant
    Dim b As Variant
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
    If as_string = True Then rgb_arr = Join(rgb_arr, ", ")
    hex_to_rgb = rgb_arr
End Function


Public Function rgb_to_hex(ByVal rgb_string As String) As String
    arr = split_rgb_string(rgb_string)
    For i = LBound(arr) To UBound(arr)
        arr(i) = WorksheetFunction.Dec2Hex(arr(i))
        If Len(arr(i)) < 2 Then arr(i) = "0" & arr(i)
    Next
    rgb_to_hex = "#" & Join(arr, "")
End Function

Public Function clean_rgb_string(ByVal rgb_string As String) As String
    Dim clean_arr As Variant
    Dim arr As Variant
    Dim i As Long
    clean_arr = Array("(", ")", " ", "r", "g", "b", "=")
    For i = LBound(clean_arr) To UBound(clean_arr)
        rgb_string = Replace(rgb_string, clean_arr(i), "")
    Next
    clean_rgb_string = rgb_string
End Function

Public Function split_rgb_string(ByVal rgb_string As String) As Variant
    Dim rgb_arr As Variant
    rgb_arr = Split(clean_rgb_string(rgb_string), ",")
    split_rgb_string = rgb_arr
End Function

Public Function apply_contrasting_font_color(ByVal background_color As Long)
    'Based on W3.org visibility recommendations:
    'https://www.w3.org/TR/AERT/#color-contrast
    Dim arr As Variant
    Dim color_constant As Long
    Dim color_brightness As Double
    
    rgb_string = long_to_rgb(background_color)
    arr = split_rgb_string(rgb_string)
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

Public Function contrast_ratio(rgb_string_1 As String, rgb_string_2 As String) As Double
    Dim lum_1 As Double
    Dim lum_2 As Double
    Dim lum_min As Double
    Dim lum_max As Double
    lum_1 = relative_luminance(rgb_string_1)
    lum_2 = relative_luminance(rgb_string_2)
    If lum_1 > lum_2 Then
        lum_max = lum_1
        lum_min = lum_2
    Else
        lum_max = lum_2
        lum_min = lum_1
    End If
    contrast_ratio = (lum_max + 0.05) / (lum_min + 0.05)
End Function

Public Function get_hue(rgb_string As String) As Double
    'Returns the hue in degrees, where Red is 0.
    get_hue = rgb_to_hsb(rgb_string)(0)
End Function

Public Function get_saturation(rgb_string As String) As Double
    'Returns the saturation as a value between 0 (white) to 1 (color saturation)
    get_saturation = rgb_to_hsb(rgb_string)(1) / 100
End Function

Public Function get_brightness(rgb_string As String) As Double
    'Returns the brightness as a value between 0 (black) to 1 (brightness)
    get_brightness = rgb_to_hsb(rgb_string)(2) / 100
End Function

Public Sub color_selection_rgb()
    'Helper function for coloring Excel cells that contain an RGB code
    Dim rgb_string As String
    For Each a_cell In Selection.Cells
        rgb_string = a_cell.Value
        back_color = rgb_to_long(rgb_string)
        font_color = apply_contrasting_font_color(back_color)
        With a_cell
            .Interior.Color = back_color
            .Font.Color = font_color
        End With
    Next
End Sub

Public Sub color_selection_hex()
    'Helper function for coloring Excel cells that contain an Hex code
    Dim hex_string As String
    For Each a_cell In Selection.Cells
        hex_string = a_cell.Value
        back_color = rgb_to_long(hex_to_rgb(hex_string))
        font_color = apply_contrasting_font_color(back_color)
        With a_cell
            .Interior.Color = back_color
            .Font.Color = font_color
        End With
    Next
End Sub

Public Sub color_selection_hsb()
    'Helper function for coloring Excel cells that contain an HSB code
    Dim hsb_string As String
    For Each a_cell In Selection.Cells
        hsb_string = a_cell.Value
        back_color = rgb_to_long(hsb_to_rgb(hsb_string))
        font_color = apply_contrasting_font_color(back_color)
        With a_cell
            .Interior.Color = back_color
            .Font.Color = font_color
        End With
    Next
End Sub

Public Sub reset_selection_color()
    'Helper function to undo the coloring of selected cells in Excel.
    With Selection
        With .Interior
            .ColorIndex = xlAutomatic
            .Pattern = xlNone
        End With
        .Font.ColorIndex = xlAutomatic
    End With
End Sub


Public Function get_complement(ByVal rgb_string As String) As String
    Dim hsb_arr As Variant
    Dim arr(2) As Variant
    hsb_arr = rgb_to_hsb(rgb_string)
    For i = 0 To 2
        arr(i) = hsb_arr(i)
    Next
    arr(0) = (arr(0) + 180) Mod 360
    get_complement = hsb_to_rgb(Join(arr, ", "))
End Function

Public Function get_triad(ByVal rgb_string As String, Optional is_transposed As Boolean = False) As Variant
    Dim hsb_1 As Variant
    Dim hsb_2 As Variant
    Dim arr(2) As Variant

    hsb_1 = rgb_to_hsb(rgb_string)
    hsb_2 = rgb_to_hsb(rgb_string)
    hsb_1(0) = (hsb_1(0) + 120) Mod 360
    hsb_2(0) = (hsb_2(0) + 240) Mod 360

    arr(0) = rgb_string
    arr(1) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
    arr(2) = hsb_to_rgb(hsb_2(0) & ", " & hsb_2(1) & ", " & hsb_2(2))
    If is_transposed Then
        get_triad = WorksheetFunction.Transpose(arr)
    Else
        get_triad = arr
    End If
End Function


Public Function get_split_complement(ByVal rgb_string As String, Optional is_transposed As Boolean = False) As Variant
    Dim hsb_1 As Variant
    Dim hsb_2 As Variant
    Dim arr(2) As Variant

    hsb_1 = rgb_to_hsb(rgb_string)
    hsb_2 = rgb_to_hsb(rgb_string)
    hsb_1(0) = (hsb_1(0) + 150) Mod 360
    hsb_2(0) = (hsb_2(0) + 210) Mod 360

    arr(0) = rgb_string
    arr(1) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
    arr(2) = hsb_to_rgb(hsb_2(0) & ", " & hsb_2(1) & ", " & hsb_2(2))
    If is_transposed Then
        get_split_complement = WorksheetFunction.Transpose(arr)
    Else
        get_split_complement = arr
    End If
End Function

Public Function get_analogous(ByVal rgb_string As String, _
    Optional is_ordered As Boolean = True, _
    Optional is_transposed As Boolean = False) As Variant
    Dim hsb_1 As Variant
    Dim hsb_2 As Variant
    Dim arr(2) As Variant

    hsb_1 = rgb_to_hsb(rgb_string)
    hsb_2 = rgb_to_hsb(rgb_string)
    hsb_1(0) = (hsb_1(0) + 30) Mod 360
    hsb_2(0) = (hsb_2(0) + 330) Mod 360

    If is_ordered Then
        arr(0) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
        arr(1) = rgb_string
    Else
        arr(1) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
        arr(0) = rgb_string
    End If
    arr(2) = hsb_to_rgb(hsb_2(0) & ", " & hsb_2(1) & ", " & hsb_2(2))
    If is_transposed Then
        get_analogous = WorksheetFunction.Transpose(arr)
    Else
        get_analogous = arr
    End If
End Function

Public Function get_tetradic(ByVal rgb_string As String, _
                            Optional rotate_CW As Boolean = True, _
                            Optional is_transposed As Boolean = False) As Variant
    Dim hsb_1 As Variant
    Dim hsb_2 As Variant
    Dim hsb_3 As Variant
    Dim arr(3) As Variant
    
    hsb_1 = rgb_to_hsb(rgb_string)
    hsb_2 = rgb_to_hsb(rgb_string)
    hsb_3 = rgb_to_hsb(rgb_string)
    
    If rotate_CW = True Then
        hsb_1(0) = (hsb_1(0) + 120) Mod 360
        hsb_2(0) = (hsb_2(0) + 180) Mod 360
        hsb_3(0) = (hsb_3(0) + 300) Mod 360
    Else
        hsb_1(0) = (hsb_1(0) + 60) Mod 360
        hsb_2(0) = (hsb_2(0) + 180) Mod 360
        hsb_3(0) = (hsb_3(0) + 240) Mod 360
    End If

    arr(0) = rgb_string
    arr(1) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
    arr(2) = hsb_to_rgb(hsb_2(0) & ", " & hsb_2(1) & ", " & hsb_2(2))
    arr(3) = hsb_to_rgb(hsb_3(0) & ", " & hsb_3(1) & ", " & hsb_3(2))
    If is_transposed Then
        get_tetradic = WorksheetFunction.Transpose(arr)
    Else
        get_tetradic = arr
    End If
End Function

Public Function get_square(ByVal rgb_string As String, Optional is_transposed As Boolean = False) As Variant
    Dim hsb_1 As Variant
    Dim hsb_2 As Variant
    Dim hsb_3 As Variant
    Dim arr(3) As Variant
    
    hsb_1 = rgb_to_hsb(rgb_string)
    hsb_2 = rgb_to_hsb(rgb_string)
    hsb_3 = rgb_to_hsb(rgb_string)
    
    hsb_1(0) = (hsb_1(0) + 90) Mod 360
    hsb_2(0) = (hsb_2(0) + 180) Mod 360
    hsb_3(0) = (hsb_3(0) + 270) Mod 360

    arr(0) = rgb_string
    arr(1) = hsb_to_rgb(hsb_1(0) & ", " & hsb_1(1) & ", " & hsb_1(2))
    arr(2) = hsb_to_rgb(hsb_2(0) & ", " & hsb_2(1) & ", " & hsb_2(2))
    arr(3) = hsb_to_rgb(hsb_3(0) & ", " & hsb_3(1) & ", " & hsb_3(2))
    If is_transposed Then
        get_square = WorksheetFunction.Transpose(arr)
    Else
        get_square = arr
    End If
End Function

Public Function random_rgb()
    random_rgb = Join(Array(WorksheetFunction.RandomBetween(0, 255), _
                            WorksheetFunction.RandomBetween(0, 255), _
                            WorksheetFunction.RandomBetween(0, 255)), ", ")
End Function

Public Function random_hex()

    random_hex = "#" & Join(Array(WorksheetFunction.DEC2HEX(WorksheetFunction.RandomBetween(0, 255)), _
                                  WorksheetFunction.DEC2HEX(WorksheetFunction.RandomBetween(0, 255)), _
                                  WorksheetFunction.DEC2HEX(WorksheetFunction.RandomBetween(0, 255)), "") 

End Function

Public Function random_long()

    random_long = rgb_to_long(random_rgb())

End Function

Public Function alpha_blend(alpha, original_rgb_string, target_rgb_string)
    Dim arr(2) As Variant
    Dim orig_arr As Variant
    Dim target_arr As Variant
    orig_arr = split_rgb_string(original_rgb_string)
    target_arr = split_rgb_string(target_rgb_string)
    For i = LBound(orig_arr) To UBound(orig_arr)
        arr(i) = CLng(orig_arr(i) + alpha * (target_arr(i) - orig_arr(i)))
    Next
    alpha_blend = Join(arr, ", ")
End Function

Public Function alpha_shade(alpha, original_rgb_string, Optional target_rgb_string As String = "0,0,0")
    Dim arr(2) As Variant
    Dim orig_arr As Variant
    Dim target_arr As Variant
    orig_arr = split_rgb_string(original_rgb_string)
    target_arr = split_rgb_string(target_rgb_string)
    For i = LBound(orig_arr) To UBound(orig_arr)
        arr(i) = CLng(orig_arr(i) + alpha * (target_arr(i) - orig_arr(i)))
    Next
    alpha_shade = Join(arr, ", ")
End Function

Public Function alpha_tint(alpha, original_rgb_string, Optional target_rgb_string As String = "255,255,255")
    Dim arr(2) As Variant
    Dim orig_arr As Variant
    Dim target_arr As Variant
    orig_arr = split_rgb_string(original_rgb_string)
    target_arr = split_rgb_string(target_rgb_string)
    For i = LBound(orig_arr) To UBound(orig_arr)
        arr(i) = CLng(orig_arr(i) + alpha * (target_arr(i) - orig_arr(i)))
    Next
    alpha_tint = Join(arr, ", ")
End Function

Public Function alpha_tone(alpha, original_rgb_string, Optional target_rgb_string As String = "128,128,128")
    Dim arr(2) As Variant
    Dim orig_arr As Variant
    Dim target_arr As Variant
    orig_arr = split_rgb_string(original_rgb_string)
    target_arr = split_rgb_string(target_rgb_string)
    For i = LBound(orig_arr) To UBound(orig_arr)
        arr(i) = CLng(orig_arr(i) + alpha * (target_arr(i) - orig_arr(i)))
    Next
    alpha_tone = Join(arr, ", ")
End Function

Public Function tint_rgb(tint_factor as Double, rgb_string as String) As String
    Dim rgb_arr as As Variant
    Dim arr as Variant
    rgb_arr = split_rgb_string(rgb_string)
    For i = 0 to 2
        arr(i) = rgb_arr(i) + ((255 - rgb_arr(i) * tint_factor))
    Next
    tint_rgb = Join(arr, ", ")
End Function

Public Function shade_rgb(shade_factor as Double, rgb_string as String) As String
    Dim rgb_arr as As Variant
    Dim arr as Variant
    rgb_arr = split_rgb_string(rgb_string)
    For i = 0 to 2
        arr(i) = rgb_arr(i) * (1 - shade_factor)
    Next
    shade_rgb = Join(arr, ", ")
End Function

Public Function tint_hex(tint_factor as Double, hex_string as String) As String
    Dim rgb_string as String
    Dim new_rgb as String
    Dim rgb_arr as As Variant
    Dim arr as Variant
    rbg_string = hex_to_rgb(hex_string)
    rgb_arr = split_rgb_string(rgb_string)
    For i = 0 to 2
        arr(i) = rgb_arr(i) + ((255 - rgb_arr(i) * tint_factor))
    Next
    new_rgb = Join(arr, ", ")
    tint_hex = rgb_to_hex(new_rgb)
End Function

Public Function shade_hex(shade_factor as Double, hex_string as String) As String
    Dim rgb_string as String
    Dim new_rgb as String
    Dim rgb_arr as As Variant
    Dim arr as Variant
    rbg_string = hex_to_rgb(hex_string)
    rgb_arr = split_rgb_string(rgb_string)
    For i = 0 to 2
        arr(i) = rgb_arr(i) * (1 - shade_factor)
    Next
    new_rgb = Join(arr, ", ")
    shade_hex = rgb_to_hex(new_rgb)
End Function