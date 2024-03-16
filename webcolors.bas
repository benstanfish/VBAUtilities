Attribute VB_Name = "webcolors"
Public Const mod_name As String = "webcolors"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.2"
Public Const module_date As Date = #3/16/2024#

' Bootstrap v5.3 Alert Colors
' Note that DARKER as Background goes well with basic
' and DARK background goes well with LIGHT

Public Const DANGER As Long = 9406186
Public Const WARNING As Long = 7002879
Public Const SUCCESS As Long = 6269045
Public Const PRIMARY As Long = 16689262
Public Const SECONDARY As Long = 11644071

Public Const DANGER_LIGHT As Long = 14342136
Public Const WARNING_LIGHT As Long = 13497343
Public Const SUCCESS_LIGHT As Long = 14542801
Public Const PRIMARY_LIGHT As Long = 16769743
Public Const SECONDARY_LIGHT As Long = 15066082

Public Const DANGER_DARK As Long = 1840472
Public Const WARNING_DARK As Long = 216422
Public Const SUCCESS_DARK As Long = 2242058
Public Const PRIMARY_DARK As Long = 6630405
Public Const SECONDARY_DARK As Long = 3288875

Public Const DANGER_DARKER As Long = 919084
Public Const WARNING_DARKER As Long = 75571
Public Const SUCCESS_DARKER As Long = 1121029
Public Const PRIMARY_DARKER As Long = 3349251
Public Const SECONDARY_DARKER As Long = 1644310

'Module that contains the 140 named webcolors as Long
'simply add this module to your project to be able to access
'the named colors

'Pink colors
Public Const MEDIUMVIOLETRED As Long = 8721863
Public Const DEEPPINK As Long = 9639167
Public Const PALEVIOLETRED As Long = 9662683
Public Const HOTPINK As Long = 11823615
Public Const LIGHTPINK As Long = 12695295
Public Const PINK As Long = 13353215

'Red colors
Public Const DARKRED As Long = 139
Public Const RED As Long = 255
Public Const FIREBRICK As Long = 2237106
Public Const CRIMSON As Long = 3937500
Public Const INDIANRED As Long = 6053069
Public Const LIGHTCORAL As Long = 8421616
Public Const SALMON As Long = 7504122
Public Const DARKSALMON As Long = 8034025
Public Const LIGHTSALMON As Long = 8036607

'Orange colors
Public Const ORANGERED As Long = 17919
Public Const TOMATO As Long = 4678655
Public Const DARKORANGE As Long = 36095
Public Const CORAL As Long = 5275647
Public Const ORANGE As Long = 42495

'Yellow colors
Public Const DARKKHAKI As Long = 7059389
Public Const GOLD As Long = 55295
Public Const KHAKI As Long = 9234160
Public Const PEACHPUFF As Long = 12180223
Public Const YELLOW As Long = 65535
Public Const PALEGOLDENROD As Long = 11200750
Public Const MOCCASIN As Long = 11920639
Public Const PAPAYAWHIP As Long = 14020607
Public Const LIGHTGOLDENRODYELLOW As Long = 13826810
Public Const LEMONCHIFFON As Long = 13499135
Public Const LIGHTYELLOW As Long = 14745599

'Brown colors
Public Const MAROON As Long = 128
Public Const BROWN As Long = 2763429
Public Const SADDLEBROWN As Long = 1262987
Public Const SIENNA As Long = 2970272
Public Const CHOCOLATE As Long = 1993170
Public Const DARKGOLDENROD As Long = 755384
Public Const PERU As Long = 4163021
Public Const ROSYBROWN As Long = 9408444
Public Const GOLDENROD As Long = 2139610
Public Const SANDYBROWN As Long = 6333684
Public Const TAN As Long = 9221330
Public Const BURLYWOOD As Long = 8894686
Public Const WHEAT As Long = 11788021
Public Const NAVAJOWHITE As Long = 11394815
Public Const BISQUE As Long = 12903679
Public Const BLANCHEDALMOND As Long = 13495295
Public Const CORNSILK As Long = 14481663

'Purple and magenta colors
Public Const INDIGO As Long = 8519755
Public Const PURPLE As Long = 8388736
Public Const DARKMAGENTA As Long = 9109643
Public Const DARKVIOLET As Long = 13828244
Public Const DARKSLATEBLUE As Long = 9125192
Public Const BLUEVIOLET As Long = 14822282
Public Const DARKORCHID As Long = 13382297
Public Const FUCHSIA As Long = 16711935
Public Const MAGENTA As Long = 16711935
Public Const SLATEBLUE As Long = 13458026
Public Const MEDIUMSLATEBLUE As Long = 15624315
Public Const MEDIUMORCHID As Long = 13850042
Public Const MEDIUMPURPLE As Long = 14381203
Public Const ORCHID As Long = 14053594
Public Const VIOLET As Long = 15631086
Public Const PLUM As Long = 14524637
Public Const THISTLE As Long = 14204888
Public Const LAVENDER As Long = 16443110

'Blue colors
Public Const MIDNIGHTBLUE As Long = 7346457
Public Const NAVY As Long = 8388608
Public Const DARKBLUE As Long = 9109504
Public Const MEDIUMBLUE As Long = 13434880
Public Const BLUE As Long = 16711680
Public Const ROYALBLUE As Long = 14772545
Public Const STEELBLUE As Long = 11829830
Public Const DODGERBLUE As Long = 16748574
Public Const DEEPSKYBLUE As Long = 16760576
Public Const CORNFLOWERBLUE As Long = 15570276
Public Const SKYBLUE As Long = 15453831
Public Const LIGHTSKYBLUE As Long = 16436871
Public Const LIGHTSTEELBLUE As Long = 14599344
Public Const LIGHTBLUE As Long = 15128749
Public Const POWDERBLUE As Long = 15130800

'Cyan colors
Public Const TEAL As Long = 8421376
Public Const DARKCYAN As Long = 9145088
Public Const LIGHTSEAGREEN As Long = 11186720
Public Const CADETBLUE As Long = 10526303
Public Const DARKTURQUOISE As Long = 13749760
Public Const MEDIUMTURQUOISE As Long = 13422920
Public Const TURQUOISE As Long = 13688896
Public Const AQUA As Long = 16776960
Public Const CYAN As Long = 16776960
Public Const AQUAMARINE As Long = 13959039
Public Const PALETURQUOISE As Long = 15658671
Public Const LIGHTCYAN As Long = 16777184

'Green colors
Public Const DARKGREEN As Long = 25600
Public Const GREEN As Long = 32768
Public Const DARKOLIVEGREEN As Long = 3107669
Public Const FORESTGREEN As Long = 2263842
Public Const SEAGREEN As Long = 5737262
Public Const OLIVE As Long = 32896
Public Const OLIVEDRAB As Long = 2330219
Public Const MEDIUMSEAGREEN As Long = 7451452
Public Const LIMEGREEN As Long = 3329330
Public Const LIME As Long = 65280
Public Const SPRINGGREEN As Long = 8388352
Public Const MEDIUMSPRINGGREEN As Long = 10156544
Public Const DARKSEAGREEN As Long = 9419919
Public Const MEDIUMAQUAMARINE As Long = 11193702
Public Const YELLOWGREEN As Long = 3329434
Public Const LAWNGREEN As Long = 64636
Public Const CHARTREUSE As Long = 65407
Public Const LIGHTGREEN As Long = 9498256
Public Const GREENYELLOW As Long = 3145645
Public Const PALEGREEN As Long = 10025880

'White colors
Public Const MISTYROSE As Long = 14804223
Public Const ANTIQUEWHITE As Long = 14150650
Public Const LINEN As Long = 15134970
Public Const BEIGE As Long = 14480885
Public Const WHITESMOKE As Long = 16119285
Public Const LAVENDERBLUSH As Long = 16118015
Public Const OLDLACE As Long = 15136253
Public Const ALICEBLUE As Long = 16775408
Public Const SEASHELL As Long = 15660543
Public Const GHOSTWHITE As Long = 16775416
Public Const HONEYDEW As Long = 15794160
Public Const FLORALWHITE As Long = 15792895
Public Const AZURE As Long = 16777200
Public Const MINTCREAM As Long = 16449525
Public Const SNOW As Long = 16448255
Public Const IVORY As Long = 15794175
Public Const WHITE As Long = 16777215

'Black and gray colors
Public Const BLACK As Long = 0
Public Const DARKSLATEGRAY As Long = 5197615
Public Const DIMGRAY As Long = 6908265
Public Const SLATEGRAY As Long = 9470064
Public Const GRAY As Long = 8421504
Public Const LIGHTSLATEGRAY As Long = 10061943
Public Const DARKGRAY As Long = 11119017
Public Const SILVER As Long = 12632256
Public Const LIGHTGRAY As Long = 13882323
Public Const GAINSBORO As Long = 14474460

Function ContrastText(bgColor As Long, _
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
