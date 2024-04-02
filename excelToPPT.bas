Attribute VB_Name = "excelToPPT"
Public Const mod_name As String = "excelToPPT"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.3.1"
Public Const module_date As Date = #4/2/2024#

' REQUIRED REFERENCES:
' - Microsoft PowerPoint 16.0 Object Model

Dim fPath As String

Public Enum TableType
    [_First]
    info = 1
    pdt = 2
    issues = 3
    sched = 4
    [_Last]
End Enum

Public Sub GenerateThisSlide()
    GenerateSlide ActiveSheet
End Sub

Public Sub GenerateAllSlides()
    Application.ScreenUpdating = False
    For i = 3 To ActiveWorkbook.Sheets.Count
        GenerateSlide ActiveWorkbook.Sheets(i)
    Next i
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateSlide(sht As Worksheet)
    
    Dim ppApp As PowerPoint.Application
    Dim ppPres As PowerPoint.Presentation
    Dim ppSlide As PowerPoint.Slide
    Dim fPath As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    On Error GoTo createNew
        Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
createNew:
    If ppApp Is Nothing Then
        Set ppApp = New PowerPoint.Application
        Set ppPres = ppApp.Presentations.Add
    Else
        ppApp.Visible = True
        Set ppPres = ppApp.ActivePresentation
    End If
    
    fPath = ppPres.Path
    
    Dim sFolders As Variant
    sFolders = Array("PDFs", "Images")
    For Each thing In sFolders
        newPath = fPath & "\" & thing & "\"
        If Not fso.FolderExists(newPath) Then fso.CreateFolder (newPath)
    Next
    
    fso.BuildPath Path:=fPath, Name:="PDFs"
    fso.BuildPath Path:=fPath, Name:="Images"
    
    Dim ppLayout As CustomLayout
    Set ppLayout = ppPres.SlideMaster.CustomLayouts(7)  ' 7 = blank slide... typically, lol
    
    Set ppSlide = ppPres.Slides.AddSlide(Index:=ppPres.Slides.Count + 1, pCustomLayout:=ppLayout)
    ppSlide.Name = "Tech Lead Slide " & ppPres.Slides.Count + 1
    ppSlide.Select
    
    Dim oLeft As Long
    Dim oTop As Long
    Dim oWidth As Long
    Dim oHeight As Long
    
    Dim projectTitle As PowerPoint.Shape
    With ppSlide
        oLeft = Application.InchesToPoints(0.3)
        oTop = Application.InchesToPoints(0.3)
        oWidth = Application.InchesToPoints(9.25)
        oHeight = Application.InchesToPoints(0.25)
        Set projectTitle = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With projectTitle.TextFrame2
            .WordWrap = False
            With .TextRange
                .Text = sht.ListObjects(TableType.info).Range(TableInfo.ProjectNameLine, 2) 'Project Title
                .Font.Name = "Aptos Narrow"
                .Font.Size = 18
                .Font.Bold = True
            End With
        End With
    End With
    
    ' Add the P2 Number
    Dim p2Number As PowerPoint.Shape
    With ppSlide
        oLeft = ppPres.PageSetup.SlideWidth - Application.InchesToPoints(0.65)
        oTop = Application.InchesToPoints(0.41)
        oWidth = Application.InchesToPoints(1.5)
        oHeight = Application.InchesToPoints(0.25)
        Set p2Number = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With p2Number.TextFrame2
            .WordWrap = False
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignRight
                .Text = "P2#: " & sht.ListObjects(TableType.info).Range(TableInfo.P2Line, 2)
                .Font.Name = "Aptos Display"
                .Font.Size = 12
                .Font.Bold = False
            End With
        End With
    End With
    
    ' Add the Project Design Team
    Dim pdtRoster As PowerPoint.Shape
    With ppSlide
        oLeft = Application.InchesToPoints(0.3)
        oTop = Application.InchesToPoints(0.75)
        oWidth = Application.InchesToPoints(6.25)
        oHeight = Application.InchesToPoints(1.25)
        
        Dim pdt As Variant
    
        Dim cnt As Long, i As Long
        cnt = sht.ListObjects(TableType.pdt).DataBodyRange.Rows.Count
        ReDim pdt(cnt - 1)
        Dim aCell As Range
        For i = 1 To sht.ListObjects(TableType.pdt).DataBodyRange.Rows.Count
            Set aCell = sht.ListObjects(TableType.pdt).DataBodyRange(i, 1)
            pdt(i - 1) = aCell.Value & ": " & aCell.Offset(0, 1).Value
        Next
    
        Set pdtRoster = ppSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With pdtRoster.TextFrame2
            .Column.Number = 3
            .WordWrap = True
            .AutoSize = msoAutoSizeNone
            With .TextRange
                .Text = Join(pdt, vbCrLf)
                .Font.Name = "Aptos"
                .Font.Size = 12
                .Font.Bold = False
            End With
        End With
        

    End With
    
    ' Highlight my own name
    Dim tRng As TextRange
    Set tRng = pdtRoster.TextFrame.TextRange
    On Error GoTo warpA
    
    Set foundText = tRng.Find(FindWhat:="TL: B Fisher")
    foundText.Font.Bold = True
        

warpA:
    On Error GoTo warpB
    ' Ghost N/As
    Set foundText = tRng.Find(FindWhat:="N/A")
    Do While Not (foundText Is Nothing)
        foundText.Font.Color = webcolors.SLATEGRAY
        Set foundText = tRng.Find(FindWhat:="N/A", After:=foundText.Start + foundText.Length - 1)
    Loop
        
warpB:
    Set foundText = tRng.Find(FindWhat:="TBD")
    Do While Not (foundText Is Nothing)
        foundText.Font.Color = webcolors.ORANGERED
        Set foundText = tRng.Find(FindWhat:="TBD", After:=foundText.Start + foundText.Length - 1)
    Loop
    On Error GoTo 0
    
    ' Add Project Info
    Dim projectInfo As PowerPoint.Shape
    With ppSlide
        oLeft = ppPres.PageSetup.SlideWidth - Application.InchesToPoints(0.5 + 2.2)
        oTop = Application.InchesToPoints(0.65)
        oWidth = Application.InchesToPoints(2.25)
        oHeight = Application.InchesToPoints(1)
        Set projectInfo = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        
        Dim pInfo As Variant
        pInfo = Array("PA: " & sht.ListObjects(TableType.info).Range(TableInfo.PALine, 2), _
                        sht.ListObjects(TableType.info).Range(TableInfo.CWELine, 2), _
                        "JES: " & sht.ListObjects(TableType.info).Range(TableInfo.JESLine, 2), _
                        sht.ListObjects(TableType.info).Range(TableInfo.ClientLine, 2), _
                        "Updated: " & Format(Now, "mm/dd/YY"))
        
        With projectInfo.TextFrame2
            .WordWrap = True
            .AutoSize = msoAutoSizeShapeToFitText
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignRight
                .Text = Join(pInfo, vbCrLf)
                .Font.Name = "Aptos"
                .Font.Size = 12
                .Font.Bold = False
            End With
        End With
        
        Set tRng = projectInfo.TextFrame.TextRange
        Set foundText = tRng.Find(FindWhat:=pInfo(1))
        With foundText
            .Font.Color = 5287936
            .Font.Bold = True
        End With
        
        Set foundText = tRng.Find(FindWhat:="CWE ? ECC")
        If Not foundText Is Nothing Then
            With foundText
                .Font.Bold = False
                .Font.Color = webcolors.SLATEGRAY
            End With
        End If
        
        Set foundText = tRng.Find(FindWhat:="CWE " & ChrW(8805) & " ECC")
        If Not foundText Is Nothing Then
            With foundText
                .Font.Color = webcolors.ORANGERED
            End With
        End If
        
    End With
    
    ' Copy in the project type logo
    Dim funding As String
    funding = LCase(sht.ListObjects(TableType.info).Range(TableInfo.FundingLine, 2))
    Select Case funding
        Case Is = "srm"
            Sheet1.Shapes("srm").Copy
        Case Is = "o&m"
            Sheet1.Shapes("om").Copy
        Case Is = "host nation"
            Sheet1.Shapes("hostnation").Copy
        Case Else
            Sheet1.Shapes("milcon").Copy
    End Select
    
    Dim logo As Variant
    Set logo = ppSlide.Shapes.Paste
    logo.Left = ppPres.PageSetup.SlideWidth - logo.Width - Application.InchesToPoints(0.55)
    logo.Top = Application.InchesToPoints(1.8)
    
    ' Add Black Bar
    Dim blackBar As PowerPoint.Shape
    With ppSlide
        oLeft = Application.InchesToPoints(0.38)
        oTop = Application.InchesToPoints(2.15)
        oWidth = ppPres.PageSetup.SlideWidth - Application.InchesToPoints(0.5)
        oHeight = Application.InchesToPoints(0)
        Set blackBar = .Shapes.AddConnector(msoConnectorStraight, oLeft, oTop, oWidth, oTop)
        With blackBar.Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 2.25
        End With
    End With

    'Add Milestones Header
    Dim milestones As PowerPoint.Shape
    With ppSlide
        oLeft = Application.InchesToPoints(0.3)
        oTop = Application.InchesToPoints(2.18)
        oWidth = Application.InchesToPoints(3)
        oHeight = Application.InchesToPoints(0.25)
        Set milestones = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With milestones.TextFrame2
            .WordWrap = False
            With .TextRange
                .Text = "Milestones"
                .Font.Name = "Aptos Display"
                .Font.Size = 18
                .Font.Bold = True
            End With
        End With
    End With
    
    'Add Updates Header
    Dim criticalUpdates As PowerPoint.Shape
    With ppSlide
        oLeft = Application.InchesToPoints(5.125)
        oTop = Application.InchesToPoints(2.18)
        oWidth = Application.InchesToPoints(3)
        oHeight = Application.InchesToPoints(0.25)
        Set criticalUpdates = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With criticalUpdates.TextFrame2
            .WordWrap = False
            With .TextRange
                .Text = "Critical / Outstanding Issues"
                .Font.Name = "Aptos Display"
                .Font.Size = 18
                .Font.Bold = True
            End With
        End With
    End With
    
    ' Copy and Paste Schedule Table
    Application.CutCopyMode = False
    sht.ListObjects(TableType.sched).Range.Copy
    
    With ppSlide.Shapes.Paste(1)
        .Top = Application.InchesToPoints(2.6)
        .Left = Application.InchesToPoints(0.4)
    End With
    
    
    Dim comments As PowerPoint.Shape
    Dim commentsArr As Variant
    ReDim commentsArr(WorksheetFunction.CountA(sht.ListObjects(TableType.issues).ListColumns("Show").DataBodyRange) - 1)
    i = 0
    For Each aRow In sht.ListObjects(TableType.issues).ListColumns("Show").DataBodyRange
        If aRow.Value <> "" Then
            commentsArr(i) = aRow.Offset(0, 1).Value
            i = i + 1
        End If
    Next
    
    With ppSlide
        oLeft = Application.InchesToPoints(5.125)
        oTop = Application.InchesToPoints(2.6)
        oWidth = ppPres.PageSetup.SlideWidth - Application.InchesToPoints(5.125 + 0.5)
        oHeight = Application.InchesToPoints(0.25)
        Set comments = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
        With comments.TextFrame2
            .WordWrap = True
            With .TextRange
                .Text = WorksheetFunction.TextJoin(vbCrLf, True, commentsArr)
                .ParagraphFormat.Bullet.Character = 8226
                .ParagraphFormat.SpaceAfter = 0.5
                .Font.Name = "Aptos"
                .Font.Size = 12.5
                .Font.Bold = False
            End With
        End With
    End With
    
    If sht.ListObjects(TableType.info).Range(10, 2) <> "" Then
        Dim watermark As PowerPoint.Shape
        oLeft = Application.InchesToPoints(1)
        oTop = Application.InchesToPoints(1)
        oWidth = Application.InchesToPoints(10)
        oHeight = Application.InchesToPoints(0.25)
        With ppSlide
            Set watermark = .Shapes.AddTextbox(msoTextOrientationHorizontal, oLeft, oTop, oWidth, oHeight)
            With watermark.TextFrame2
                .WordWrap = False
                .AutoSize = msoAutoSizeShapeToFitText
                With .TextRange
                    .Text = sht.ListObjects(TableType.info).Range(TableInfo.WatermarkLine, 2)
                    .Font.Name = "Aptos Black"
                    .Font.Size = 84
                End With
            End With
            watermark.Rotation = -20
            watermark.Left = (ppPres.PageSetup.SlideWidth - watermark.Width) / 2
            watermark.Top = (ppPres.PageSetup.SlideHeight - watermark.Height) / 2
        End With
    End If
    
    Application.CutCopyMode = False
    
    ppSlide.Export Filename:=fPath & "\PDFs\" & sht.Name & ".pdf", FilterName:="PDF"
    ppSlide.Export Filename:=fPath & "\Images\" & sht.Name & ".jpg", FilterName:="JPG"
    
    Set fso = Nothing
    Set ppApp = Nothing
    Set ppPres = Nothing
    Set ppLayout = Nothing
    Set ppSlide = Nothing

End Sub














