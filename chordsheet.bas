Attribute VB_Name = "chordsheet"
Private Const mod_name As String = "chordsheet"
Private Const module_author As String = "Ben Fisher"
Private Const module_version As String = "1.1"
Private Const module_date As String = "2024/01/21"

'NOTE: Set a reference to Microsoft Scripting Runtime for Scripting.Dictionary

'==================================  GLOBALS  ==================================
' Note: GLOBAL constant names are in all-caps.

Public Const SHARPTAGS = "c,cs,d,ds,e,f,fs,g,gs,a,as,b" 'es and cs in F# and C#
Public Const FLATTAGS = "c,db,d,eb,e,f,gb,g,ab,a,bb,b"  'cf and ff in Gb and Cb

Public Const SHARPS = "C,C#,D,D#,E,F,F#,G,G#,A,A#,B"    'E# and C# in F# and C#
Public Const FLATS = "C,Db,D,Eb,E,F,Gb,G,Ab,A,Bb,B"     'Cb and Fb in Gb and Cb

'==============================  HELPER METHODS  ===============================

Function ParseToArray(namedConstant As String) As Variant
    ParseToArray = Split(namedConstant, ",")
End Function

Function XMOD(i As Variant, Optional div As Long = 12)
    XMOD = i - (div * Int(i / div))
End Function

Private Function GetCurrentKey()
    GetCurrentKey = LCase(ActiveDocument.ContentControls(1).Range.Text)
End Function

Private Sub DeletePart()
    ThisDocument.CustomXMLParts("chordnamespace").Delete
End Sub

Function TagDictionary() As Scripting.Dictionary

    Dim tagIndices As Scripting.Dictionary
    Set tagIndices = New Scripting.Dictionary
    
    tagIndices.Add "c", 0
    tagIndices.Add "cs", 1
    tagIndices.Add "df", 1
    tagIndices.Add "d", 2
    tagIndices.Add "ds", 3
    tagIndices.Add "ef", 3
    tagIndices.Add "e", 4
    tagIndices.Add "es", 5
    tagIndices.Add "ff", 4
    tagIndices.Add "f", 5
    tagIndices.Add "fs", 6
    tagIndices.Add "gf", 6
    tagIndices.Add "g", 7
    tagIndices.Add "gs", 8
    tagIndices.Add "af", 8
    tagIndices.Add "a", 9
    tagIndices.Add "as", 10
    tagIndices.Add "bf", 10
    tagIndices.Add "b", 11
    tagIndices.Add "cf", 11
    
    tagIndices.Add "c#", 1
    tagIndices.Add "db", 1
    tagIndices.Add "d#", 3
    tagIndices.Add "eb", 3
    tagIndices.Add "e#", 5
    tagIndices.Add "fb", 4
    tagIndices.Add "f#", 6
    tagIndices.Add "gb", 6
    tagIndices.Add "g#", 8
    tagIndices.Add "ab", 8
    tagIndices.Add "a#", 10
    tagIndices.Add "bb", 10
    tagIndices.Add "cb", 11
    
    Set TagDictionary = tagIndices
    Set tagIndices = Nothing
    
End Function

Function GetKeyType(aKey As Variant) As Variant
    Select Case LCase(aKey)
        Case Is = "g", "d", "a", "e", "b", "f#", "c#", "fs", "cs"
            GetKeyType = ParseToArray(SHARPS)
        Case Is = "f", "bb", "eb", "ab", "db", "gb", "cb"
            GetKeyType = ParseToArray(FLATS)
        Case Else
            GetKeyType = ParseToArray(SHARPS)
    End Select
End Function

Private Sub CleanXML()

    Dim parts As CustomXMLParts
    Dim part As CustomXMLPart
    Dim nodes As CustomXMLNodes
    Dim node As CustomXMLNode
    Dim rootNode As CustomXMLNode
    
    Set part = ThisDocument.CustomXMLParts("https://www.musicalnotes.com/")
    Set rootNode = part.DocumentElement
    
    For Each node In rootNode.ChildNodes
        If node.BaseName = "#text" Then node.Delete
    Next

End Sub

Sub ChangeKey()

    Application.ScreenUpdating = False
    
    Dim tags As Scripting.Dictionary
    Dim aKey As String
    Dim keyIndex As Variant
    Dim notesOfNewKey As Variant
    Dim newNotes(11) As String

    Dim part As CustomXMLPart
    Dim nodes As CustomXMLNodes
    Dim node As CustomXMLNode
    Dim rootNode As CustomXMLNode
    
    aKey = LCase(InputBox("Which key would you like to transpose to?", "Transposition Dialog"))
    
    If aKey <> "" Then
        
        ResetKey

        Set tags = TagDictionary()
        
        currentKey = GetCurrentKey()
        currentKeyIndex = tags(currentKey)
        
        keyIndex = tags(aKey)
        notesOfNewKey = GetKeyType(aKey)
        
        Set part = ThisDocument.CustomXMLParts("https://www.musicalnotes.com/")
        Set rootNode = part.DocumentElement

        For i = 0 To 11
            newIndex = XMOD(i + keyIndex - currentKeyIndex, 12)
            newNotes(i) = notesOfNewKey(newIndex)
        Next i

        CleanXML
    
        For Each node In rootNode.ChildNodes
            If node.BaseName <> "#text" Then
                currentIndex = tags(node.BaseName)
                node.Text = newNotes(currentIndex)
                If (aKey = "f#" Or aKey = "c#") And node.Text = "F" Then node.Text = "E#"
                If aKey = "c#" And node.Text = "C" Then node.Text = "B#"
                If (aKey = "gb" Or aKey = "cb") And node.Text = "B" Then node.Text = "Cb"
                If aKey = "cb" And node.Text = "E" Then node.Text = "Fb"
            End If

        Next
    End If
    
    Set tags = Nothing
    Set part = Nothing
    Set nodes = Nothing
    Set rootNode = Nothing

    Application.ScreenUpdating = True

End Sub


'===============================  MAIN METHODS  ================================

Sub ResetKey()

    Dim parts As CustomXMLParts
    Dim part As CustomXMLPart
    Dim nodes As CustomXMLNodes
    Dim node As CustomXMLNode
    Dim rootNode As CustomXMLNode
    
    Set parts = ThisDocument.CustomXMLParts
    If parts.Count > 0 Then
        
        Set part = parts("https://www.musicalnotes.com/")
        Set rootNode = part.DocumentElement
        For Each node In rootNode.ChildNodes
            bName = node.BaseName
            If Len(bName) = 2 Then
                Select Case Right(bName, 1)
                    Case Is = "s"
                        node.Text = UCase(Left(node.BaseName, 1)) & "#"
                    Case Is = "f"
                        node.Text = UCase(Left(node.BaseName, 1)) & "b"
                    Case Else
                End Select
            Else
                node.Text = UCase(node.BaseName)
            End If
        Next
    End If

End Sub


