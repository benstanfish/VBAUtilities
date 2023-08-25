Attribute VB_Name = "reflection"

'To use the VBProject object, set a reference to
'Microsoft Visual Basic for Applications Extensibility 5.3 (the VBIDE library).
'as well as Microsoft VBScript Regular Expressions 5.5

Public Enum proc_kind
    vbext_pk_Proc = 0
    vbext_pk_Let = 1
    vbext_pk_Set = 2
    vbext_pk_Get = 3
End Enum

Public Function proc_kind_text(proc_kind As Long)

    Dim arr As Variant
    arr = Array("vbext_pk_Proc", "vbext_pk_Let", "vbext_pk_Set", "vbext_pk_Get")
    proc_kind_text = arr(proc_kind)

End Function

Private Sub count_module_lines()

    Dim vb_project As VBIDE.VBProject
    Dim vb_component As VBIDE.VBComponent
    Dim code_module As VBIDE.CodeModule
    
    Set vb_project = ActiveWorkbook.VBProject
    Set vb_component = vb_project.VBComponents("colorfuncs")
    Set code_module = vb_component.CodeModule
    
    Debug.Print code_module.CountOfLines

End Sub

Private Sub get_dec_lines()

    Dim vb_project As VBIDE.VBProject
    Dim vb_component As VBIDE.VBComponent
    Dim code_module As VBIDE.CodeModule
    
    Set vb_project = ActiveWorkbook.VBProject
    Set vb_component = vb_project.VBComponents("colorfuncs")
    Set code_module = vb_component.CodeModule
    
    For i = 1 To code_module.CountOfDeclarationLines
        Debug.Print i; vbTab; code_module.Lines(i, 1)
    Next

End Sub

Private Sub list_all_proc_lines()

    Dim vb_project As VBIDE.VBProject
    Dim vb_component As VBIDE.VBComponent
    Dim code_module As VBIDE.CodeModule
    
    Set vb_project = ActiveWorkbook.VBProject
    Set vb_component = vb_project.VBComponents("colorfuncs")
    Set code_module = vb_component.CodeModule
    
    With code_module
        line_number = .CountOfDeclarationLines + 1
        Do Until line_number >= .CountOfLines
            proc_name = .ProcOfLine(line_number, vbext_pk_Proc)

            Debug.Print line_number
            'skip over the rest of the current procedure's lines
            line_number = .ProcStartLine(proc_name, vbext_pk_Proc) + _
                .ProcCountLines(proc_name, vbext_pk_Proc) + 1
        Loop
    End With

End Sub

Private Sub list_all_procedures()

    Dim vb_project As VBIDE.VBProject
    Dim vb_component As VBIDE.VBComponent
    Dim code_module As VBIDE.CodeModule
    Dim count_string As String * 5
    Dim i As Long
    
    Set vb_project = ActiveWorkbook.VBProject
    Set vb_component = vb_project.VBComponents("colorfuncs")
    Set code_module = vb_component.CodeModule
    
    i = 1
    With code_module
        line_number = .CountOfDeclarationLines + 1
        Do Until line_number >= .CountOfLines
            proc_name = .ProcOfLine(line_number, vbext_pk_Proc)
            count_string = CStr(i)
            Debug.Print count_string; proc_name & "()"
            i = i + 1
            'skip over the rest of the current procedure's lines
            line_number = .ProcStartLine(proc_name, vbext_pk_Proc) + _
                .ProcCountLines(proc_name, vbext_pk_Proc) + 1
        Loop
    End With

End Sub

Private Sub list_arguments()

    Dim vb_project As VBIDE.VBProject
    Dim vb_component As VBIDE.VBComponent
    Dim code_module As VBIDE.CodeModule
    Dim count_string As String * 5
    Dim i As Long
    
    Set vb_project = ActiveWorkbook.VBProject
    Set vb_component = vb_project.VBComponents("colorfuncs")
    Set code_module = vb_component.CodeModule
    
    Dim reg_string As String
    Dim regex As Object
    Set regex = New RegExp
    
    reg_string = "\(([^)]+)?\)"
    regex.Pattern = reg_string
    
    i = 1
    With code_module
        line_number = .CountOfDeclarationLines + 1
        Do Until line_number >= .CountOfLines
            proc_name = .ProcOfLine(line_number, vbext_pk_Proc)
            'TODO: may be missing arguments from lines with returns
            Set matches = regex.Execute(.Lines(line_number, 1))
            For Each Match In matches
                Debug.Print line_number; vbTab; Replace(Replace(Match, "(", ""), ")", "")
            Next
            i = i + 1
            'skip over the rest of the current procedure's lines
            line_number = .ProcStartLine(proc_name, vbext_pk_Proc) + _
                .ProcCountLines(proc_name, vbext_pk_Proc) + 1
        Loop
    End With

End Sub


Private Sub ExportModules()
    Dim me_path As String
    Dim comp As VBIDE.VBComponent
    
    me_path = Application.ActiveWorkbook.Path & "\"
    
    For Each comp In ActiveWorkbook.VBProject.VBComponents
        is_export = True
        Select Case comp.Type
            Case vbext_ct_ClassModule
                comp.Export me_path & comp.Name & ".cls"
            Case vbext_ct_MSForm
                comp.Export me_path & comp.Name & ".frm"
            Case vbext_ct_StdModule
                comp.Export me_path & comp.Name & ".bas"
            Case vbext_ct_Document
                ' Don't export
        End Select
    Next
End Sub