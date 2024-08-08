Public Sub ScrapeAttachmentsFromEmail()
    ' Created by Ben Fisher, 8 July 2024
    ' Saves all inline attachements (images, etc.) to the downloads folder

    Dim aMail As Outlook.MailItem, aInspector As Inspector
    Set aInspector = Application.ActiveInspector
   
    ' You must have an email opened in an inspector window.
    If aInspector Is Nothing Then
        Debug.Print "No active inspector"
    Else
        Set aMail = aInspector.CurrentItem
        Dim oAtt As Outlook.Attachment, fPath As String
        ' This will save the attachments to the User's downloads folder
        For Each oAtt In aMail.Attachments
            fPath = Environ("USERPROFILE") & "\Downloads\" & oAtt.DisplayName
            oAtt.SaveAsFile Path:=fPath
        Next
    End If
 
End Sub


Public Sub MarkAllSubfoldersRead()


    ' Marks all mail items in subfolders of Inbox - not in the top-level of the inbox

 
    Dim oNameSpace As Outlook.NameSpace
    Set oNameSpace = Application.GetNamespace("MAPI")
  

    Dim myInbox As Outlook.Folder
    Set myInbox = oNameSpace.GetDefaultFolder(olFolderInbox)

   
    Dim aFolder As Outlook.Folder
    Dim aMessage As Outlook.MailItem
   

    Dim thing As Object
    For Each aFolder In myInbox.Folders
        For Each thing In aFolder.Items
            If thing.Class = olMail Then
                Set aMessage = thing
                If aMessage.UnRead Then aMessage.UnRead = False
            End If
        Next
    Next

 
    Set oNameSpace = Nothing
    Set myInbox = Nothing
    Set aFolder = Nothing
    Set aMessage = Nothing
    Set thing = Nothing

 
End Sub

 

 

Public Sub PreviewName()

 
    Dim aMail As Outlook.MailItem, aInspector As Inspector
    Set aInspector = Application.ActiveInspector

 
    If aInspector Is Nothing Then
        Debug.Print "No active inspector"
    Else
        Set aMail = aInspector.CurrentItem
        Dim oAtt As Outlook.Attachment, fPath As String
        ' This will save the attachments to the User's downloads folder
        For Each oAtt In aMail.Attachments
            'fPath = Environ("USERPROFILE") & "\Downloads\" & oAtt.FileName
            'oAtt.SaveAsFile Path:=fPath
            Debug.Print oAtt.DisplayName
        Next
    End If

 
End Sub
