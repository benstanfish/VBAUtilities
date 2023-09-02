Attribute VB_Name = "emails"

Public Sub SendTest()

    Dim outlook_app As New Outlook.Application
    Dim outlook_mail As Outlook.MailItem
    
    Set outlook_mail = outlook_app.CreateItem(olMailItem)
    
    With outlook_mail
        .BodyFormat = olFormatHTML
        .HTMLBody = "Send email from VBA using Outlook. <br><br>Test conducted at: " & Format(Now(), "yyyy/mm/dd hh:mm:ss")
        .To = "someone@gmail.com; someone@hotmail.com"
        .Cc = "someone_else@outlook.com"
        .BCC  = "still_another_person@yahoo.com"
        .Subject = "Test of Automated Emails Sent from VBA"
        .Importance = olImportanceHigh
        ' Add .Attachements to add attachements
        ' Use .Send to silently email, or use .Display to show the email draft
        .Send
    End With

    Set outlook_app = Nothing
    Set outlook_mail = Nothing

End Sub
