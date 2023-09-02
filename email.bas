Attribute VB_Name = "emails"

Public Sub SendTest()

    Dim outlook_app As New Outlook.Application
    Dim outlook_mail As Outlook.MailItem
    
    Set outlook_mail = outlook_app.CreateItem(olMailItem)
    
    With outlook_mail
        .BodyFormat = olFormatHTML
        .Display
        .HTMLBody = "Send email from VBA using Outlook. <br><br>Test conducted at: " & Format(Now(), "yyyy/mm/dd hh:mm:ss") & .HTMLBody 'Need to do this to add the default signature (also, must use .Display to generate it)
        .To = "someone@gmail.com; someone@hotmail.com"
        .Cc = "someone_else@outlook.com"
        .BCC  = "still_another_person@yahoo.com"
        .Subject = "Test of Automated Emails Sent from VBA"
        .Importance = olImportanceHigh
        .Attachments.Add ThisWorkbook.FullName
        .Send
    End With

    Set outlook_app = Nothing
    Set outlook_mail = Nothing

End Sub
