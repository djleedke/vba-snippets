Attribute VB_Name = "EmailHandling"

Function SendEmail(ByVal toEmail As String, Optional ByVal subject As String, Optional ByVal attachmentPath)

    '   Opens and displays an email and fills with the provided parameters, does not send.
    '
    '   Arguments:
    '       toEmail: The to email address.
    '       subject: The subject of the email address
    '       attachmentPath: The file path of a single attachment to be added to the email.

    Dim outlook As Object
    Dim outlookMailItem As Object
    Dim attachments As Object

    Set outlook = CreateObject("outlook.application")
    Set outlookMailItem = outlook.createitem(0)
    Set attachments = outlookMailItem.attachments
    
    With outlookMailItem
    
        .To = toEmail
        .Body = ""
        
        If IsMissing(attachmentPath) = False Then
            attachments.Add attachmentPath
        End If
        
        If IsMissing(subject) = False Then
            .subject = subject
        End If

        .Display
    
    End With
    
    Set outlookMailItem = Nothing
    Set outlook = Nothing

End Function

