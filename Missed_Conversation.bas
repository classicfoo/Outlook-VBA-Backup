Attribute VB_Name = "Missed_Conversation"
Public Sub Forward_To_Self_As_Plain_Text()

    'loop though inbox items
    Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
    Dim Item As Object
    
    For Each Item In olFolder.Items
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            
            'check if the item is a missed conversation;
            'by checking if subject contains "Conversation with"
            'once message is forwarded it will have "[PlainText]" appended to subject
            'make that the exception
            If InStr(oMail.subject, "Conversation with") > 0 & InStr(oMail.subject, " [PlainText]") < 0 Then
            
                'forward to self
                Dim oForward As MailItem
                Set oForward = oMail.Forward
                
                'set format to plain text
                oForward.BodyFormat = olFormatPlain
                oForward.Body = oMail.Body

                oForward.Recipients.Add olFolder.Session.CurrentUser
                oForward.subject = oForward.subject & " [PlainText]"
                oForward.Save
                oForward.Send
                
                'Delete the missed conversation
                oMail.Delete
           
            End If
            
        End If
    Next

End Sub

