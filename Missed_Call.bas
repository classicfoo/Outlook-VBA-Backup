Attribute VB_Name = "Missed_Call"
Public Sub Forward_To_Self_As_Plain_Text()

    'loop though inbox items
    Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
    Dim Item As Object
    
    For Each Item In olFolder.Items
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            
            'check if the item is a missed call
            If oMail.subject = "Missed Call" Then
            
                'forward to self
                Dim oForward As MailItem
                Set oForward = oMail.Forward
                
                'set format to plain text
                oForward.BodyFormat = olFormatPlain
                oForward.Body = oMail.Body
                oForward.Recipients.Add "michael_huynh2@dell.com"
                oForward.Save
                oForward.Send
                
                'Delete the missed call
                oMail.Delete
            End If
        End If
    Next

End Sub
