VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'on item send event handler
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
Dim objRecip As Recipient
    Dim strMsg As String
    Dim res As Integer
    Dim strBcc As String
    On Error Resume Next
     
    strBcc = "emailtosalesforce@e-259u20awl32532bg7qxmb13ykmbmdmojif1vu0o50z8rp9kre3.3-6urqeaa.na7.le.salesforce.com"
    
    'if "To:" field contains "@eq.edu.au"
    'case insensitive
    
    If InStr(1, Item.To, "@eq.edu.au") Then
        
        'Testing
        'MsgBox "contains @eq.edu.au"
        
        Set objRecip = Item.Recipients.Add(strBcc)
        objRecip.Type = olBCC
        If Not objRecip.Resolve Then
            strMsg = "Could not resolve the Bcc recipient. " & _
                     "Do you want still to send the message?"
            res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, _
                     "Could Not Resolve Bcc Recipient")
            If res = vbNo Then
                Cancel = True
            End If
        End If
         
        Set objRecip = Nothing
    
    End If
    
    

End Sub

Private Sub test()
    'Item.To

    x = InStr(1, "asfsadfafassdf@eq.edu.au", "@eq.edu.au")
    
    If x Then
        MsgBox "is eq"
    End If
    
    x = "asdf"
    
End Sub




'on new email event handler
Private Sub Application_NewMail()

    Call Missed_Call.Forward_To_Self_As_Plain_Text
    Call Missed_Conversation.Forward_To_Self_As_Plain_Text
    
    'subject = Outlook.Application.ActiveExplorer.Selection.Item(1).subject
    'Call Polog.AddRow(Date, subject, "", Polog.
    
End Sub
