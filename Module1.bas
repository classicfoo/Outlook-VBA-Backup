Attribute VB_Name = "Module1"

Sub EUC()

'website: https://www.howto-outlook.com/howto/replyinhtml.htm
'=================================================================
    
    Dim objOL As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Set objOL = Outlook.Application
    
    'Get the selected item
    Select Case TypeName(objOL.ActiveWindow)
        Case "Explorer"
            Set objSelection = objOL.ActiveExplorer.Selection
            If objSelection.Count > 0 Then
                Set objItem = objSelection.Item(1)
            Else
                Result = MsgBox("No item selected. " & _
                            "Please make a selection first.", _
                            vbCritical, "Forward in HTML")
                Exit Sub
            End If
        
        Case "Inspector"
            Set objItem = objOL.ActiveInspector.CurrentItem
            
        Case Else
            Result = MsgBox("Unsupported Window type." & _
                        vbNewLine & "Please make a selection" & _
                        " or open an item first.", _
                        vbCritical, "Forward in HTML")
            Exit Sub
    End Select

    Dim olMsg As Outlook.MailItem
    Dim olMsgForward As Outlook.MailItem
    Dim IsPlainText As Boolean
    
    'Change the message format and forward
    If objItem.Class = olMail Then
        Set olMsg = objItem
        If olMsg.BodyFormat = olFormatPlain Then
            IsPlainText = True
        End If
        olMsg.BodyFormat = olFormatHTML
        Set olMsgForward = olMsg.Forward
        If IsPlainText = True Then
            olMsg.BodyFormat = olFormatPlain
        End If
        
        'create message signature
        '*********
        Dim strBuffer As String
        enviro = CStr(Environ("appdata"))
        Debug.Print enviro
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        ' Edit the signature file name on the following line
        strSigFilePath = enviro & "\Microsoft\Signatures\"
        Debug.Print strSigFilePath
        Set objSignatureFile = objFSO.OpenTextFile(strSigFilePath & "ORL QUOPP.htm")
        strBuffer = objSignatureFile.ReadAll
        objSignatureFile.Close
        
        With olMsgForward
            '.Subject = "Subject goes here"
            .SendUsingAccount = Application.GetNamespace("MAPI").accounts.Item(1) 'Macro using the Default Account 'https://www.slipstick.com/developer/send-using-default-or-specific-account/
            .To = "ORL_OP_AU_QUOPP <ORL_OP_AU_QUOPP@Dell.com>"
            .CC = "David Brown (Brisbane) (David_Brown1@Dell.com); Brookes, Joel (Joel_Brookes@Dell.com)"
            
            Dim accountnum As String
            Dim accountname As String
            accountname = Eqorder_Subject_Parser.GetAccountName(.subject)
            accountnum = Eqorder_Subject_Parser.GetCustomerNumber(.subject)
            
            If accountnum = "" Then
                MsgBox "No Account Found"
                accountnum = "XXXXXXXXXX"
            End If
                        
            .HTMLBody = "Customer number: " & accountnum & strBuffer & olMsgForward.HTMLBody
            .Display
            
            'Resolve each Recipient's name.
             For Each objOutlookRecip In olMsgForward.Recipients
               objOutlookRecip.Resolve
             Next
           
          
        End With
        '*********
        
        
        olMsg.Close (olSave)
        olMsgForward.Display
        
    'Selected item isn't a mail item
    Else
        Result = MsgBox("No message item selected. " & _
                    "Please make a selection first.", _
                    vbCritical, "Forward in HTML")
        Exit Sub
    End If
    
    'Cleanup
    Set objOL = Nothing
    Set objItem = Nothing
    Set objSelection = Nothing
    Set olMsg = Nothing
    Set olMsgForward = Nothing
       
End Sub

Sub ISG()

'website: https://www.howto-outlook.com/howto/replyinhtml.htm
'=================================================================
    
    Dim objOL As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Set objOL = Outlook.Application
    
    'Get the selected item
    Select Case TypeName(objOL.ActiveWindow)
        Case "Explorer"
            Set objSelection = objOL.ActiveExplorer.Selection
            If objSelection.Count > 0 Then
                Set objItem = objSelection.Item(1)
            Else
                Result = MsgBox("No item selected. " & _
                            "Please make a selection first.", _
                            vbCritical, "Forward in HTML")
                Exit Sub
            End If
        
        Case "Inspector"
            Set objItem = objOL.ActiveInspector.CurrentItem
            
        Case Else
            Result = MsgBox("Unsupported Window type." & _
                        vbNewLine & "Please make a selection" & _
                        " or open an item first.", _
                        vbCritical, "Forward in HTML")
            Exit Sub
    End Select

    Dim olMsg As Outlook.MailItem
    Dim olMsgForward As Outlook.MailItem
    Dim IsPlainText As Boolean
    
    'Change the message format and forward
    If objItem.Class = olMail Then
        Set olMsg = objItem
        If olMsg.BodyFormat = olFormatPlain Then
            IsPlainText = True
        End If
        olMsg.BodyFormat = olFormatHTML
        Set olMsgForward = olMsg.Forward
        If IsPlainText = True Then
            olMsg.BodyFormat = olFormatPlain
        End If
        
        'create message signature
        '*********
        Dim strBuffer As String
        enviro = CStr(Environ("appdata"))
        Debug.Print enviro
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        ' Edit the signature file name on the following line
        strSigFilePath = enviro & "\Microsoft\Signatures\"
        Debug.Print strSigFilePath
        Set objSignatureFile = objFSO.OpenTextFile(strSigFilePath & "ORL OFFLINE.htm")
        strBuffer = objSignatureFile.ReadAll
        objSignatureFile.Close
        
        With olMsgForward
          '.Subject = "Subject goes here"
          .SendUsingAccount = Application.GetNamespace("MAPI").accounts.Item(1) 'Macro using the Default Account 'https://www.slipstick.com/developer/send-using-default-or-specific-account/
          .To = "ORL_OP_AU_QLD <ORL_OP_AU_QLD@Dell.com>"
          .CC = "David Brown (Brisbane) (David_Brown1@Dell.com); Brookes, Joel (Joel_Brookes@Dell.com); Arifin, Andry <Andry_Arifin@Dell.com>"
          .HTMLBody = strBuffer & olMsgForward.HTMLBody
          .Display
          
          'Resolve each Recipient's name.
           For Each objOutlookRecip In olMsgForward.Recipients
             objOutlookRecip.Resolve
           Next
           
          
        End With
        '*********
        
        
        olMsg.Close (olSave)
        olMsgForward.Display
        
    'Selected item isn't a mail item
    Else
        Result = MsgBox("No message item selected. " & _
                    "Please make a selection first.", _
                    vbCritical, "Forward in HTML")
        Exit Sub
    End If
    
    'Cleanup
    Set objOL = Nothing
    Set objItem = Nothing
    Set objSelection = Nothing
    Set olMsg = Nothing
    Set olMsgForward = Nothing
       
End Sub


