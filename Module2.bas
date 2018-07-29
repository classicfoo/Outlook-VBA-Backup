Attribute VB_Name = "Module2"

Sub ForceForwardInHTML()

'=================================================================
'Description: Outlook macro to Forward to a message in HTML
'             regardless of the current message format.
'             The forward will use your HTML signature as well.
'
'author : Robert Sparnaaij
'version: 1.0
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




