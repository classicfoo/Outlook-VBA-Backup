Attribute VB_Name = "Numbers"
Public Function GetTotal(path) As String

    alltext = Text_Tools.ReadTextFile(path)
    
    d1 = "Total incl. GST: AUD"
    d2 = "Unless otherwise stated"
    
    GetTotal = Text_Tools.GetTwoLinesBefore(alltext, d1)

End Function

Sub total()
    MsgBox GetTotal("C:\Users\michael_huynh2\Desktop\Purchase Order 2063 2002099.txt")
End Sub


Sub saveattachments()
    Call Mail.saveattachments
    mydocs = CreateObject("WScript.Shell").SpecialFolders(16)
    Dim strFolderpath As Variant
    
    f = Dir(mydocs & "\OLAttachments\")
    strFolderpath = mydocs & "\OLAttachments\"
    
    While f <> ""
        If InStr(f, ".txt") Then
            MsgBox GetTotal(strFolderpath & f)
        End If

        'MsgBox GetTotal(strFolderpath)
        'Set the fileName to the next file
        f = Dir
    Wend
End Sub


Sub DisplayCurrentUser()

 Dim myNamespace As Outlook.NameSpace

 Set myNamespace = Application.GetNamespace("MAPI")

 MsgBox myNamespace.CurrentUser

End Sub
