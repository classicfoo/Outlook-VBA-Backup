Attribute VB_Name = "Eqorder_Subject_Parser"
Public Function GetAccountName(subject) As String

    Dim account As String
    Dim custnum As String
    Dim noaccount As Boolean
    
    noaccount = True
        
    'get the account name by splitting keywords
    account = subject
    
    'try parse the string
    On Error GoTo noaccount
    account = Split(account, "from")(1)
    account = Split(account, "- (")(0)
    account = Trim(account)
    noaccount = False
    
    'on error set account name to nothing
noaccount:
    If noaccount = True Then
        account = ""
    End If
    
    'return the account name
    GetAccountName = account
    
End Function

Public Function GetCustomerNumber(subject) As String

    Dim account As String
    account = GetAccountName(subject)
    
    If account = "" Then
        Exit Function
    End If
            
    
    'read_whole_file()
    Dim file As String, sWhole As String
    Dim v As Variant
    file = "C:\Users\michael_huynh2\Documents\DET\DET_May_2018_Accounts_With_Omega_Customer_Numbers.csv"
    Open file For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine)
    
    Dim accountname As String 'the account name in our list
    Dim accountnum As String 'the account number in our list
    
    On Error GoTo custnumnotfound
    For Each element In v
        accountname = Split(element, ",")(0)
                
        If (accountname = account) Then
            accountnum = Split(element, ",")(1)
            Exit For
        End If
        Debug.Print (element)
    Next element

    GetCustomerNumber = accountnum
    Exit Function
    
custnumnotfound:
    GetCustomerNumber = ""
    
End Function


Public Sub test_GetAccountName()

    Dim subject As String
    subject = "Purchase Order 2000043 from Albert State School - (0038)"
    
    MsgBox GetAccountName(subject)
    
End Sub

Public Sub test_GetCustomerNumber()

    MsgBox GetCustomerNumber("Purchase Order 2000043 from Albert State School - (0038)")
    
End Sub






