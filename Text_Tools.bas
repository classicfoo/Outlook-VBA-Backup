Attribute VB_Name = "Text_Tools"
Public Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Public Function GetTextBetween(alltext, delimiter1, delimiter2) As String
    alltext = Split(alltext, delimiter1)(1)
    alltext = Split(alltext, delimiter2)(0)
    GetTextBetween = Trim(alltext)
End Function

Public Function GetTwoLinesBefore(alltext, delimiter1) As String
MsgBox (alltext)
    alltext = Split(alltext, delimiter1)(0)
    GetTwoLinesBefore = Trim(alltext)
End Function


Public Function ReadTextFile(path) As String

    Dim FileNum As Integer
    Dim DataLine As String
    Dim alltext As String
    
    FileNum = FreeFile()
    Open path For Input As #FileNum
    
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine ' read in data 1 line at a time
        ' decide what to do with dataline,
        ' depending on what processing you need to do for each case
        alltext = alltext + DataLine
    Wend
    
    ReadTextFile = alltext
End Function

