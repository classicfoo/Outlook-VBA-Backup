Attribute VB_Name = "Polog"
Public Function GetTotal(path) As String

    alltext = Text_Tools.ReadTextFile(path)
    
    d1 = "Total incl. GST: AUD"
    d2 = "Unless otherwise stated"
    
    GetTotal = Text_Tools.GetTextBetween(alltext, d1, d2)

End Function

Sub total()
    MsgBox GetTotal("C:\Users\michael_huynh2\Desktop\Purchase Order 2063 2002099.txt")
End Sub

'https://stackoverflow.com/questions/35769065/email-from-outlook-with-excel-cell-value-in-subject
Public Function AddRow(subject, exgst) As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    Dim sPath As String

    sPath = "C:\Users\michael_huynh2\Documents\DET\DET_May_2018_Polog.xlsx" '<- Update"

'   // Excel
    Set xlApp = CreateObject("Excel.Application")
'   // Workbook
    Set xlBook = xlApp.Workbooks.Open(sPath)
'   // Sheet
    Set xlSht = xlBook.Sheets("Auto")

    'https://stackoverflow.com/questions/8295276/function-or-sub-to-add-new-row-and-data-to-table
    'Public Sub addDataToTable(ByVal strTableName As String, ByVal strData As String, ByVal col As Integer)
    strTableName = "Table1"
    Dim lLastRow As Long
    Dim iHeader As Integer

    With xlSht.ListObjects(strTableName)
        'find the last row of the list
        lLastRow = xlSht.ListObjects(strTableName).ListRows.Count
        'shift from an extra row if list has header
        If .Sort.Header = xlYes Then
            iHeader = 1
        Else
            iHeader = 0
        End If
    End With
    'add the data a row after the end of the list
    xlSht.Cells(lLastRow + 1 + iHeader, 1).Value = Date 'add today's date
    xlSht.Cells(lLastRow + 1 + iHeader, 2).Value = "PO num" 'add today's date
    

'   // Close
    xlBook.Close SaveChanges:=True
'   // Quit
    xlApp.Quit

    '// CleanUp
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSht = Nothing

End Function

