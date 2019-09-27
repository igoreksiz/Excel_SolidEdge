Public Sub openSelectedPdf()

Dim objMail As Outlook.MailItem
Set objMail = ThisOutlookSession.ActiveExplorer.Selection.Item(1)
Dim wordSelected As String
wordSelected = objMail.GetInspector.WordEditor.Application.Selection.Text
wordSelected = Trim(wordSelected)
wordSelected = Replace(Replace(wordSelected, Chr(10), ""), Chr(13), "")             '去除换行符

Dim myPDFstore As String

Dim p As Variant
p = Split(GetSetting("Domisoft", "Config", "PDF_Store", ""), "|")

Dim filename As String
Dim done As Boolean
done = False
For i = LBound(p) To UBound(p)

    myPDFstore = p(i)
    
    filename = wordSelected
    
    If InStr(1, filename, Chr(10), vbTextCompare) > 0 Then
        filename = Split(filename, Chr(10))(0)           ' TODO 一格里含有多个文件名
    End If
    If Len(filename) = 8 And Left(filename, 1) = 8 Then filename = "00" & filename    '解决00问题
    
    filename = myPDFstore & "\" & filename & ".pdf"
    
    If IsFileExists(filename) Then
        Shell "explorer.exe " & filename
        done = True
        Exit For
    End If
Next
If Not done Then MsgBox "file not found: " & filename, vbOKOnly, "File Not Found"
End Sub
Public Function IsFileExists(ByVal strFileName As String) As Boolean
    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

