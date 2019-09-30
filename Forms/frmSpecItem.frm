VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpecItem 
   Caption         =   "技术要求条目列表"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10125
   OleObjectBlob   =   "frmSpecItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpecItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbConn As ADODB.Connection
Private dbRs As ADODB.Recordset
Private dbRsD As ADODB.Recordset
Private dbPath As String

Private Sub CommandButton1_Click()
If tList.ListIndex = -1 Then Exit Sub
frmSpec.tView.AddItem
frmSpec.tView.list(frmSpec.tView.ListCount - 1, 0) = tList.list(tList.ListIndex, 0)
frmSpec.tView.list(frmSpec.tView.ListCount - 1, 1) = tList.list(tList.ListIndex, 1)

frmSpec.tView.Selected(frmSpec.tView.ListCount - 1) = True
frmSpec.tView.ListIndex = frmSpec.tView.ListCount - 1
Unload Me
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
Dim iDesp As String
iDesp = InputBox("请输入新项目描述:", "添加新项目")
If Len(iDesp) = 0 Then Exit Sub

Dim sql As String
sql = "INSERT INTO SE_SPEC_ITEM (sTxt) VALUES ('" & iDesp & "')"

dbConn.Execute sql
sql = "select * from SE_SPEC_ITEM where sTxt='" & iDesp & "'"
dbRs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly

tList.AddItem
tList.list(tList.ListCount - 1, 0) = iDesp
tList.list(tList.ListCount - 1, 1) = dbRs.Fields("ID")
tList.Selected(tList.ListCount - 1) = True

dbRs.Close

End Sub

Private Sub CommandButton4_Click()
If tList.ListIndex = -1 Then Exit Sub
If tList.list(tList.ListIndex, 1) = -1 Then Exit Sub    '新添加的条目要重启窗口后才能删除

Dim sql As String

sql = "DELETE * from SE_SPEC_ITEM where ID=" & tList.list(tList.ListIndex, 1)

dbConn.Execute sql

tList.RemoveItem tList.ListIndex
End Sub

Private Sub TextBox1_Change()
If Len(TextBox1.Text) = 0 Then
    tList.ListIndex = -1
    Exit Sub
Else
    Dim i As Integer
    For i = 0 To tList.ListCount - 1
        If InStr(1, tList.list(i, 0), TextBox1.Text) > 0 Then
            tList.ListIndex = i
            Exit For
        End If
    Next
End If
End Sub



Private Sub tList_Click()
'ddd
End Sub

Private Sub tList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call CommandButton1_Click
End Sub


Private Sub tList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If tList.ListIndex = -1 Then Exit Sub
If Button = 2 Then
    Dim Clipboard As New MSForms.DataObject
    Clipboard.SetText tList.list(tList.ListIndex)
    Clipboard.PutInClipboard
End If
End Sub

Private Sub UserForm_Initialize()
dbPath = GetSetting("Domisoft", "Config", "Spec_db_path", "")

Select Case LCase(Environ("UserName"))
    Case "ccl100100"
    Case Else
        CommandButton4.Visible = False
End Select

Set dbConn = New ADODB.Connection
Set dbRs = New ADODB.Recordset
Set dbRsD = New ADODB.Recordset

dbConn.Provider = "Microsoft.Jet.oledb.4.0"
dbConn.Open dbPath

Dim sql As String
sql = "select * from SE_SPEC_ITEM ORDER BY ID"

dbRs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly

Dim n As Integer
n = 0

    Do While (Not dbRs.EOF)

        tList.AddItem
        tList.list(n, 0) = dbRs.Fields("sTxt")
        tList.list(n, 1) = dbRs.Fields("ID")
        n = n + 1
        
        dbRs.MoveNext
    Loop
dbRs.Close
End Sub

Private Sub UserForm_Terminate()

dbConn.Close

Set dbRs = Nothing
Set dbRsD = Nothing
Set dbConn = Nothing

End Sub


