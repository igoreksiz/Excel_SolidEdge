VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpec 
   Caption         =   "技术要求编辑器"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9390
   OleObjectBlob   =   "frmSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const KEY_PRESSING As Integer = &H8000
Private Const KEY_PRESSED As Integer = &H1000
Private Const VK_CONTROL As Long = &H11

Private dbConn As ADODB.Connection
Private dbRs As ADODB.Recordset
Private dbRsD As ADODB.Recordset
Dim dbPath As String

Dim dft As SolidEdgeDraft.DraftDocument


Private Sub cb_add_Click()
frmSpecItem.Show

'Dim ss As String
'ss = InputBox("输入自定义技术要求", "input")
'tView.AddItem ss
'tView.Selected(tView.ListCount - 1) = True
'tView.ListIndex = tView.ListCount - 1

End Sub

Private Sub cb_del_Click()
If tView.ListIndex >= 0 Then
    tView.RemoveItem tView.ListIndex
End If
End Sub

Private Sub cb_down_Click()
If tView.ListIndex = -1 Then Exit Sub
Dim ss As Variant
ReDim ss(0 To 1) As String
Dim nn As Integer
Dim ckd As Boolean
With tView
    If .ListIndex = .ListCount - 1 Then Exit Sub
    ss(0) = .list(.ListIndex, 0)
    ss(1) = .list(.ListIndex, 1)
    nn = .ListIndex
    ckd = .Selected(.ListIndex)
    .RemoveItem nn
    .AddItem ss(0), nn + 1
    .list(nn + 1, 1) = ss(1)
    .Selected(nn + 1) = ckd
    .ListIndex = nn + 1
End With
End Sub
Private Sub cb_up_Click()
If tView.ListIndex = -1 Then Exit Sub
Dim ss As Variant
ReDim ss(0 To 1) As String
Dim nn As Integer
Dim ckd As Boolean
With tView
    If .ListIndex = 0 Then Exit Sub
    ss(0) = .list(.ListIndex, 0)
    ss(1) = .list(.ListIndex, 1)
    nn = .ListIndex
    ckd = .Selected(.ListIndex)
    .RemoveItem nn
    .AddItem ss(0), nn - 1
    .list(nn - 1, 1) = ss(1)
    .Selected(nn - 1) = ckd
    .ListIndex = nn - 1
End With
End Sub
Private Sub cb_ok_Click()
If tView.ListCount = 0 Then Exit Sub

Dim sht As SolidEdgeDraft.sheet
Set sht = dft.ActiveSheet

Dim txts As SolidEdgeFrameworkSupport.TextBoxes
Set txts = sht.TextBoxes

Dim txt As SolidEdgeFrameworkSupport.TextBox
Set txt = txts.Add(0.25, 0.12, 0)

Dim tEdit As SolidEdgeFrameworkSupport.TextEdit

With txt
    .HorizontalAlignment = igTextHzAlignLeft
    .FlowOrientation = igTextHorizontal
    .Justification = igTextJustifyTop
    .TextControlType = igTextFitToContent
    .Text = "技术要求:" & vbNewLine
    .Text = .Text & " " & vbNewLine
    
    Dim i As Integer
    Dim n As Integer
    n = 1
    For i = 0 To tView.ListCount - 1
        If tView.Selected(i) = True Then
            .Text = .Text & n & ". " & tView.list(i) & vbNewLine
            n = n + 1
        End If
    Next
    
'备用代码,目前有bug
'    Set tEdit = .Edit'
'    tEdit.SetSelect 8, Len(.Text), seTextSelectParagraph
'    tEdit.SetNumberList igPlain, igNoFormat, igRightJustification
    
End With

'dft.SelectSet.RemoveAll
'dft.SelectSet.Add txt

txt.Cut

AppActivate seApp.Name

Call SendPaste

Unload Me
End Sub

Private Function ReplacePara(str As String)
Dim outstr As String

'If dft.ModelLinks.Item(1).ModelDocument.Type <> igSheetMetalDocument Then
'    ReplacePara = str
''    Debug.Print "not sheet"
'    Exit Function
'End If

If InStr(1, str, "{thk}", vbTextCompare) < 1 And InStr(1, str, "{radii}", vbTextCompare) < 1 Then
    ReplacePara = str
    Exit Function
End If


Dim smdl As SolidEdgePart.SheetMetalDocument
Dim i As Integer
For i = 1 To dft.ModelLinks.Count
    If dft.ModelLinks.Item(i).ModelDocument.Type = igSheetMetalDocument Then
        Set smdl = dft.ModelLinks.Item(i).ModelDocument
        Exit For
    End If
Next i

If smdl Is Nothing Then
    ReplacePara = outstr
    Exit Function
End If

Dim thk As Variant
Dim radii As Variant
Call smdl.GetGlobalParameter(seSheetMetalGlobalMaterialThickness, thk)
Call smdl.GetGlobalParameter(seSheetMetalGlobalBendRadius, radii)


outstr = str
outstr = Replace(outstr, "{thk}", Format(CDbl(thk) * 1000, "#0.0;\E\r\r"), , , vbTextCompare)
outstr = Replace(outstr, "{radii}", Format(CDbl(radii) * 1000, "#0.0;\E\r\r"), , , vbTextCompare)

ReplacePara = outstr

End Function
Private Sub SpecSaveNew_Click()
Dim overwrite As Boolean
overwrite = False

If GetKeyState(VK_CONTROL) And KEY_PRESSING Then
    If LCase(Environ("UserName")) = "ccl100100" Then
        overwrite = True
    End If
End If

If tView.ListCount = 0 Then Exit Sub
Dim combSTR As Variant
Dim selSTR As Variant
ReDim combSTR(1 To tView.ListCount)
ReDim selSTR(1 To tView.ListCount)

Dim i As Integer
For i = 1 To tView.ListCount
combSTR(i) = tView.list(i - 1, 1)
selSTR(i) = IIf(tView.Selected(i - 1), 1, 0)
Next

Dim tName As String
Dim sql As String

If overwrite = True Then
    tName = tList.list(tList.ListIndex)
    sql = "select * from SE_SPEC_TEMPLATE where tName='" & tName
Else
    tName = InputBox("输入模版名称:", "名称")
    If Len(tName) = 0 Then
        Exit Sub
    End If
        
    For i = 0 To tList.ListCount - 1
        If tName = tList.list(i) Then
            MsgBox "此名称已经存在!", vbCritical, "保存失败"
    '        Exit For
            Exit Sub
        End If
    Next
    sql = "select * from SE_SPEC_TEMPLATE where 1=2"
End If

dbRs.Open sql, dbConn, adOpenForwardOnly, adLockOptimistic
If overwrite = False Then
    dbRs.AddNew
    dbRs.Fields("tName") = tName
End If
    dbRs.Fields("tCombine") = Join(combSTR, ",")
    dbRs.Fields("tSelect") = Join(selSTR, ",")
dbRs.MoveFirst  '必须的.否则close时出错
dbRs.Close

'MsgBox "OK"
If overwrite = False Then
    tList.AddItem tName
    tList.Selected(tList.ListCount - 1) = True
End If
End Sub

Private Sub tList_Click()
tView.Clear

Dim sql As String
sql = "select tCombine,tSelect from SE_SPEC_TEMPLATE where tName='" & tList.list(tList.ListIndex) & "'"

dbRs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly
    Dim v As Variant
    Dim vs As Variant
    
     v = Split(dbRs.Fields("tCombine"), ",")
     vs = Split(dbRs.Fields("tSelect"), ",")
dbRs.Close

Dim i As Integer
Dim qSql As String

For i = LBound(v) To UBound(v)
    qSql = "select sTxt from SE_SPEC_ITEM where ID=" & v(i)
    dbRsD.Open qSql, dbConn, adOpenForwardOnly, adLockReadOnly
        tView.AddItem ReplacePara(dbRsD.Fields("sTxt"))
        tView.Selected(i) = CBool(vs(i))
        tView.list(i, 1) = v(i)
    dbRsD.Close
Next
tView.SetFocus
End Sub

Private Sub cb_Del_temp_Click()
If tList.ListIndex = -1 Then Exit Sub
If MsgBox("删除此项?", vbOKCancel + vbInformation, "确认") <> vbOK Then Exit Sub

Dim sql As String
sql = "DELETE * from SE_SPEC_TEMPLATE where tName='" & tList.list(tList.ListIndex) & "'"

dbConn.Execute sql

tList.RemoveItem tList.ListIndex

End Sub

Private Sub UserForm_Initialize()
If seApp Is Nothing Then Call Conn2se

Select Case LCase(Environ("UserName"))
    Case "ccl100100"
    Case Else
        cb_Del_temp.Visible = False
'        SpecSaveNew.Visible = False
End Select

Set dft = seApp.ActiveDocument

dbPath = GetSetting("Domisoft", "Config", "Spec_db_path", "")

Set dbConn = New ADODB.Connection
Set dbRs = New ADODB.Recordset
Set dbRsD = New ADODB.Recordset

dbConn.Provider = "Microsoft.Jet.oledb.4.0"
dbConn.Open dbPath

Dim sql As String
sql = "select * from SE_SPEC_TEMPLATE ORDER BY ID"

dbRs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly
    Do While (Not dbRs.EOF)
        tList.AddItem dbRs.Fields("tName")
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
