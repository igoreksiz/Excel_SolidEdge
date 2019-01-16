VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSE_tbm 
   Caption         =   "Title Block Manager for SolidEdge"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "frmSE_tbm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSE_tbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private seDFT As DraftDocument
Private seBlk As SolidEdgeDraft.BlockOccurrence
Private seLbs As SolidEdgeDraft.BlockLabelOccurrences
Private seVerBlk As SolidEdgeDraft.BlockOccurrence
Private seVerLbs As SolidEdgeDraft.BlockLabelOccurrences
Private isLegacyDoc As Boolean

Private dbConn As ADODB.Connection
Private dbRs As ADODB.Recordset
Private dbRsD As ADODB.Recordset
Private SeDftBlockId As BlkId
Dim dbPath As String



Private Sub cb_approve_date_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_approve_date.Text = qDate(VBA.Date)
End Sub

Private Sub cb_design_date_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_design_date = qDate(VBA.Date)
End Sub
Private Sub cb_review_date_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_review_date = qDate(VBA.Date)
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
If isLegacyDoc Then
    Call writeToLegacyActiveDocument
Else
    Call writeToActiveDocument
End If
AppActivate seApp.Name
Unload Me
End Sub

Private Sub CommandButton3_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Designer", cb_designer.Text
End Sub

Private Sub CommandButton4_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Reviewer", cb_reviewer.Text
End Sub

Private Sub CommandButton5_Click()
SaveSetting "Domisoft", "TBM_SE", "Default_Approver", cb_approver.Text
End Sub
Private Sub designer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_designer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Designer", "")
End Sub



Private Sub reviewer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_reviewer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Reviewer", "")
End Sub
Private Sub approver_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
cb_approver.Text = GetSetting("Domisoft", "TBM_SE", "Default_Approver", "")
End Sub

Private Sub UserForm_Initialize()
If seApp Is Nothing Then Call Conn2se


Set seDFT = seApp.ActiveDocument

If seDFT.Sheets.Count = 0 Then Exit Sub
If seDFT.ActiveSheet.BlockOccurrences.Count = 0 Then Exit Sub

Dim i As Integer
For i = 1 To seDFT.ActiveSheet.BlockOccurrences.Count
    If seDFT.ActiveSheet.BlockOccurrences.Item(i).Block.Name = "Title" Then
        Set seBlk = seDFT.ActiveSheet.BlockOccurrences.Item(i)
        Set seLbs = seBlk.BlockLabelOccurrences
        isLegacyDoc = False
    End If
    If seDFT.ActiveSheet.BlockOccurrences.Item(i).Block.Name = "Title-SRDC_V1" Then
        Set seBlk = seDFT.ActiveSheet.BlockOccurrences.Item(i)
        Set seLbs = seBlk.BlockLabelOccurrences
        isLegacyDoc = True
    End If
    If seDFT.ActiveSheet.BlockOccurrences.Item(i).Block.Name = "SRDC_Ver" Then
        Set seVerBlk = seDFT.ActiveSheet.BlockOccurrences.Item(i)
        Set seVerLbs = seVerBlk.BlockLabelOccurrences
    End If
Next

For i = 1 To seLbs.Count
    Select Case seLbs.Item(i).Name
        Case "型号/项目名称"
            SeDftBlockId.Model = i
        Case "零件名称"
            SeDftBlockId.name_cn = i
        Case "专用号"
            SeDftBlockId.drw_no = i
        Case "材料"
            SeDftBlockId.material = i
        Case "钣厚"
            SeDftBlockId.thk = i
        Case "质量/体积"
            SeDftBlockId.weight = i
        Case "喷粉标准"
            SeDftBlockId.paint_std = i
        Case "公差等级"
            SeDftBlockId.tol = i
        Case "设计"
            SeDftBlockId.designer = i
        Case "审核"
            SeDftBlockId.reviewer = i
        Case "批准"
            SeDftBlockId.approver = i
        Case "设计日期"
            SeDftBlockId.design_date = i
        Case "审核日期"
            SeDftBlockId.review_date = i
        Case "批准日期"
            SeDftBlockId.approve_date = i
    End Select
Next

If isLegacyDoc Then
     Call getLegacyFromActive
Else
     Call getFromActive
End If

Call loadDefault
Call loadLists
seDocPath.Caption = seDFT.FullName

End Sub
Private Sub loadDefault()
If cb_designer.Text = "" Or cb_designer.Text = "-" Then cb_designer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Designer", "")
If cb_reviewer.Text = "" Or cb_reviewer.Text = "-" Then cb_reviewer.Text = GetSetting("Domisoft", "TBM_SE", "Default_Reviewer", "")
If cb_approver.Text = "" Or cb_approver.Text = "-" Then cb_approver.Text = GetSetting("Domisoft", "TBM_SE", "Default_Approver", "")
If cb_design_date.Text = "YYYY.MM.DD" Or cb_design_date.Text = "" Or cb_design_date.Text = "2016.11.19" Then cb_design_date.Text = qDate(VBA.Date)
If cb_review_date.Text = "YYYY.MM.DD" Or cb_review_date.Text = "" Or cb_review_date.Text = "2016.11.18" Then cb_review_date.Text = qDate(VBA.Date)
If cb_approve_date.Text = "YYYY.MM.DD" Or cb_approve_date.Text = "" Or cb_approve_date.Text = "2016.11.18" Then cb_approve_date.Text = qDate(VBA.Date)
End Sub
Private Sub loadLists()
dbPath = GetSetting("Domisoft", "Config", "Spec_db_path", "")
Set dbConn = New ADODB.Connection
Set dbRs = New ADODB.Recordset
Set dbRsD = New ADODB.Recordset
dbConn.Provider = "Microsoft.Jet.oledb.4.0"
dbConn.Open dbPath

Dim sql As String
sql = "select * from SE_TBM_Lists"

dbRs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly
    Do While (Not dbRs.EOF)
        Select Case dbRs.Fields("ListName")
            Case "name_cn"
                cb_name_cn.AddItem dbRs.Fields("Title")
            Case "model_no"
                cb_model_no.AddItem dbRs.Fields("Title")
            Case "material"
                cb_material.AddItem dbRs.Fields("Title")
            Case "weight"
                cb_weight.AddItem dbRs.Fields("Title")
            Case "designer"
                cb_designer.AddItem dbRs.Fields("Title")
            Case "reviewer"
                cb_reviewer.AddItem dbRs.Fields("Title")
            Case "approver"
                cb_approver.AddItem dbRs.Fields("Title")
        End Select
        dbRs.MoveNext
    Loop
dbRs.Close
dbConn.Close

Set dbRs = Nothing
Set dbRsD = Nothing
Set dbConn = Nothing
End Sub

Private Sub getFromActive()

cb_name_cn.Text = seLbs.Item(SeDftBlockId.name_cn).Value
cb_model_no.Text = seLbs.Item(SeDftBlockId.Model).Value
cb_material.Text = seLbs.Item(SeDftBlockId.material).Value
cb_weight.Text = seLbs.Item(SeDftBlockId.weight).Value
'version.Text = seLBS.Item(SeDftBlockId.version).Value
cb_designer.Text = seLbs.Item(SeDftBlockId.designer).Value
cb_design_date.Text = seLbs.Item(SeDftBlockId.design_date).Value
cb_reviewer.Text = seLbs.Item(SeDftBlockId.reviewer).Value
cb_review_date.Text = seLbs.Item(SeDftBlockId.review_date).Value
cb_approver.Text = seLbs.Item(SeDftBlockId.approver).Value
cb_approve_date.Text = seLbs.Item(SeDftBlockId.approve_date).Value
'add ver here
End Sub
Private Sub writeToActiveDocument()
seLbs.Item(SeDftBlockId.name_cn).Value = Trim(cb_name_cn.Text)
seLbs.Item(SeDftBlockId.Model).Value = Trim(cb_model_no.Text)
seLbs.Item(SeDftBlockId.material).Value = Trim(cb_material.Text)
seLbs.Item(SeDftBlockId.weight).Value = Trim(cb_weight.Text)
'seLBS.Item(SeDftBlockId.version).Value = Trim(version.Text)
seLbs.Item(SeDftBlockId.designer).Value = Trim(cb_designer.Text)
seLbs.Item(SeDftBlockId.design_date).Value = Trim(cb_design_date.Text)
seLbs.Item(SeDftBlockId.reviewer).Value = Trim(cb_reviewer.Text)
seLbs.Item(SeDftBlockId.review_date).Value = Trim(cb_review_date.Text)
seLbs.Item(SeDftBlockId.approver).Value = Trim(cb_approver.Text)
seLbs.Item(SeDftBlockId.approve_date).Value = Trim(cb_approve_date.Text)
End Sub

'=====================Legacy Template=========================
Private Sub getLegacyFromActive()
cb_name_cn.Text = seLbs.Item(LegacySeDftBlockId.name_cn).Value
cb_model_no.Text = seLbs.Item(LegacySeDftBlockId.Model).Value
cb_material.Text = seLbs.Item(LegacySeDftBlockId.material).Value
cb_weight.Text = seLbs.Item(LegacySeDftBlockId.weight).Value
'version.Text = seLBS.Item(LegacySeDftBlockId.version).Value
cb_designer.Text = seLbs.Item(LegacySeDftBlockId.designer).Value
cb_design_date.Text = seLbs.Item(LegacySeDftBlockId.design_date).Value
cb_reviewer.Text = seLbs.Item(LegacySeDftBlockId.reviewer).Value
cb_review_date.Text = seLbs.Item(LegacySeDftBlockId.review_date).Value
cb_approver.Text = seLbs.Item(LegacySeDftBlockId.approver).Value
cb_approve_date.Text = seLbs.Item(LegacySeDftBlockId.approve_date).Value
End Sub
Private Sub writeToLegacyActiveDocument()
seLbs.Item(LegacySeDftBlockId.name_cn).Value = Trim(cb_name_cn.Text)
seLbs.Item(LegacySeDftBlockId.Model).Value = Trim(cb_model_no.Text)
seLbs.Item(LegacySeDftBlockId.material).Value = Trim(cb_material.Text)
seLbs.Item(LegacySeDftBlockId.weight).Value = Trim(cb_weight.Text)
'seLBS.Item(LegacySeDftBlockId.version).Value = Trim(version.Text)
seLbs.Item(LegacySeDftBlockId.designer).Value = Trim(cb_designer.Text)
seLbs.Item(LegacySeDftBlockId.design_date).Value = Trim(cb_design_date.Text)
seLbs.Item(LegacySeDftBlockId.reviewer).Value = Trim(cb_reviewer.Text)
seLbs.Item(LegacySeDftBlockId.review_date).Value = Trim(cb_review_date.Text)
seLbs.Item(LegacySeDftBlockId.approver).Value = Trim(cb_approver.Text)
seLbs.Item(LegacySeDftBlockId.approve_date).Value = Trim(cb_approve_date.Text)
End Sub

Private Sub vUP()
If seApp Is Nothing Then Call Conn2se

Dim sedoc As SolidEdgeDocument
Set sedoc = seApp.ActiveDocument

Dim seDFT As DraftDocument
Set seDFT = sedoc


Dim seBlk As SolidEdgeDraft.BlockOccurrence
Set seBlk = seDFT.ActiveSheet.BlockOccurrences.Item(1)


Dim seLbs As SolidEdgeDraft.BlockLabelOccurrences
Set seLbs = seBlk.BlockLabelOccurrences

Dim vstr
vstr = Split(Trim(seLbs.Item(SeDftBlockId.version).Value), ".")
seLbs.Item(SeDftBlockId.version).Value = vstr(0) & "." & CInt(vstr(1)) + 1

seApp.StatusBar = "Reversion changed to " & seLbs.Item(SeDftBlockId.version).Value
End Sub
Private Sub copyTB()
If seApp Is Nothing Then Call Conn2se

Dim sedoc As SolidEdgeDocument
Set sedoc = seApp.ActiveDocument


'\\ccnsia0u\SolidEdge\Solidedge_Template\Template_ST7\iso metric draft.dft"

Dim seDFT As DraftDocument
Set seDFT = sedoc


Dim seBlk As SolidEdgeDraft.BlockOccurrence
Set seBlk = seDFT.ActiveSheet.BlockOccurrences.Item(1)

Dim neBLK As SolidEdgeDraft.BlockOccurrence
Set neBLK = seDFT.ActiveSheet.BlockOccurrences.Item(2)

Dim seLbs As SolidEdgeDraft.BlockLabelOccurrences
Set seLbs = seBlk.BlockLabelOccurrences

Dim nelbs As SolidEdgeDraft.BlockLabelOccurrences
Set nelbs = neBLK.BlockLabelOccurrences


For i = 1 To nelbs.Count
nelbs.Item(i).Value = seLbs.Item(i).Value
Next


seApp.StatusBar = "Title block copied!"
End Sub


