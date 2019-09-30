VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "²å¼þÅäÖÃ"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cb_ok_Click()
Call SaveSetting("Domisoft", "Config", "SE_Working", seWork.Text)
Call SaveSetting("Domisoft", "Config", "SE_Output", seOutput.Text)
Dim i As Integer
Dim ss As String
For i = 0 To pdf_store.ListCount - 1
ss = ss & "|" & pdf_store.list(i)
Next
Call SaveSetting("Domisoft", "Config", "PDF_Store", Right(ss, Len(ss) - 1))
Call SaveSetting("Domisoft", "Config", "Spec_db_path", spec_db_path.Text)
If seApp Is Nothing Then
'
Else
    Call seApp.SetGlobalParameter(seApplicationGlobalLinkMgmt, LinkMgrPath.Value)
End If
Excel.Application.Cursor = xlDefault '»Ö¸´Êó±ê

Application.EnableEvents = True '»Ö¸´´¥·¢ÊÂ¼þ
Application.Calculation = xlCalculationAutomatic    '×Ô¶¯ÖØËã
Application.ScreenUpdating = True   '¿ªÆôÆÁÄ»Ë¢ÐÂ

Unload Me
End Sub

Private Sub CB_EXIT_Click()
Unload Me
End Sub

Private Sub cb_add_Click()
Dim ss As String
ss = InputBox("paste full path here", "input")
pdf_store.AddItem ss
End Sub

Private Sub cb_del_Click()
If pdf_store.ListIndex >= 0 Then
    pdf_store.RemoveItem pdf_store.ListIndex
End If
End Sub

Private Sub cb_MoveUP_Click()
Dim ss As String
Dim nn As Integer
With pdf_store
If .ListIndex = 0 Then Exit Sub
ss = .list(.ListIndex)
nn = .ListIndex
.RemoveItem nn
.AddItem ss, nn - 1
.Selected(nn - 1) = True
.ListIndex = nn - 1
End With
End Sub

Private Sub cb_MoveDOWN_Click()
Dim ss As String
Dim nn As Integer
With pdf_store
If .ListIndex = .ListCount - 1 Then Exit Sub
ss = .list(.ListIndex)
nn = .ListIndex
.RemoveItem nn
.AddItem ss, nn + 1
.Selected(nn + 1) = True
End With
End Sub



Private Sub UserForm_Initialize()
seWork.Text = GetSetting("Domisoft", "Config", "SE_Working", "")
If Len(seWork.Text) = 0 Then seWork.Text = "S:\Cabinet"

seOutput.Text = GetSetting("Domisoft", "Config", "SE_Output", "")
If Len(seOutput.Text) = 0 Then seOutput.Text = "d:\workspaces"
Dim ss As String
ss = GetSetting("Domisoft", "Config", "PDF_Store", "")

If Len(ss) = 0 Then
    ss = "d:\workspaces|\\CCNSIA1A\SEParts\Cabinet\PDFÍ¼Ö½¿â"
End If

Dim v As Variant
v = Split(ss, "|")
Dim i As Integer
For i = LBound(v) To UBound(v)
pdf_store.AddItem v(i)
Next
spec_db_path.Text = GetSetting("Domisoft", "Config", "Spec_db_path", "")
spec_db_path.AddItem Defualt_DB

lbl_update.Caption = lbl_update.Caption & VBA.FileDateTime(Excel.AddIns.Item(VBA_name).FullName)

Dim lmp As Variant
If seApp Is Nothing Then Exit Sub
Call seApp.GetGlobalParameter(seApplicationGlobalLinkMgmt, lmp)
LinkMgrPath.Text = CStr(lmp)
LinkMgrPath.AddItem "\\CCNSIA1A\SEParts\Admin\Settings\LinkMgmt.txt"
LinkMgrPath.AddItem "\\ccnsif0g\srdc\CCR\A02-Project\B06-Project_2014\Next_Gen_Service_Counter\03-engneering\12-3D_Drawings\01-model\LinkMgmt.txt"
End Sub
