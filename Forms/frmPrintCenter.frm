VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintCenter 
   Caption         =   "UserForm1"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
   OleObjectBlob   =   "frmPrintCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub page_sn_Change()
Select Case page_sn.SelectedItem.Index
    Case 0  ' SE
        list_sn.Clear
        Call getSnList
    Case 1  'EXCEL
        list_excel.Clear
        Call getExcelList
    Case 2  'FOLDER
End Select
End Sub
Private Sub getSnList()
Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim fname As String
Dim i As Integer
Dim sedoc As SolidEdgeDocument

For i = 1 To Docs.Count
    Set sedoc = Docs.Item(i)
    If sedoc.Type = igDraftDocument Then
        list_sn.AddItem sedoc.Name
    End If
Next
End Sub
Private Sub getExcelList()
Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")


For i = 1 To uRg.Rows.Count
    If uRg.Cells(i, 1).Value = "" Then GoTo nSkip
    filename = Split(uRg.Cells(i, 1).Value, ".")(0)
    
    filename = Trim(filename)
    
    If filename = "" Then GoTo nSkip
    filename = filename & ".dft"

'    If IsFileExists(filename) Then
'
'    End If
    list_excel.AddItem filename
nSkip:
Next i


End Sub

Private Sub UserForm_Initialize()

If seApp Is Nothing Then Call Conn2se

Select Case page_sn.SelectedItem.Index
    Case 0  ' SE
Call getSnList
    Case 1  'EXCEL
Call getExcelList
    Case 2  'FOLDER

End Select

End Sub
