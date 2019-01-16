VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBatchOpen 
   Caption         =   "打开指定DFT"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   OleObjectBlob   =   "frmBatchOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBatchOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Me.Hide
AppActivate seApp.Name
End Sub

Private Sub UserForm_Activate()
ListBox1.Clear
'TextBox1.Value = "Loading...please wait..."
frmBatchOpen.Height = 97
Call batchOpen

End Sub

Sub batchOpen()
If seApp Is Nothing Then Call Conn2se

Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")

Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim filename As String
Dim msG As String
msG = ""

seApp.Application.DisplayAlerts = False
For j = 1 To uRg.Columns.Count
    For i = 1 To uRg.Rows.Count
        If uRg.Cells(i, j).Value = "" Then GoTo nSkip
        filename = Split(uRg.Cells(i, j).Value, ".")(0)
        
        filename = Trim(filename)
        'filename = Replace(filename, cha(16), "")
        'filename = Left(filename, 10)
        
        If filename = "" Then GoTo nSkip
        filename = filename & ".dft"
        filename = myWorkspace & "\" & filename
        If IsFileExists(filename) Then
            Call Docs.Open(filename)
        Else
            msG = msG & "," & uRg.Cells(i, j).Value
        End If
nSkip:
    Next i
Next j
seApp.Application.DisplayAlerts = True

Set Docs = Nothing

If msG = "" Then
'    TextBox1.ForeColor = vbBlack
'    TextBox1.Value = "All done"
'CommandButton2.Enabled = True
    Me.Hide
    AppActivate seApp.Name
Else
    ListBox1.Visible = True
    frmBatchOpen.Height = 184
    
    Dim lit As Variant
    lit = Split(Trim(Join(Split(msG, ","), " ")), " ")

    For k = LBound(lit) To UBound(lit)
            ListBox1.AddItem lit(k), k
            ListBox1.list(k, 1) = "找不到此图纸!"
    Next k
    CommandButton2.Caption = "  Click to return to Solid Edge >>"
End If
End Sub

