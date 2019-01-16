VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDFselected 
   Caption         =   "Create PDF from selection"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "frmPDFselected.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPDFselected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub batchOpenPrint()
If seApp Is Nothing Then Call Conn2se

Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")



Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

seApp.Application.DisplayAlerts = False

Dim seDFT As DraftDocument

Dim dftName As String
Dim pdfName As String


For i = 0 To ListBox1.ListCount - 1

    If Left(ListBox1.list(i, 2), 3) = "DFT" Then GoTo nSkip
    
    dftName = myWorkspace & "\" & ListBox1.list(i, 1) & ".dft"
    pdfName = toFolder & "\" & ListBox1.list(i, 1) & ".pdf"

    Set seDFT = Docs.Open(dftName)
    seDFT.Activate '如果不Activate,只会重复保存当前文件
    seDFT.SaveCopyAs pdfName  'savecopyas 和saveas效果一样
    seDFT.Close False 'TODO: 如果已经在SE里打开了的文件如何处理
    ListBox1.list(i, 2) = "PDF created!"
nSkip:
Next i

seApp.Application.DisplayAlerts = True

Set Docs = Nothing

End Sub


Private Sub CommandButton1_Click()

Call batchOpenPrint
CommandButton1.Enabled = False
End Sub

Private Sub UserForm_Initialize()

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")


Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim filename As String
Dim dftName As String
Dim pdfName As String

Dim n As Integer
n = 1

For j = 1 To uRg.Columns.Count
    For i = 1 To uRg.Rows.Count
        If uRg.Cells(i, j).Value = "" Then GoTo nSkip
        filename = Split(uRg.Cells(i, j).Value, ".")(0)
        If filename = "" Then GoTo nSkip
        dftName = myWorkspace & "\" & filename & ".dft"
        pdfName = toFolder & "\" & filename & ".pdf"
        With ListBox1
            .AddItem
            .list(n - 1, 0) = n
            .list(n - 1, 1) = filename
            If Not IsFileExists(dftName) Then .list(n - 1, 2) = "DFT file not exists!"
            If IsFileExists(pdfName) Then .list(n - 1, 2) = "PDF file exists! overwrite!"
        End With
        n = n + 1
nSkip:
    Next i
Next j
End Sub
