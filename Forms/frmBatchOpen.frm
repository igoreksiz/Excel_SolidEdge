VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBatchOpen 
   Caption         =   "打开指定DFT"
   ClientHeight    =   1380
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

Dim allowDeepSearch As Boolean
allowDeepSearch = True

Dim fl As New Scripting.Dictionary
If allowDeepSearch Then
    Set fl = ListAllFsoDic(myWorkspace)
Else
    Dim fso As New Scripting.FileSystemObject
    Dim fs As Scripting.Files
    Set fs = fso.GetFolder(myWorkspace).Files
    If fs.Count > 1 Then
        For Each f In fs
            fl(f.Name) = f.path
        Next
    Else
       '
    End If
End If

seApp.Application.DisplayAlerts = False

'If uRg.Find("*") Is Nothing Then
'    filename = InputBox("请输入要打开的文件名\n 例如0080191234:", "文件名")
'    If Len(filename) > 0 Then
'        Call Docs.Open(fl(filename))
'    End If
'End If

For j = 1 To uRg.Columns.Count
    For i = 1 To uRg.Rows.Count
        If uRg.Cells(i, j).Value = "" Then GoTo nSkip
        filename = Split(uRg.Cells(i, j).Value, ".")(0)
        If InStr(1, filename, Chr(10), vbTextCompare) > 0 Then
            filename = Split(filename, Chr(10))(0)           ' TODO 一格里含有多个文件名
        End If
        filename = Trim(filename)
        
        If Len(filename) = 8 And Left(filename, 1) = 8 Then filename = "00" & filename    '解决00问题
        
        'filename = Replace(filename, cha(16), "")
        'filename = Left(filename, 10)
        
        If filename = "" Then GoTo nSkip
        filename = filename & ".dft"
        'filename = myWorkspace & "\" & filename
        
        If fl.Exists(filename) Then
            Call Docs.Open(fl(filename))
        Else
            msG = msG & "," & filename
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

Private Function ListAllFsoDic(myPath$) As Scripting.Dictionary '
    Dim i&, j&
    Dim d1 As New Scripting.Dictionary '字典d1记录子文件夹的绝对路径名
    Dim d2 As New Scripting.Dictionary '字典d2记录文件名(key)和路径(items)
     
     d1(myPath) = ""           '以当前路径myPath作为起始记录，以便开始循环检查
     
    Dim fso As New Scripting.FileSystemObject
    Dim f As Scripting.File
    Do While i < d1.Count
    '当字典1文件夹中有未遍历处理的key存在时进行Do循环 直到 i=d1.Count即所有子文件夹都已处理时停止
 
        kr = d1.Keys '取出文件夹中所有的key即所有子文件夹路径 （注意每次都要更新）
        For Each f In fso.GetFolder(kr(i)).Files '遍历该子文件夹中所有文件 （注意仅从新的kr(i) 开始）
            j = j + 1
            Select Case f.Type
                Case "Solid Edge Draft Document"
                    d2(f.Name) = f.path
            End Select
           '把该子文件夹内的所有文件名作为字典key项加入字典d2 ,重名将被复写*******************************
        Next
 
        i = i + 1 '已经处理过的子文件夹数目 i +1 （避免下次产生重复处理）
        For Each fd In fso.GetFolder(kr(i - 1)).SubFolders '遍历该文件夹中所有新的子文件夹
            d1(fd.path) = " " & fd.Name & ""
            '把新的子文件夹路径存入字典d1以便在下一轮循环中处理
        Next
    Loop
    
    Set ListAllFsoDic = d2

End Function
