Attribute VB_Name = "mdl_SolidEdge_pub"
Public seApp As SolidEdgeFramework.Application

Public Sub Conn2se()
Set seApp = GetObject(, "SolidEdge.Application")
End Sub
Public Sub Disconn()
Set seApp = Nothing
End Sub
Public Function qDate(d)
qDate = Format(d, "YYYY.MM.DD")
End Function
Public Sub openwithRevMGR()

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")

Dim filename As String
Dim uRg As Excel.Range
Set uRg = Excel.Selection
filename = uRg.Cells(1, 1).Value
If filename = "" Then Exit Sub
filename = filename & ".dft"
filename = myWorkspace & "\" & filename
If IsFileExists(filename) Then '
Dim rmdoc As RevisionManager.Document
'Call rm.OpenFileInRevisionManager(fileName)
Shell "C:\Program Files\Silent Solid Edge\Program\win32\iCnct.exe /r " & filename

Else
MsgBox filename
End If

End Sub
