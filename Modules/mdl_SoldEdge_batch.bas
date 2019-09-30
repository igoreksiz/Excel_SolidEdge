Attribute VB_Name = "mdl_SoldEdge_batch"
Public Sub closeallDFT(saveChanges As Boolean)
If seApp Is Nothing Then Call Conn2se

Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim sedoc As SolidEdgeDocument
Dim i As Integer

Excel.Application.Cursor = xlWait '修改鼠标为等待

For i = Docs.Count To 1 Step -1
    Set sedoc = Docs.Item(i)
    If sedoc.Type = igDraftDocument Then
        sedoc.Close saveChanges
    End If
Next

Excel.Application.Cursor = xlDefault '恢复鼠标

Set sedoc = Nothing
Set Docs = Nothing
End Sub

Public Sub openSelectedPDF()
Dim myPDFstore As String

Dim p As Variant
p = Split(GetSetting("Domisoft", "Config", "PDF_Store", ""), "|")

Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim filename As String
Dim done As Boolean
done = False
For i = LBound(p) To UBound(p)

    myPDFstore = p(i)
    
    filename = Trim(uRg.Cells(1, 1).Value)
    
    If InStr(1, filename, Chr(10), vbTextCompare) > 0 Then
        filename = Split(filename, Chr(10))(0)           ' TODO 一格里含有多个文件名
    End If
    If Len(filename) = 8 And Left(filename, 1) = 8 Then filename = "00" & filename    '解决00问题
    
    filename = myPDFstore & "\" & filename & ".pdf"
    
    
    
    If IsFileExists(filename) Then
        Excel.Application.Cursor = xlWait '修改鼠标为等待
        Shell "explorer.exe " & filename
        Excel.Application.Cursor = xlDefault '恢复鼠标
        done = True
        Exit For
    End If
Next
If Not done Then MsgBox "file not found: " & filename, vbOKOnly, "File Not Found"
End Sub
Public Sub editBB()
If seApp Is Nothing Then Call Conn2se



Dim seDFT As DraftDocument
Set seDFT = seApp.ActiveDocument

Dim seSht As SolidEdgeDraft.sheet
Set seSht = seDFT.ActiveSheet

Dim BBS As Balloons
Set BBS = seSht.Balloons

Dim ee As String
ee = "a"

Dim revBlk As SolidEdgeDraft.BlockOccurrence
Dim revBlkLBLs As SolidEdgeDraft.BlockLabelOccurrences

For i = 1 To seSht.BlockOccurrences.Count
    If seSht.BlockOccurrences.Item(i).Block.Name = "变更修改" Then
        Set revBlk = seSht.BlockOccurrences.Item(i)
        Set revBlkLBLs = revBlk.BlockLabelOccurrences
        Exit For
    End If
Next
For i = 1 To revBlkLBLs.Count
    If revBlkLBLs.Item(i).Name = "标记" Then
        ee = revBlkLBLs.Item(i).Value
        Exit For
    End If
Next

Dim e As String
e = InputBox("What symbol change bolloons to ?", "input", ee)

If Len(e) = 0 Then Exit Sub

For i = 1 To BBS.Count

If BBS.Item(i).BalloonType = igDimBalloonCircle Then
    If BBS.Item(i).Leader = False Then
        BBS.Item(i).BalloonText = e
    End If
End If

Next i

Set seDFT = Nothing

AppActivate seApp.Name
End Sub
Public Sub deleteBB()
If seApp Is Nothing Then Call Conn2se



Dim seDFT As DraftDocument
Set seDFT = seApp.ActiveDocument

Dim seSht As SolidEdgeDraft.sheet
Set seSht = seDFT.ActiveSheet

Dim BBS As Balloons
Set BBS = seSht.Balloons

Dim ee As String
ee = "a"
For i = 1 To BBS.Count
    If BBS.Item(i).BalloonType = igDimBalloonCircle Then
        If BBS.Item(i).Leader = False Then
            ee = BBS.Item(i).BalloonText
            Exit For
        End If
    End If
Next

Dim e As String
e = InputBox("What symbol bolloons you want to delete ?", "input", ee)

If Len(e) = 0 Then Exit Sub

For i = BBS.Count To 1 Step -1

If BBS.Item(i).BalloonType = igDimBalloonCircle Then
    If BBS.Item(i).Leader = False Then
        If BBS.Item(i).BalloonText = e Then
            BBS.Item(i).Delete
        End If
    End If
End If

Next i

Set seDFT = Nothing

AppActivate seApp.Name
End Sub
Public Sub saveDXF()
If seApp Is Nothing Then Call Conn2se

Dim seDFT As DraftDocument
Set seDFT = seApp.ActiveDocument

Dim seMdl As SheetMetalDocument
Set seMdl = seDFT.ModelLinks.Item(1).ModelDocument

'seMdl.Activate

Dim sepsms As SolidEdgePart.Models
Set sepsms = seMdl.Models

Dim fn As String
fn = GetSetting("Domisoft", "Config", "SE_Output", "") & "\" & Split(seDFT.Name, ".")(0) & ".dxf"

seApp.Application.DisplayAlerts = False

Call sepsms.SaveAsFlatDXFEx(fn, Nothing, Nothing, Nothing, True)

seApp.Application.DisplayAlerts = True

Set seDFT = Nothing
Set seMdl = Nothing

End Sub


