Attribute VB_Name = "mdl_SE_Print"
Public Sub printPDF()
If seApp Is Nothing Then Call Conn2se

Dim dft As DraftDocument

If seApp.ActiveDocumentType = igDraftDocument Then
    Set dft = seApp.ActiveDocument
Else
    MsgBox "DFT ONLY"
    Exit Sub
End If

Dim fname As String
Dim toFolder As String
'toFolder = dft.Path

toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")
fname = toFolder & "\" & Split(dft.Name, ".")(0) & ".pdf"

'需要引用Constants库
Dim currentV As Variant
Call seApp.GetGlobalParameter(seApplicationGlobalDraftSaveAsPDFSheetOptions, currentV)  '保存现有值

seApp.SetGlobalParameter seApplicationGlobalDraftSaveAsPDFSheetOptions, SolidEdgeConstants.DraftSaveAsPDFSheetOptionsConstants.seDraftSaveAsPDFSheetOptionsConstantsAllSheets

seApp.Application.DisplayAlerts = False
Excel.Application.Cursor = xlWait '修改鼠标为等待

dft.SaveAs fname, False, False

Excel.Application.Cursor = xlDefault '恢复鼠标
seApp.Application.DisplayAlerts = True

seApp.SetGlobalParameter seApplicationGlobalDraftSaveAsPDFSheetOptions, CInt(currentV)


Set dft = Nothing

seApp.StatusBar = fname & vbTab & "Done"

Shell "explorer.exe " & fname

End Sub


Public Sub printDWG()
If seApp Is Nothing Then Call Conn2se

Dim dft As DraftDocument
Set dft = seApp.ActiveDocument

Dim gname As String
Dim toFolder As String
'toFolder = dft.Path

toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

gname = toFolder & "\" & Split(dft.Name, ".")(0) & ".dwg"

seApp.Application.DisplayAlerts = False

dft.SaveAs gname, False, False

seApp.Application.DisplayAlerts = True

Set dft = Nothing

seApp.StatusBar = fname & vbTab & "Done"

End Sub
Public Sub printAll_PDF()
If seApp Is Nothing Then Call Conn2se

Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

Dim fname As String
Dim i As Integer
Dim sedoc As SolidEdgeDocument

Dim currentV As Variant
Call seApp.GetGlobalParameter(seApplicationGlobalDraftSaveAsPDFSheetOptions, currentV)  '保存现有值

seApp.SetGlobalParameter seApplicationGlobalDraftSaveAsPDFSheetOptions, SolidEdgeConstants.DraftSaveAsPDFSheetOptionsConstants.seDraftSaveAsPDFSheetOptionsConstantsAllSheets

seApp.Application.DisplayAlerts = False
Excel.Application.Cursor = xlWait '修改鼠标为等待

For i = 1 To Docs.Count
    Set sedoc = Docs.Item(i)
    If sedoc.Type = igDraftDocument Then

        fname = toFolder & "\" & Split(sedoc.Name, ".")(0) & ".pdf"
        sedoc.Activate '如果不Activate,只会重复保存当前文件
        sedoc.SaveCopyAs fname  'savecopyas 和saveas效果一样
    End If
Next
Excel.Application.Cursor = xlDefault '恢复鼠标
seApp.Application.DisplayAlerts = True
seApp.SetGlobalParameter seApplicationGlobalDraftSaveAsPDFSheetOptions, CInt(currentV)


Set sedoc = Nothing
Set Docs = Nothing

seApp.StatusBar = "all pdf Done"

End Sub
Public Sub PrintAll_DXF()
If seApp Is Nothing Then Call Conn2se
Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

Dim fname As String
Dim i As Integer
Dim sedoc As SolidEdgeDocument

Dim seDFT As DraftDocument
Dim seMdl As SheetMetalDocument
Dim sepsms As SolidEdgePart.Models
Dim fn As String
Excel.Application.Cursor = xlWait '修改鼠标为等待

Dim k As Integer
For i = 1 To Docs.Count
    Set sedoc = Docs.Item(i)
    If sedoc.Type = igDraftDocument Then

        Set seDFT = sedoc

        For k = 1 To seDFT.ActiveSheet.DrawingViews.Count
            If seDFT.ActiveSheet.DrawingViews.Item(k).ModelLink.ModelDocument.Type = igSheetMetalDocument Then
                Set seMdl = seDFT.ActiveSheet.DrawingViews.Item(k).ModelLink.ModelDocument
                GoTo continueGo
            End If
        Next k
    
        If seMdl Is Nothing Then GoTo errH

continueGo:

'       Set seMdl = seDft.ModelLinks.Item(1).ModelDocument

        
        'seMdl.Activate
        Set sepsms = seMdl.Models
        fn = toFolder & "\" & Split(seDFT.Name, ".")(0) & ".dxf"
        
        seApp.Application.DisplayAlerts = False
        
        Call sepsms.SaveAsFlatDXFEx(fn, Nothing, Nothing, Nothing, True)
 
'Debug.Print seDft.Name & vbTab & seDft.ActiveSheet.DrawingViews.Item(1).CaptionDefinitionTextPrimary & vbTab & seMdl.Name & vbTab & GetFileSize(fn)
        
        seApp.Application.DisplayAlerts = True
        
        If GetFileSize(fn) = 157715 Then
            MsgBox "文件:" & seDFT.Name & " 存在未知错误! DXF文件生成失败! 点击OK继续下一个", vbCritical + vbOKOnly, "Error"
            Kill fn
        End If
        
    End If
Next
Excel.Application.Cursor = xlDefault '恢复鼠标

Set seDFT = Nothing
Set seMdl = Nothing
Set sedoc = Nothing
Set Docs = Nothing
Exit Sub

errH:
Excel.Application.Cursor = xlDefault '恢复鼠标
MsgBox seDFT.Name & "出错了!没有找到钣金视图!"
End Sub
Public Sub PrintAll_DWG()
If seApp Is Nothing Then Call Conn2se
Dim Docs As SolidEdgeFramework.Documents
Set Docs = seApp.Documents

Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

Dim i As Integer
Dim sedoc As SolidEdgeDocument
Dim seDFT As DraftDocument

Dim gname As String

Excel.Application.Cursor = xlWait '修改鼠标为等待

For i = 1 To Docs.Count
    Set sedoc = Docs.Item(i)
    If sedoc.Type = igDraftDocument Then
        Set seDFT = sedoc
        gname = toFolder & "\" & Split(seDFT.Name, ".")(0) & ".dwg"
        seApp.Application.DisplayAlerts = False
        seDFT.SaveAs gname, False, False
        seApp.Application.DisplayAlerts = True
        Call adjustCAD(gname)
    End If
Next

Excel.Application.Cursor = xlDefault '恢复鼠标

Set seDFT = Nothing
Set sedoc = Nothing
Set Docs = Nothing

End Sub
Sub adjustCAD(fn As String)



Dim cadapp As AutoCAD.AcadApplication
Set cadapp = GetObject(, "AutoCAD.Application")

Dim cadDoc As AutoCAD.AcadDocument
Set cadDoc = cadapp.Documents.Open(fn)

Dim i As Integer
For i = 1 To cadDoc.Layouts.Count - 1   ' skip item(0) "Model"
    cadDoc.Layouts.Item(i).CanonicalMediaName = "ISO_A3_(420.00_x_297.00_MM)"
    cadDoc.Layouts.Item(i).PlotRotation = ac0degrees
Next
cadDoc.Regen acActiveViewport
cadDoc.Close True
Set cadDoc = Nothing
Set cadapp = Nothing
End Sub

Sub TESTCAD()
Dim cadapp As AutoCAD.AcadApplication
Set cadapp = GetObject(, "AutoCAD.Application")

Dim cadDoc As AutoCAD.AcadDocument
Set cadDoc = cadapp.Documents.Open("D:\Selene_Corner_Case\out_put\0080131451.dwg")

Dim k As Integer
Dim sAll As AutoCAD.AcadSelectionSet
'For k = 1 To cadDoc.Layouts.Count
    cadDoc.Layouts.Item(1).CanonicalMediaName = "ISO_A3_(420.00_x_297.00_MM)"
    cadDoc.Layouts.Item(1).PlotRotation = ac0degrees

    
'Next k
cadDoc.Regen acActiveViewport
Set sAll = cadDoc.SelectionSets.Add("all")

sAll.Select acSelectionSetAll




'Dim a As Variant
'a = cadDoc.Layouts.Item(1).GetCanonicalMediaNames
'
'For i = LBound(a) To UBound(a)
' Debug.Print a(i)
'Next

End Sub
