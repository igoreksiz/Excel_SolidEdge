Attribute VB_Name = "Module1"
Sub insertPaintTable()
If seApp Is Nothing Then Call Conn2se
Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim dft As SolidEdgeDraft.DraftDocument
Set dft = seApp.ActiveDocument
Dim sht As SolidEdgeDraft.sheet
Set sht = dft.ActiveSheet

Dim paintTable As SolidEdgeDraft.Table
Set paintTable = dft.Tables.Add
paintTable.ShowColumnHeader = True

Dim row As SolidEdgeDraft.TableRow
Set row = paintTable.Rows.Add(0, True)

paintTable.HeaderFixedRowHeight = 0.007
paintTable.DataFixedRowHeight = 0.007

Dim col As SolidEdgeDraft.TableColumn
Set col = paintTable.Columns.Add(0, True)
col.Width = 0.035
col.HeaderRowValue = "专用号(未喷粉)"

paintTable.Cell(1, 1).Value = uRg.Cells(1, 1).Text
Call formatCol(col)

Set col = paintTable.Columns.Add(1, True)
col.Width = 0.035
col.HeaderRowValue = "专用号(已喷粉)"
paintTable.Cell(1, 2).Value = uRg.Cells(2, 1).Text
Call formatCol(col)

dft.SelectSet.RemoveAll
dft.SelectSet.Add paintTable
dft.SelectSet.Cut

VBA.AppActivate seApp.Name

Call SendPaste

End Sub
Private Sub formatCol(seCol As SolidEdgeDraft.TableColumn)
seCol.HeaderRowVerticalAlignment = igTextHzAlignVCenter
seCol.HeaderRowHorizontalAlignment = igTextHzAlignCenter
seCol.DataHorizontalAlignment = igTextHzAlignCenter
seCol.DataVerticalAlignment = igTextHzAlignVCenter
End Sub
Public Sub FormatExistPDF()
Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim toFolder As String
toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")

Dim filename As String
Dim fRow As Excel.Range

For i = 1 To uRg.Rows.Count
    filename = toFolder & "\" & Trim(uRg.Cells(i, 1).Text) & ".pdf"
    If IsFileExists(filename) Then
        Set fRow = uRg.Rows(i)
        fRow.Interior.Color = RGB(146, 208, 80)
    End If
Next
End Sub
Public Sub MoveSelectedPDF()
Dim uRg As Excel.Range
Set uRg = Excel.Selection

'Dim fromFolder As String, toFolder As String
'toFolder = GetSetting("Domisoft", "Config", "SE_Output", "")
Const fromFolder = "S:\Cabinet\PDF图纸库"
Const toFolder = "Y:\A02-Project\B11-Project_2019\01_Monaxis Low Front\11-Drawings to QHC\已发图纸\外购件"

Dim filename As String
Dim MoveName As String
Dim fRow As Excel.Range

'Dim folderName As String
'folderName = InputBox(toFolder & vbCrLf & "输入新文件夹名称:", "新建文件夹", "自制件 外购件")

'folderName = Trim(folderName)
'If Len(folderName) = 0 Then Exit Sub

'folderName = toFolder & "\" & folderName

Dim fso As New FileSystemObject

'If fso.FolderExists(folderName) Then
'    MsgBox "文件夹已经存在:" & vbCrLf & folderName
'Else
'    fso.CreateFolder folderName
'End If

Debug.Print Err.Number
If Err.Number > 0 Then Exit Sub


Dim nCount As Integer
nCount = 0
For i = 1 To uRg.Rows.Count
    filename = fromFolder & "\" & Trim(uRg.Cells(i, 1).Text) & ".pdf"
    MoveName = toFolder & "\" & Trim(uRg.Cells(i, 1).Text) & ".pdf"
    If IsFileExists(filename) Then
    
        fso.CopyFile filename, MoveName '复制PDF
        
'        filename = Replace(filename, ".pdf", ".dxf", , , vbTextCompare)
'        MoveName = Replace(MoveName, ".pdf", ".dxf", , , vbTextCompare)
'        If IsFileExists(filename) Then
'            fso.MoveFile filename, MoveName '移动同名DXF
'        End If
        
        nCount = nCount + 1 '计数
    End If
Next
Set fso = Nothing
MsgBox "移动了" & nCount & "个文件!"
End Sub
Public Sub verticalMerge()


Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim tRg As Excel.Range

Excel.Application.DisplayAlerts = False

For i = 1 To uRg.Columns.Count
    If i > 100 Then Exit For
    Set tRg = uRg.Columns(i)
    tRg.Merge
    tRg.VerticalAlignment = xlCenter
    tRg.HorizontalAlignment = xlCenter
Next

Excel.Application.DisplayAlerts = True

End Sub
Public Sub verticalUnmerge()
Dim uRg As Excel.Range
Set uRg = Excel.Selection

Dim tRg As Excel.Range
For i = 1 To uRg.Columns.Count
    Set tRg = uRg.Columns(i)
    tRg.UnMerge
Next
End Sub
Public Sub openDummy()

If seApp Is Nothing Then Call Conn2se

Dim pAsm As SolidEdgeAssembly.AssemblyDocument

If seApp.ActiveDocumentType <> igAssemblyDocument Then Exit Sub

Set pAsm = seApp.ActiveDocument

If pAsm.SelectSet.Count > 1 Then Exit Sub


Dim pOc As SolidEdgeAssembly.Occurrence


Select Case pAsm.SelectSet.Item(1).Type
    Case igReference    '子装配里的子装配
        Dim pRasm As SolidEdgeFramework.Reference
        Set pRasm = pAsm.SelectSet.Item(1)
        Set pOc = pRasm.Object
    Case igSubAssembly
        Set pOc = pAsm.SelectSet.Item(1)
    Case Else
        Debug.Print pAsm.SelectSet.Item(1).Type
        Exit Sub
End Select

Dim oldName As String
oldName = pOc.OccurrenceFileName

seApp.Documents.Open (Split(oldName, "!")(0))

End Sub
Public Sub editDim()
    Dim SelSet As Object
    If seApp Is Nothing Then Call Conn2se
    
    Set SelSet = seApp.ActiveDocument.SelectSet
    
    If SelSet.Count <> 1 Then
        MsgBox "You must select a single dimension.", , "Edit Dimension"
        Exit Sub
    Else
        ' Make sure selected object is a dimension
        If SelSet(1).Type <> igDimension Then
            MsgBox "You must select a single dimension.", , "Edit Dimension"
            Exit Sub
        Else
            Dim s As String
            s = InputBox("input text to replace dimension", "Edit Dim", "L")
            If Len(Trim(s)) > 0 Then
                SelSet(1).OverrideString = s
                SelSet(1).Style.NTSSymbol = igDimStyleNTSNone
            End If
        End If
    End If
End Sub
Public Sub closeThenOpenReadonly()
Dim wb As Excel.Workbook
Set wb = Excel.ActiveWorkbook

'If wb.ReadOnly = True Then Exit Sub

Dim path As String
path = wb.FullName

wb.Close False

Excel.Application.Workbooks.Open path, False, True

End Sub
