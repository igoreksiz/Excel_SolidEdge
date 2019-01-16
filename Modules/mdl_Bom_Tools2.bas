Attribute VB_Name = "mdl_Bom_Tools2"
Option Explicit



Sub testRest()
Dim RawSource As Variant
RawSource = Array( _
       Array("0.5", "0000180451", "热锌板 DC51D+Z 0.5*1250"), _
       Array("0.6", "0000180476", "热锌板 DC51D+Z 0.6*1250"), _
       Array("0.7", "0000180477", "热锌板 DC51D+Z 0.7*1250"), _
       Array("0.8", "0000180503", "热锌板 DC51D+Z 0.8*1250"), _
       Array("1.0", "0000180520", "热锌板 DC51D+Z 1.0*1250"), _
       Array("1.2", "0000180586", "热锌板 DC52D+Z 1.2*1250"), _
       Array("1.5", "0000180540", "热锌板 DC51D+Z 1.5*1250"), _
       Array("2.0", "0000180544", "热锌板 DC52D+Z 2.0*1250") _
        )
RawSource = Excel.Application.Transpose(RawSource)
RawSource = Excel.Application.Transpose(RawSource)

'Dim a As New Recordset
'a.
'
'Debug.Print a.Fields(1).Name
''a.Fields(2).Name = "code"
''a.Fields(3).Name = "desp"

End Sub

Public Sub OpenTableTemplate(act As String)
Dim tfile As String
Dim tPath As String
tPath = "\\CCNSIF0G\sRdc\CCR\A05-Design\Structure Design\QHC\"

Dim xlDoc As Excel.Workbook

    Select Case act
        Case "Apply_BOM_CODE"
            Set xlDoc = Excel.Workbooks.Add(tPath & "专用号申请表.xlsx")
        Case "Apply_BOM_DESP"
            Set xlDoc = Excel.Workbooks.Add(tPath & "主数据更改申请表.xlsx")
        Case "Apply_BOM_EDIT"
            Set xlDoc = Excel.Workbooks.Add(tPath & "BOM更改单.xlsx")
        Case "Apply_PRODUCT_CODE"
            Set xlDoc = Excel.Workbooks.Add(tPath & "产成品编码申请表.xls")
        Case "Apply_PRODUCT_LABLE"
            Set xlDoc = Excel.Workbooks.Add(tPath & "铭牌模板.xlsx")
    End Select

    Select Case act
        Case "Apply_BOM_CODE", "Apply_BOM_DESP", "Apply_BOM_EDIT"
        'Debug.Print xlDoc.Name
        'xlDoc.Name = Join(Array(xlDoc.Name, Format(Date, "yyyy-mm-dd"), GetSetting("Domisoft", "TBM_SE", "Default_Designer", Excel.Application.UserName)), "_")
        'ActiveWorkbook.Windows(1).Caption = "new name"
    End Select
End Sub
Public Sub formatMultiWorkbooks()
Const path = "D:\Cabinets\E6 SV\BOM\整理物料表\*.xls"
Dim s As String
s = VBA.FileSystem.Dir(path)

Dim list As Variant
list
Dim n As Integer
n = 1
Do Until Len(s) = 0
    s = VBA.FileSystem.Dir
    n = n + 1
    Debug.Print s
Loop

End Sub
