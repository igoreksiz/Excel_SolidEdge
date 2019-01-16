Attribute VB_Name = "mdl_Addin_Pub"
Public Type ColDefType
    DefRow As Integer
    LvlCol As Integer    '层级
    CodeCol As Integer   '专用号
    DespCol As Integer  '物料描述
    TypeCol As Integer  '物料属性
    UnitCol As Integer  '单位
    QtyCol As Integer   '单位用量
    LocCol As Integer   '工位
End Type

'' Public Enum SeDftBlockId
'    Model = 1
'    name_cn = 2
'    drw_no = 3
'    material = 4
'    weight = 5
'    designer = 6
'    design_date = 7
'    reviewer = 8
'    review_date = 9
'    approver = 10
'    approve_date = 11
'    paint_std = 12
'End Enum
    
Public Type BlkId
    Model As Integer
    name_cn As Integer
    drw_no As Integer
    material As Integer
    weight As Integer
    designer As Integer
    design_date As Integer
    reviewer As Integer
    review_date As Integer
    approver As Integer
    approve_date As Integer
    paint_std As Integer
    thk As Integer
    tol As Integer
End Type
 
Public Enum SeDftVerBlockId
    Rev = 1
    Ver = 2
    Phase = 3
End Enum

Public Enum LegacySeDftBlockId
    Model = 1
    name_cn = 2
    name_en = 3
    drw_no = 4
    material = 5
    weight = 6
    version = 7
    designer = 8
    design_date = 9
    reviewer = 10
    review_date = 11
    approver = 12
    approve_date = 13
End Enum
Public Const VBA_name = "Dominic"
Public Const Defualt_DB = "\\ccnsif0g\sdc\Refrigeration\Structure Design\Addons\domisoft.mdb"
Public Function IsFileExists(ByVal strFileName As String) As Boolean
    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function
'Public Function IsFileExists(ByVal strFileName As String) As Boolean
'    Dim objFileSystem As Object
'    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'    If objFileSystem.FileExists(strFileName) = True Then
'        IsFileExists = True
'    Else
'        IsFileExists = False
'    End If
'End Function
Public Function GetFileSize(filespec)
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    GetFileSize = f.Size
End Function


