Attribute VB_Name = "mdl_Bom_Tools"
Option Explicit

Dim ColDef As ColDefType

'备忘
'钢板:7.84*长（mm）*宽（mm）*厚（mm）*10-6/0.90
'铝：2.7*长（mm）*宽（mm）*厚（mm）*10-6/0.95
'ABS: 1.05*长（mm）*宽（mm）*厚（mm）*10-6/0.90
'PS: 1.07*长（mm）*宽（mm）*厚（mm）*10-6/0.90
'粉末：0.18*长（mm）*宽（mm）*面数（1面或两面）*10-6/0.9（白粉）、0.18*长（mm）*宽（mm）*面数（1面或两面）*10-6/0.5（特粉）
'发泡:M（总质量）=体积（M3）*44   白料=M/2.2  黑料=白料*1.2或黑料=M-白料


Public Sub switchCell()
    Dim i As Integer
    Dim j As Integer
    Dim uRg As Excel.Range
    If Excel.Selection Is Nothing Then Exit Sub
    Set uRg = Excel.Selection
    If uRg.Rows.Count < 2 Then Exit Sub

    Dim ccc()
    ReDim ccc(uRg.Rows.Count, uRg.Columns.Count)
    
    Dim nf As String
    If Not uRg.NumberFormat = Null Then
    nf = uRg.NumberFormat
    End If
    uRg.NumberFormat = "@"  '先设计为文本格式,解决专用号自动转变成数字的问题
    
    ccc = uRg.Value
    
    Dim rCt As Integer
    rCt = uRg.Rows.Count
    If rCt Mod 2 <> 0 Then rCt = rCt - 1    '解决误选了奇数行的问题
    
    For j = 1 To rCt Step 2
        For i = 1 To uRg.Columns.Count
            uRg.Cells(j, i).Value = ccc(j + 1, i)
            uRg.Cells(j + 1, i).Value = ccc(j, i)
        Next
    Next
    If nf = "" Then
        uRg.NumberFormat = Null
    Else
        uRg.NumberFormat = nf
    End If
End Sub
Public Sub demoteLvl()
    Dim oSht As Excel.Worksheet
    Set oSht = Excel.ActiveSheet
    
    Dim uRg As Excel.Range
    Set uRg = Excel.Selection
    Dim i As Integer
    For i = 1 To uRg.Rows.Count
    
        If Len(uRg.Cells(i, 1).Value) > 0 Then
            uRg.Cells(i, 1).Value = getSubLevel(uRg.Cells(i, 1).Value)
        End If
    Next
End Sub
Private Function getSubLevel(currentLevel As String) As String
'Debug.Print currentLevel
        getSubLevel = Left(currentLevel, 1) & Left(currentLevel, Len(currentLevel) - 1) & CStr(CInt(Right(currentLevel, 1)) + 1)
End Function
Private Function getUpLevel(currentLevel As String) As String
        getUpLevel = Left(currentLevel, Len(currentLevel) - 2) & CStr(CInt(Right(currentLevel, 1)) - 1)
End Function
Private Function getUpLevelRow(ByVal currentRow As Integer) As Integer
        Dim cLvl As Integer
        cLvl = getLvlNum(Cells(currentRow, ColDef.LvlCol).Text)
        
        Dim tLvl As Integer
        tLvl = cLvl
        
        Dim n As Integer
        n = currentRow
        
        Do Until cLvl - tLvl = 1
            n = n - 1
            tLvl = getLvlNum(Cells(n, ColDef.LvlCol).Text)
        Loop
        
        getUpLevelRow = n
        
End Function
Private Function getLvlNum(lvlStr As String) As Integer
        Dim s As Variant
        s = Split(lvlStr, ".")
        getLvlNum = CInt(s(UBound(s)))
End Function
Public Sub addRawMaterial()
    Dim oSht As Excel.Worksheet
    If Excel.ActiveSheet Is Nothing Then Exit Sub
    Set oSht = Excel.ActiveSheet
    
    Call GetBomColumn
    
   
    Dim uRg As Excel.Range
    
    If Excel.Selection Is Nothing Then Exit Sub
    
    Set uRg = Excel.Selection
    
    Dim rangeHeight As Integer
    rangeHeight = uRg.Rows.Count
    
    Dim i As Integer
    For i = 1 To rangeHeight + 1
        If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "DC51D+Z", vbTextCompare) > 0 Or InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "DC52D+Z", vbTextCompare) > 0 Or InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "热锌钣", vbTextCompare) > 0 Or InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "热锌板", vbTextCompare) > 0 Then
            If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "已喷", vbTextCompare) Then
                If InStr(1, uRg.Cells(i + 1, ColDef.DespCol).Value, "未喷", vbTextCompare) < 1 Then
                    MsgBox "缺少[未喷]项在line:" & i
                    Exit Sub
                End If
                i = i + 1   '跳过已喷粉行
                
                '跳过已有行
                If uRg.Cells(i, ColDef.TypeCol).Value = "自制" Then
                    If uRg.Cells(i + 1, ColDef.LvlCol).Value = getSubLevel(uRg.Cells(i, ColDef.LvlCol).Value) And uRg.Cells(i + 2, ColDef.CodeCol).Value = "0088200085" Then
                        i = i + 2
                        GoTo nSkip
                    End If
                Else
                    If uRg.Cells(i + 1, ColDef.LvlCol).Value = uRg.Cells(i, ColDef.LvlCol).Value And uRg.Cells(i + 1, ColDef.CodeCol).Value = "0088200085" Then
                        i = i + 1
                        GoTo nSkip
                    End If
                End If
                '跳过结束
                
                uRg.Rows(i + 1).Insert
                rangeHeight = rangeHeight + 1
                uRg.Cells(i + 1, ColDef.LvlCol).Value = uRg.Cells(i, ColDef.LvlCol).Text
                uRg.Cells(i + 1, ColDef.CodeCol).Value = "'" & "0088200085"
                uRg.Cells(i + 1, ColDef.DespCol).Value = "白色粉沫 RAL9003"
                uRg.Cells(i + 1, ColDef.LocCol).Value = "601"
                uRg.Cells(i + 1, ColDef.QtyCol).Value = CalcPdWT(uRg.Cells(i, ColDef.DespCol).Text)
                uRg.Cells(i + 1, ColDef.UnitCol).Value = "公斤"
                uRg.Cells(i + 1, ColDef.TypeCol).Value = "外购"
                
                If uRg.Cells(i, ColDef.TypeCol).Value = "自制" Then
                    uRg.Rows(i + 1).Insert
                    rangeHeight = rangeHeight + 1
                    uRg.Cells(i + 1, ColDef.LvlCol).Value = getSubLevel(uRg.Cells(i, ColDef.LvlCol).Text)
                    uRg.Cells(i + 1, ColDef.CodeCol).Value = "'" & getRawShtNo(uRg.Cells(i, ColDef.DespCol).Text)
                    uRg.Cells(i + 1, ColDef.DespCol).Value = getRawShtDesp(uRg.Cells(i + 1, ColDef.CodeCol).Text)
                    uRg.Cells(i + 1, ColDef.LocCol).Value = "101"
                    uRg.Cells(i + 1, ColDef.QtyCol).Value = CalcShtWT(uRg.Cells(i, ColDef.DespCol).Text)
                    uRg.Cells(i + 1, ColDef.UnitCol).Value = "公斤"
                    uRg.Cells(i + 1, ColDef.TypeCol).Value = "外购"
                    i = i + 1
                End If
                i = i + 1
            Else
                '跳过已有行
                If uRg.Cells(i, ColDef.TypeCol).Value = "自制" Then
                    If uRg.Cells(i + 1, ColDef.LvlCol).Value = getSubLevel(uRg.Cells(i, ColDef.LvlCol).Text) Then
                        i = i + 1
                        GoTo nSkip
                    End If
                Else
                    GoTo nSkip
                End If
                '跳过结束
                If uRg.Cells(i, ColDef.TypeCol).Value = "自制" Then
                    uRg.Rows(i + 1).Insert
                    rangeHeight = rangeHeight + 1
                    uRg.Cells(i + 1, ColDef.LvlCol).Value = getSubLevel(uRg.Cells(i, ColDef.LvlCol).Text)
                    uRg.Cells(i + 1, ColDef.CodeCol).Value = "'" & getRawShtNo(uRg.Cells(i, ColDef.DespCol).Text)
                    uRg.Cells(i + 1, ColDef.DespCol).Value = getRawShtDesp(uRg.Cells(i + 1, ColDef.CodeCol).Text)
                    uRg.Cells(i + 1, ColDef.LocCol).Value = "101"
                    uRg.Cells(i + 1, ColDef.QtyCol).Value = CalcShtWT(uRg.Cells(i, ColDef.DespCol).Text)
                    uRg.Cells(i + 1, ColDef.UnitCol).Value = "公斤"
                    uRg.Cells(i + 1, ColDef.TypeCol).Value = "外购"
                End If
                i = i + 1
            End If
        Else
            If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "麻纹不锈钢板", vbTextCompare) > 0 Then
                 uRg.Cells(i, ColDef.QtyCol).Value = CalcShtWT(uRg.Cells(i - 1, ColDef.DespCol).Value)
            End If
        End If
nSkip:
    Next
End Sub

Private Sub GetBomColumn()
Dim asht As Excel.Worksheet

If Excel.ActiveSheet Is Nothing Then Exit Sub

Set asht = Excel.ActiveSheet



Dim i As Integer

For i = 1 To asht.UsedRange.Rows.Count
    If asht.Cells(i, 1).Text = "层级" Or asht.Cells(i, 1).Text = "层次" Or asht.Cells(i, 2).Text = "展开层" Then
        ColDef.DefRow = i
        Exit For
    End If
Next

ColDef.LvlCol = 1
        
If ColDef.DefRow = 0 Then Exit Sub


For i = 1 To asht.UsedRange.Columns.Count
Select Case asht.Cells(ColDef.DefRow, i).Text
    Case "子项物料代码", "专用号", "物料代码", "对象标识"
        ColDef.CodeCol = i
    Case "物料名称", "物料描述", "对象描述"
        ColDef.DespCol = i
    Case "物料属性", "属性", "物料类型"
        ColDef.TypeCol = i
    Case "单位", "组件单位"
        ColDef.UnitCol = i
    Case "数量", "单位用量", "用量", "组件数量(CUn)"
        ColDef.QtyCol = i
    Case "工位", "排序字符串"
        ColDef.LocCol = i
End Select

If i > 100 Then Exit For

Next

If ColDef.CodeCol = 0 Then Debug.Print "没找到专用号所在列"
If ColDef.DespCol = 0 Then Debug.Print "没找到物料描述所在列"
If ColDef.TypeCol = 0 Then Debug.Print "没找到物料属性所在列"
If ColDef.UnitCol = 0 Then Debug.Print "没找到单位所在列"
If ColDef.QtyCol = 0 Then Debug.Print "没找到单位用量所在列 "
If ColDef.LocCol = 0 Then Debug.Print "没找到工位所在列"

End Sub
Private Function getRawShtNo(partdesp As String) As String

Dim uv As Variant
uv = Split(partdesp, " ")

Dim v As String
Dim i As Integer
Dim j As Integer

For i = LBound(uv) To UBound(uv)
    If InStr(1, uv(i), "*", vbTextCompare) > 1 Then
        v = uv(i)
        Exit For
    End If
Next

If Len(v) = 0 Then
    MsgBox "getRawShtNo error:" & partdesp
    getRawShtNo = ""
    Exit Function
End If

Dim s As Variant
s = Split(v, "*")

'冒泡排序
For i = UBound(s) - 1 To LBound(s) Step -1
    For j = LBound(s) To i
        If CDbl(s(j)) > CDbl(s(j + 1)) Then
            s(j) = CDbl(s(j)) + CDbl(s(j + 1))
            s(j + 1) = CDbl(s(j)) - CDbl(s(j + 1))
            s(j) = CDbl(s(j)) - CDbl(s(j + 1))
        End If
    Next j
Next i

Dim t As String
t = Round(s(LBound(s)), 1)

Select Case t
    Case "0.5", ".5", 0.5
        getRawShtNo = "0087500182"
    Case "0.6", ".6", 0.6
        getRawShtNo = "0087500185"
    Case "0.7", ".7", 0.7
        getRawShtNo = "0087500186"
    Case "0.8", ".8", 0.8
        getRawShtNo = "0087500191"
    Case "1.0", "1", 1
        getRawShtNo = "0087500284"
    Case "1.2", 1.2
        getRawShtNo = "0087500193"
    Case "1.5", 1.5
        getRawShtNo = "0087500188"
    Case "2.0", "2", 2
        getRawShtNo = "0087500190"
    Case Else
        Debug.Print "厚度不存在:" & t
        getRawShtNo = ""
End Select

'87500182    热锌板 0.5*1250
'87500280    热锌板 0.5*1536
'87500185    热锌板 0.6*1250 不钝化首钢
'87500186    热锌板 0.7*1250
'87500282    热锌板 0.7*1400
'87500191    热锌板 0.8*1250
'87500284    热锌板 1.0*1300
'87500193    热锌板 1.2*1250
'87500188    热锌板 1.5*1250
'87500190    热锌板 2.0*1250

End Function
Private Function getRawShtDesp(partnumber As String) As String
    Select Case partnumber
        Case "0000180451"
            getRawShtDesp = "热锌板 DC51D+Z 0.5*1250"
        Case "0000180476"
            getRawShtDesp = "热锌板 DC51D+Z 0.6*1250"
        Case "0000180477"
            getRawShtDesp = "热锌板 DC51D+Z 0.7*1250"
        Case "0000180503"
            getRawShtDesp = "热锌板 DC51D+Z 0.8*1250"
        Case "0000180520"
            getRawShtDesp = "热锌板 DC51D+Z 1.0*1250"
        Case "0000180586"
            getRawShtDesp = "热锌板 DC52D+Z 1.2*1250"
        Case "0000180540"
            getRawShtDesp = "热锌板 DC51D+Z 1.5*1250"
        Case "0000180544"
            getRawShtDesp = "热锌板 DC52D+Z 2.0*1250"
        Case Else
            getRawShtDesp = ""
    End Select
End Function

Private Function CalcShtWT(r As String) As Double
Dim L As Double
Dim h As Double
Dim t As Double
Dim v As Variant

'r = Replace(r, "x", "*")
'r = Replace(r, "X", "*")

Dim uv As Variant
uv = Split(r, " ")

Dim i As Integer

For i = LBound(uv) To UBound(uv)
    If InStr(1, uv(i), "*", vbTextCompare) > 1 Then
        v = uv(i)
        Exit For
    End If
Next

If Len(v) = 0 Then
    CalcShtWT = 0
    Exit Function
End If

Dim s As Variant
s = Split(v, "*")
L = s(LBound(s))
h = s(LBound(s) + 1)
t = s(LBound(s) + 2)
Dim rez As Double
rez = L * h * t * 7.84 / 0.9 / 1000000
If rez < 0.01 Then
    CalcShtWT = 0.01
Else
    CalcShtWT = Round(rez, 2)
End If
End Function
Private Function CalcPdWT(r As String) As Double
Dim L As Double
Dim h As Double
Dim t As Double
Dim v As Variant
'r = Replace(r, "x", "*")
'r = Replace(r, "X", "*")
Dim uv As Variant
uv = Split(r, " ")

Dim i As Integer
For i = LBound(uv) To UBound(uv)
    If InStr(1, uv(i), "*", vbTextCompare) > 1 Then
        v = uv(i)
        Exit For
    End If
Next
If Len(v) = 0 Then
    CalcPdWT = 0
    Exit Function
End If
Dim s As Variant
 s = Split(v, "*")
L = s(LBound(s))
h = s(LBound(s) + 1)
t = s(LBound(s) + 2)
Dim rez As Double
rez = L * h * 2 * 0.18 / 0.9 / 1000000  '白色粉末 0.9, 特殊粉末0.5
If rez < 0.01 Then
    CalcPdWT = 0.01
Else
    CalcPdWT = Round(rez, 2)
End If
End Function
Public Sub findDespbyNo()
'    Dim uRg As Excel.Range
'    Set uRg = Excel.Selection
'
'    Dim wb As Workbook
'    Set wb = Excel.Workbooks("Selene角柜专用号申请记录.xlsx")
'
'    Dim st As Excel.Worksheet
'    Set st = wb.Sheets("专用号申请")
'
'    Dim rg As Excel.Range
'    Set rg = st.UsedRange
'
'    CELLS().Row, ColDef.UnitCol).Value
'
'    st.Cells(rg.Find(uRg.Cells(1, 1).Value).Row, ColDef.UnitCol).Value
    
End Sub
Public Sub Trans2SpotWeld()
    Dim uRg As Excel.Range
    Set uRg = Excel.Selection
    
        Call GetBomColumn
        
    Dim rowHeight As Integer
    rowHeight = uRg.Rows.Count
    
    Dim pw As Double
    
    Dim nname As String
    
    nname = InputBox("输入点焊件喷粉专用号,未喷粉专用号和名称,用逗号分隔", "输入", "0080101233,0080101234,点焊件")
    If Len(nname) = 0 Then Exit Sub
    
    Dim i As Integer
    
    For i = rowHeight To 1 Step -1

        If Len(uRg.Cells(i, ColDef.LvlCol).Value) > 0 Then
            uRg.Cells(i, ColDef.LvlCol).Value = getSubLevel(uRg.Cells(i, ColDef.LvlCol).Value)
        End If
        
        If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "DC51D+Z", vbTextCompare) > 0 Or InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "DC52D+Z", vbTextCompare) > 0 Then
            If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "已喷", vbTextCompare) > 0 Then
                uRg.Rows(i).Delete
                rowHeight = rowHeight - 1
              End If
            If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "未喷", vbTextCompare) > 0 Then
                uRg.Cells(i, ColDef.DespCol).Value = Trim(Replace(uRg.Cells(i, ColDef.DespCol).Value, "(未喷粉)", ""))
                uRg.Cells(i, ColDef.LocCol).Value = "201"
            End If
        End If
        
        If InStr(1, uRg.Cells(i, ColDef.DespCol).Value, "白色粉沫", vbTextCompare) > 0 Then
            pw = pw + CDbl(uRg.Cells(i, ColDef.QtyCol).Value)
            uRg.Rows(i).Delete
            rowHeight = rowHeight - 1
        End If

'        If i > 9999 Then
'            MsgBox "loop!"
'            Exit For
'        End If
    Next
    
    Dim vs As Variant
    vs = Split(nname, ",")
    If UBound(vs) >= 2 Then
        uRg.Rows(1).Insert
        uRg.Cells(0, ColDef.LvlCol) = getUpLevel(getUpLevel(uRg.Cells(1, ColDef.LvlCol).Value))
        uRg.Cells(0, ColDef.CodeCol) = vs(LBound(vs))
        uRg.Cells(0, ColDef.DespCol) = vs(UBound(vs)) & " (已喷粉)"
        uRg.Cells(0, ColDef.LocCol) = "201"
        uRg.Cells(0, ColDef.QtyCol) = "1"
        uRg.Cells(0, ColDef.UnitCol) = "EA"
        uRg.Cells(0, ColDef.TypeCol) = "自制"
        
        uRg.Rows(1).Insert
        uRg.Cells(0, ColDef.LvlCol) = getUpLevel(uRg.Cells(1, ColDef.LvlCol).Value)
        uRg.Cells(0, ColDef.CodeCol) = vs(LBound(vs) + 1)
        uRg.Cells(0, ColDef.DespCol) = vs(UBound(vs)) & " (未喷粉)"
        uRg.Cells(0, ColDef.LocCol) = "601"
        uRg.Cells(0, ColDef.QtyCol) = "1"
        uRg.Cells(0, ColDef.UnitCol) = "EA"
        uRg.Cells(0, ColDef.TypeCol) = "自制"
        
        rowHeight = rowHeight + 1
        uRg.Rows(rowHeight).Insert
        uRg.Cells(rowHeight, ColDef.LvlCol) = getUpLevel(uRg.Cells(1, ColDef.LvlCol).Value)
        uRg.Cells(rowHeight, ColDef.CodeCol) = "0088200085"
        uRg.Cells(rowHeight, ColDef.DespCol) = "白色粉沫 RAL9003"
        uRg.Cells(rowHeight, ColDef.LocCol) = "601"
        uRg.Cells(rowHeight, ColDef.QtyCol) = pw
        uRg.Cells(rowHeight, ColDef.UnitCol) = "公斤"
        uRg.Cells(rowHeight, ColDef.TypeCol) = "外购"
    End If
    If UBound(vs) = 1 Then
        uRg.Rows(1).Insert
        uRg.Cells(0, ColDef.LvlCol) = getUpLevel(uRg.Cells(1, ColDef.LvlCol).Value)
        uRg.Cells(0, ColDef.CodeCol) = vs(LBound(vs))
        uRg.Cells(0, ColDef.DespCol) = vs(UBound(vs))
        uRg.Cells(0, ColDef.LocCol) = "201"
        uRg.Cells(0, ColDef.QtyCol) = "1"
        uRg.Cells(0, ColDef.UnitCol) = "EA"
        uRg.Cells(0, ColDef.TypeCol) = "自制"
    End If
   
End Sub

Public Sub initExportedBom()

Dim sht As Excel.Worksheet
If Excel.ActiveSheet Is Nothing Then Exit Sub


'Application.ScreenUpdating = False   '关闭屏幕刷新
Application.EnableEvents = False '先禁止触发事件
Application.Calculation = xlCalculationManual    '手动重算
    

Set sht = Excel.ActiveSheet

    Call GetBomColumn
    
    sht.UsedRange.FormatConditions.Delete   '清楚当前颜色

Dim i As Integer

If Not isThisSheetRawBom(sht) Then Exit Sub

'删除多余列
For i = sht.UsedRange.Columns.Count To 1 Step -1
    Select Case sht.Cells(ColDef.DefRow, i).Text
    Case "备注3", "备注2", "备注1", "备注", "是否禁用", "使用状态", "使用状态", "关键件标志", "使用状态", _
        "发料仓库", "工序", "工序号", "子项类型", "坯料数", "坯料尺寸", "位置号", "损耗率(%)", "计划百分比(%)", _
        "用量", "图号", "规格型号", "直接材料", "直接人工", "变动制造费用", "固定制造费用", "委外材料费", "委外加工费", "项目编号", "特定工厂的物料状态"
        sht.Columns.Item(i).Delete
    End Select
Next

    Call GetBomColumn

If sht.Cells(2, ColDef.CodeCol).Text = sht.Cells(2, ColDef.DespCol).Text Then
    sht.Cells(2, ColDef.CodeCol).Value = ""  '解决列宽问题
    sht.Cells(1, ColDef.CodeCol).Value = ""
End If

'调整列宽和对齐方式
For i = 1 To sht.UsedRange.Columns.Count
    sht.Columns(i).AutoFit
    Select Case sht.Cells(ColDef.DefRow, i).Text
        Case "层次", "物料属性", "展开层", "对象描述"
            sht.Columns(i).HorizontalAlignment = xlLeft
        Case "子项物料代码", "单位", "单位用量", "工位", "对象标识", "组件数量(CUn)", "组件单位", "物料类型", "排序字符串"
            sht.Columns(i).HorizontalAlignment = xlCenter
    End Select
Next

'设置表格线
sht.UsedRange.Borders.LineStyle = xlContinuous
sht.UsedRange.Borders.weight = xlThin
'sht.UsedRange.Borders.ColorIndex = 1

'设置颜色
Dim fmtCds As Excel.FormatCondition
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1="".1"")")
fmtCds.Interior.Color = RGB(0, 176, 80)
fmtCds.StopIfTrue = False
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1=""..2"")")
fmtCds.Interior.Color = RGB(79, 98, 40)
fmtCds.StopIfTrue = False
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1=""...3"")")
fmtCds.Interior.Color = RGB(118, 147, 60)
fmtCds.StopIfTrue = False
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1=""....4"")")
fmtCds.Interior.Color = RGB(196, 215, 155)
fmtCds.StopIfTrue = False
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1="".....5"")")
fmtCds.Interior.Color = RGB(216, 228, 188)
fmtCds.StopIfTrue = False
Set fmtCds = sht.UsedRange.FormatConditions.Add(xlExpression, , "=($A1=""......6"")")
fmtCds.Interior.Color = RGB(235, 241, 222)
fmtCds.StopIfTrue = False

'调整用量小数位
If sht.Cells(ColDef.DefRow, ColDef.QtyCol).Text = "组件数量(CUn)" Then
    Dim hCode As String
    For i = ColDef.DefRow + 1 To sht.UsedRange.Rows.Count
        hCode = sht.Cells(i, ColDef.CodeCol).Text
        If Left(hCode, 1) = "H" Then
            sht.Cells(i, ColDef.CodeCol).Value = "'" & Right(hCode, Len(hCode) - 1)
        End If
        Select Case sht.Cells(i, ColDef.TypeCol).Text
            Case "HALB"
                sht.Cells(i, ColDef.TypeCol).Value = "自制"
            Case "ROH"
                sht.Cells(i, ColDef.TypeCol).Value = "外购"
            Case "LEER"
                sht.Cells(i, ColDef.TypeCol).Value = "特征类"
        End Select
        
        Select Case sht.Cells(i, ColDef.UnitCol).Text
            Case "KG"
                sht.Cells(i, ColDef.UnitCol).Value = "公斤"
            Case "M"
                sht.Cells(i, ColDef.UnitCol).Value = "米"
        End Select
    Next
Else
    For i = ColDef.DefRow + 1 To sht.UsedRange.Rows.Count
        If sht.Cells(i, ColDef.LvlCol).Text = "" And sht.Cells(i, ColDef.CodeCol).Text = "" Then Exit Sub
        
        
        If InStr(1, sht.Cells(i, ColDef.QtyCol).Text, ".0000", vbTextCompare) Then
                sht.Cells(i, ColDef.QtyCol).Value = Replace(sht.Cells(i, ColDef.QtyCol).Text, ".0000", "", , , vbTextCompare)
        Else
            If Len(Trim(sht.Cells(i, ColDef.QtyCol).Text)) > 0 Then
                If CDbl(sht.Cells(i, ColDef.QtyCol).Text) <> Int(CDbl(sht.Cells(i, ColDef.QtyCol).Text)) Then
                    sht.Cells(i, ColDef.QtyCol).Value = Round(CDbl(sht.Cells(i, ColDef.QtyCol).Text), 2)
                End If
            End If
        End If
    Next
End If

'开启AutoFilter
sht.Rows(ColDef.DefRow).AutoFilter

Application.EnableEvents = True '恢复触发事件
Application.Calculation = xlCalculationAutomatic    '自动重算
'Application.ScreenUpdating = True   '开启屏幕刷新

End Sub
Private Function isThisSheetRawBom(sheet As Excel.Worksheet) As Boolean
'If sheet.UsedRange.Columns.Count = 25 Then
    If ColDef.CodeCol <> 0 Then
        If ColDef.LvlCol <> 0 Then
            isThisSheetRawBom = True
            Exit Function
        Else
            Debug.Print "层次<>" & sheet.Cells(ColDef.DefRow, ColDef.LvlCol).Value
        End If
    Else
        Debug.Print "物料代码<>" & sheet.Cells(1, ColDef.CodeCol).Value
    End If
'End If
isThisSheetRawBom = False
Debug.Print "可能不是BOM"
End Function
Public Sub bom_CheckUsage()
Dim sht As Excel.Worksheet
If Excel.ActiveSheet Is Nothing Then Exit Sub
Set sht = Excel.ActiveSheet

Call GetBomColumn

Dim i As Integer
Dim calcPw As Double
Dim shtPw As Double
For i = ColDef.DefRow + 1 To sht.UsedRange.Rows.Count
    
    Select Case sht.Cells(i, ColDef.CodeCol).Text
        Case "0088200085"   '白色粉
'            Debug.Print i
            calcPw = CalcPdWT(sht.Cells(getUpLevelRow(i), ColDef.DespCol).Text)
            If Abs(sht.Cells(i, ColDef.QtyCol).Value - calcPw) > 0.1 Then
                If calcPw > 0 Then  '无法获取时不提示
                    sht.Cells(i, ColDef.QtyCol).Font.Color = vbRed
                    sht.Cells(i, "H").Value = calcPw
                End If
            End If
            'TODO 焊接件喷粉如何计算
        Case "0000180451", _
            "0000180476", _
            "0000180477", _
            "0000180503", _
            "0000180520", _
            "0000180586", _
            "0000180540", _
            "0000180544"
        shtPw = CalcShtWT(sht.Cells(getUpLevelRow(i), ColDef.DespCol).Text)
        If Abs(sht.Cells(i, ColDef.QtyCol).Value - shtPw) > 0.1 Then
            If shtPw > 0 Then   '无法获取时不提示
                sht.Cells(i, ColDef.QtyCol).Font.Color = vbRed
                sht.Cells(i, "H").Value = calcPw
            End If
        End If
        Case Else
    
    End Select
Next
End Sub
Public Sub BBBbom_Get_cutsize()


Dim uRg As Excel.Range
If Excel.Selection Is Nothing Then Exit Sub
Set uRg = Excel.Selection

Dim filename As String
filename = Trim(uRg.Cells(1, 1).Text)

Dim myWorkspace As String
myWorkspace = GetSetting("Domisoft", "Config", "SE_Working", "")

filename = filename & ".dft"
filename = myWorkspace & "\" & filename

If IsFileExists(filename) = False Then Exit Sub     '文件是否存在


Excel.Application.Cursor = xlWait '修改鼠标为等待

If seApp Is Nothing Then Call Conn2se
Dim dft As SolidEdgeDraft.DraftDocument
Set dft = seApp.Documents.Open(filename)    '打开文件

If dft.ModelLinks.Item(1).ModelDocument.Type <> igSheetMetalDocument Then
'    dft.Close false '已经打开的文件怎么办?
    Exit Sub
End If

'读取关联模型
Dim sht As SheetMetalDocument
Set sht = dft.ModelLinks.Item(1).ModelDocument

Dim fMdl As FlatPatternModel
Set fMdl = sht.FlatPatternModels.Item(1)

'获取展开尺寸
Dim L As Double
Dim W As Double
Call fMdl.GetCutSize(L, W)

Dim thk As Variant
Call sht.GetGlobalParameter(seSheetMetalGlobalMaterialThickness, thk)

'获取名称和材料
Dim seBlk As SolidEdgeDraft.BlockOccurrence
Dim seLbs As SolidEdgeDraft.BlockLabelOccurrences
Dim isLegacyDoc As Boolean
Dim i As Integer

For i = 1 To dft.Sheets.Item(1).BlockOccurrences.Count
    Select Case dft.Sheets.Item(1).BlockOccurrences.Item(i).Block.Name
        Case "Title"
            Set seBlk = dft.ActiveSheet.BlockOccurrences.Item(i)
            Set seLbs = seBlk.BlockLabelOccurrences
            isLegacyDoc = False
        Case "Title-SRDC_V1"
            Set seBlk = dft.ActiveSheet.BlockOccurrences.Item(i)
            Set seLbs = seBlk.BlockLabelOccurrences
            isLegacyDoc = True
    End Select
Next
Dim name_cn As String
Dim material As String

name_cn = IIf(isLegacyDoc, seLbs.Item(LegacySeDftBlockId.name_cn).Value, seLbs.Item(SeDftBlockId.name_cn).Value)
material = IIf(isLegacyDoc, seLbs.Item(LegacySeDftBlockId.material).Value, seLbs.Item(SeDftBlockId.material).Value)

Excel.Application.Cursor = xlDefault '恢复鼠标

'导出字符串
Dim outstr As String
outstr = name_cn & " " & material & " " & Round((L * 1000), 1) & "*" & Round((W * 1000), 1) & "*" & Format(CDbl(thk) * 1000, "0.0")

Dim a As String
a = InputBox("点击确定将结果复制到剪贴板.", "展开尺寸", outstr)

If a = outstr Then

Dim MyData As DataObject
Set MyData = New DataObject
MyData.SetText outstr
MyData.PutInClipboard
Set MyData = Nothing

End If

dft.Close False
Set dft = Nothing
End Sub

Public Sub bom_Get_cutsize() 'batchExpand() '借用按钮临时改了名字

Application.ScreenUpdating = False   '关闭屏幕刷新
Application.EnableEvents = False '先禁止触发事件
Application.Calculation = xlCalculationManual    '手动重算


Dim aht As Excel.Worksheet
Set aht = Excel.ActiveSheet

Call GetBomColumn

Dim pNumber As String

Dim i As Integer
i = aht.UsedRange.Rows.Count

Do Until i = 6
    pNumber = aht.Cells(i, ColDef.CodeCol).Text
    If Len(Trim(pNumber)) > 0 Then
        If BOMExpandable(aht.Rows(i)) Then
            If IsBOMExpanded(aht.Rows(i)) = False Then
                Call BOMExpand(aht.Rows(i))
            End If
        End If
    End If
    i = i - 1
Loop

Application.EnableEvents = True '恢复触发事件
Application.Calculation = xlCalculationAutomatic    '自动重算
Application.ScreenUpdating = True   '开启屏幕刷新

End Sub
Private Sub BOMExpand(erow As Excel.Range)
'Const BOMSource_path = "Y:\A02-Project\B07-Project_2015\P1004_E6 semi Multideck\03-engineering\313-BOM\02-Meat case\2018.05.04\整合SV_BOM_DATA.xlsx"

Dim Cht As Excel.Worksheet
Set Cht = Excel.Workbooks("整合SV_BOM_DATA.xlsx").Sheets(1)

Dim aht As Excel.Worksheet
Set aht = erow.Parent

Dim r As Integer
r = erow.row

Dim pNumber As String
pNumber = aht.Cells(r, ColDef.CodeCol).Text

Dim cr As Integer

Dim i As Integer
For i = 2 To Cht.UsedRange.Rows.Count
    If pNumber = Cht.Cells(i, 2).Text Then
        cr = i
        Exit For
    End If
Next

If cr = 0 Then Exit Sub

Dim crb As Integer
Dim clt As Integer
clt = getLvlNum(Cht.Cells(cr, 1))

For i = cr + 1 To Cht.UsedRange.Rows.Count

    If i = Cht.UsedRange.Rows.Count Then
            crb = i
            Exit For
    Else
        If clt >= getLvlNum(Cht.Cells(i, 1)) Then    '>= 解决BUG
            crb = i - 1
            Exit For
        End If
    End If
Next

'If crb = 0 Then Exit Sub

Dim cRng As Excel.Range

Set cRng = Cht.Rows(CStr(cr + 1) & ":" & CStr(crb))

cRng.Copy

aht.Rows(r + 1).Insert Shift:=xlDown


End Sub
Private Function BOMExpandable(r As Excel.Range) As Boolean  '判断所选行是否可以展开
    Dim mtype As String
    mtype = Mid(r.Cells(1, ColDef.CodeCol).Text, 4, 2)
    Select Case mtype
        Case "08", "01", "09"
                Dim ptype As String
                ptype = r.Cells(1, ColDef.TypeCol).Text
                Select Case ptype
                    Case "外购", "外购件", "外购类", "虚拟", "虚拟件", "虚拟类"
                        BOMExpandable = False
                    Case Else
                        BOMExpandable = True
                End Select
        Case Else
            BOMExpandable = False
    End Select

End Function
Private Function IsBOMExpanded(r As Excel.Range) As Boolean '判断所选行是否已经展开
    Dim rr As Range
    Set rr = r.Parent.Rows(r.row + 1)
    If getLvlNum(r.Cells(1, 1).Text) < getLvlNum(rr.Cells(1, 1).Text) Then
        IsBOMExpanded = True
    Else
        IsBOMExpanded = False
    End If
End Function
Public Sub demoteAll()
Dim asht As Excel.Worksheet
Set asht = Excel.ActiveSheet
Dim i As Integer
For i = 8 To 56
    asht.Cells(i, 1).Value = getSubLevel(CStr(asht.Cells(i, 1).Text))
Next
End Sub
Public Sub CheckBOM()

Dim sht As Excel.Worksheet

If Excel.ActiveSheet Is Nothing Then Exit Sub
Set sht = Excel.ActiveSheet

Call GetBomColumn

Dim i As Integer
Dim j As Integer
Dim haveTZL As Boolean


If Not isThisSheetRawBom(sht) Then Exit Sub

For i = ColDef.DefRow + 1 To sht.UsedRange.Rows.Count
    Select Case sht.Cells(i, ColDef.TypeCol).Text
    Case "特征类"
        If sht.Cells(getUpLevelRow(i), ColDef.TypeCol).Text <> "配置类" Then
            Debug.Print "第" & i & "行, 存在错误: 特征类上一级不是配置类"
        End If
        
    Case "配置类"
        j = i
        haveTZL = False
        
        Do Until sht.Cells(j, ColDef.LvlCol).Text = sht.Cells(i, ColDef.LvlCol).Text
            j = j + 1
            If j > sht.UsedRange.Rows.Count Then Exit Sub
            If sht.Cells(j, ColDef.TypeCol).Text = "特征类" Then
                haveTZL = True
            End If
        Loop
        
        If haveTZL = False Then
            Debug.Print "第" & i & "行, 存在错误: 配置类下一级无特征类"
        End If
    End Select
Next

End Sub
Public Function getWeight(str As String)
    If InStr(1, str, "未喷粉", vbTextCompare) > 1 Then getWeight = CalcShtWT(str)
    If InStr(1, str, "已喷粉", vbTextCompare) > 1 Then getWeight = CalcPdWT(str)
    If InStr(1, str, "热锌板", vbTextCompare) > 1 And InStr(1, str, "已喷粉", vbTextCompare) < 1 Then getWeight = CalcShtWT(str)
End Function


