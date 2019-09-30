Attribute VB_Name = "Model_MaterialList"
Public Sub fillCutsizeTable()
Const L = "G"
Const W = "H"
Const t = "I"
Dim s As String
Dim ss As String
Dim u As Variant
Dim UU As Variant
Dim js
Set js = CreateObject("msscriptcontrol.scriptcontrol")
js.Language = "javascript"
js.addcode "function aa(bb){js=bb.split(',');js.sort(function(a,b){return a-b;});js.reverse();return js;}"

For i = 2 To ActiveSheet.UsedRange.Rows.Count
    s = Cells(i, "B").Value
    ss = ""
    u = Split(s, " ")
    For j = LBound(u) To UBound(u)
        If InStr(1, u(j), "*") > 1 Then
            ss = u(j)
            Exit For
        End If
    Next
    If ss = "" Then GoTo nSkip
    ss = Replace(ss, "*", ",", , , vbTextCompare)
    ss = js.Eval("aa('" & ss & "')")
    UU = Split(ss, ",")

    Cells(i, L) = UU(0)
    Cells(i, W) = UU(1)
    Cells(i, t) = UU(2)
nSkip:
Next
End Sub
Public Sub fillSimpleName()
Dim s As String
Dim ss As String

For i = 2 To ActiveSheet.UsedRange.Rows.Count
    s = Cells(i, "B").Value
    ss = ""
    u = Split(s, " ")
    For j = LBound(u) To UBound(u)
        If InStr(1, u(j), "*") > 1 Then GoTo mSkip
        If InStr(1, u(j), "ÈÈÐ¿") > 0 Then GoTo mSkip
        If InStr(1, u(j), "Åç") > 0 Then GoTo mSkip
        If InStr(1, u(j), "SV") > 0 Then GoTo mSkip
        If InStr(1, u(j), "E6") > 0 Then GoTo mSkip
        If InStr(1, u(j), "EU") > 0 Then GoTo mSkip
        If InStr(1, u(j), "LU") > 0 Then GoTo mSkip
        If Application.WorksheetFunction.IsNumber(u(j)) Then GoTo mSkip
        If Len(ss) = 0 Then ss = u(j)
mSkip:
    Next
    If ss = "" Then GoTo nSkip
    ss = Replace(ss, "SD", "", , , vbTextCompare)
    Cells(i, "C") = ss
nSkip:
Next
End Sub
Private Function biggerNumber(a, b)
    If Val(a) > Val(b) Then
        biggerNumber = a
    Else
        biggerNumber = b
    End If
End Function
Private Function smallerNumber(a, b)
    If Val(a) < Val(b) Then
        smallerNumber = a
    Else
        smallerNumber = b
    End If
End Function
