Attribute VB_Name = "Model_MaterialList"
Public Sub fillCutsizeTable()
Const L = "G"
Const W = "H"
Const T = "I"
Dim s As String
Dim ss As String
Dim u As Variant
Dim UU As Variant

For i = 2 To ActiveSheet.UsedRange.Rows.Count
    s = Cells(i, "C").Value
    ss = ""
    u = Split(s, " ")
    For j = LBound(u) To UBound(u)
        If InStr(1, u(j), "*") > 1 Then
            ss = u(j)
            Exit For
        End If
    Next
    If ss = "" Then GoTo nSkip
    UU = Split(ss, "*")
    
    Cells(i, L) = biggerNumber(UU(LBound(UU)), UU(LBound(UU) + 1))
    Cells(i, W) = smallerNumber(UU(LBound(UU)), UU(LBound(UU) + 1))
    Cells(i, T) = UU(UBound(UU))
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
