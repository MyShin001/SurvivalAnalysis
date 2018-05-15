Attribute VB_Name = "Mat"
Option Explicit

Sub printMat(name, mat)
    Dim i As Long, j As Long
    
    Debug.Print ("-----------------")
    Debug.Print (name)
    For i = 1 To UBound(mat, 1)
        For j = 1 To UBound(mat, 2)
            Debug.Print (mat(i, j));
        Next
        Debug.Print ("")
    Next
    Debug.Print ("")
End Sub

Function PlusMat(m1, m2)
    Dim ans()
    Dim i As Long, j As Long
    
    If (UBound(m1, 1) <> UBound(m2, 1) Or UBound(m1, 2) <> UBound(m2, 2)) Then
        PlusMat = -1
        Exit Function
    End If

    ReDim ans(1 To UBound(m1, 1), 1 To UBound(m1, 2))
    For i = 1 To UBound(ans, 1)
        For j = 1 To UBound(ans, 2)
            ans(i, j) = m1(i, j) + m2(i, j)
        Next
    Next
    PlusMat = ans
End Function

Function MinusMat(m1, m2)
    Dim ans()
    Dim i As Long, j As Long
    
    If (UBound(m1, 1) <> UBound(m2, 1) Or UBound(m1, 2) <> UBound(m2, 2)) Then
        MinusMat = -1
        Exit Function
    End If
    
    ReDim ans(1 To UBound(m1, 1), 1 To UBound(m1, 2))
    For i = 1 To UBound(ans, 1)
        For j = 1 To UBound(ans, 2)
            ans(i, j) = m1(i, j) - m2(i, j)
        Next
    Next
    MinusMat = ans
End Function

Function MulMat(m1, m2)
    Dim ans()
    Dim i As Long, j As Long, k As Long
    
    If UBound(m1, 2) <> UBound(m2, 1) Then
        MulMat = -1
        Exit Function
    End If
    
    ReDim ans(1 To UBound(m1, 1), 1 To UBound(m2, 2))
    For i = 1 To UBound(ans, 1)
        For j = 1 To UBound(ans, 2)
            ans(i, j) = 0
            For k = 1 To UBound(m1, 2)
                ans(i, j) = ans(i, j) + m1(i, k) * m2(k, j)
            Next
        Next
    Next
    MulMat = ans
End Function

Function MulValMat(val, mat)
    Dim ans()
    Dim i As Long, j As Long
    
    ReDim ans(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    For i = 1 To UBound(ans, 1)
        For j = 1 To UBound(ans, 2)
            ans(i, j) = val * mat(i, j)
        Next
    Next
    MulValMat = ans
End Function

Function TMat(mat)
    Dim ans()
    Dim i As Long, j As Long
    
    ReDim ans(1 To UBound(mat, 2), 1 To UBound(mat, 1))
    For i = 1 To UBound(ans, 1)
        For j = 1 To UBound(ans, 2)
            ans(i, j) = mat(j, i)
        Next
    Next
    TMat = ans
End Function

Function InvMat(mat)
    Dim org(), inv()
    Dim UM As Long, buf As Double
    Dim i As Long, j As Long, k As Long
    
    If UBound(mat, 1) <> UBound(mat, 2) Then
        InvMat = -1
        Exit Function
    End If
    
    UM = UBound(mat, 1)
    org = mat
    ReDim inv(1 To UM, 1 To UM)
    For i = 1 To UBound(inv, 1)
        inv(i, i) = 1
    Next
    
    For i = 1 To UM
        buf = 1 / org(i, i)
        For j = 1 To UM
            org(i, j) = org(i, j) * buf
            inv(i, j) = inv(i, j) * buf
        Next
        
        For j = 1 To UM
            If i <> j Then
                buf = org(j, i)
                For k = 1 To UM
                    org(j, k) = org(j, k) - org(i, k) * buf
                    inv(j, k) = inv(j, k) - inv(i, k) * buf
                Next
            End If
        Next
    Next
    InvMat = inv
End Function
