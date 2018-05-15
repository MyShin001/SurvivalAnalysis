Attribute VB_Name = "MatEX"
Option Explicit

Public LDS() 'Value
Public SGMW() 'Value
Public SGMWX() 'Vector
Public SGMWXX() 'Matrix

Sub setLDS(data(), dataInf(), LDSMax, xkey)
    Dim row As Long, vec()
    Dim i As Long, j As Long, k As Long, L As Long
    
    ReDim LDS(1 To dataInf(1), 1 To 3)
    
    row = 1
    For i = 2 To dataInf(1)
        j = i
        Do While data(j, 1) = data(j + 1, 1)
            If data(j, 2) = 1 Then LDS(row, 2) = LDS(row, 2) + 1
            j = j + 1
        Loop
        If data(j, 2) = 1 Then LDS(row, 2) = LDS(row, 2) + 1
        
        LDS(row, 1) = j - i + 1
        
        If IsArray(xkey) Then: ReDim vec(1 To UBound(xkey), 1 To 1)
        If Not IsArray(xkey) Then: ReDim vec(1 To dataInf(2) - 2, 1 To 1)
        
        For k = i To j
            If data(k, 2) = 1 Then 'event = 1
                If IsArray(xkey) Then
                    For L = 1 To UBound(xkey)
                        vec(L, 1) = vec(L, 1) + data(k, xkey(L))
                    Next
                Else
                    For L = 1 To dataInf(2) - 2
                        vec(L, 1) = vec(L, 1) + data(k, L + 2)
                    Next
                End If
            End If
        Next
        LDS(row, 3) = vec
        
        row = row + 1
        i = j
    Next
    LDSMax = row - 1
End Sub

Sub setSGMs(data(), dataInf(), beta(), xkey)
    Dim x(), tx(), W, wx(), wxx()
    Dim i As Long, j As Long
    ReDim SGMW(2 To dataInf(1))
    ReDim SGMWX(2 To dataInf(1))
    ReDim SGMWXX(2 To dataInf(1))
    
    If IsArray(xkey) Then: ReDim x(1 To UBound(xkey), 1 To 1)
    If Not IsArray(xkey) Then: ReDim x(1 To dataInf(2) - 2, 1 To 1)
    
    For i = dataInf(1) To 2 Step -1
'x, tx, w, wx, wxx
        W = 0
        If IsArray(xkey) Then
            For j = 1 To UBound(xkey)
                x(j, 1) = data(i, xkey(j))
                W = W + beta(j, 1) * x(j, 1)
            Next
        Else
            For j = 1 To UBound(x)
                x(j, 1) = data(i, j + 2)
                W = W + beta(j, 1) * x(j, 1)
            Next
        End If
        tx = TMat(x)
        W = Exp(W)
        wx = MulValMat(W, x)
        wxx = MulMat(wx, tx)
        
'SGMs
        Select Case i
        Case dataInf(1)
            SGMW(i) = W
            SGMWX(i) = wx
            SGMWXX(i) = wxx
        Case Else
            SGMW(i) = SGMW(i + 1) + W
            SGMWX(i) = PlusMat(SGMWX(i + 1), wx)
            SGMWXX(i) = PlusMat(SGMWXX(i + 1), wxx)
        End Select
    Next
End Sub

Function UpdateBeta(LDSMax, invI, beta)
    Dim at(), dat(), St(), i1(), i2(), U(), Ut(), II(), IIt()
    Dim row As Long
    Dim i As Long, j As Long
    
    ReDim U(1 To UBound(beta), 1 To 1)
    ReDim II(1 To UBound(beta), 1 To UBound(beta))
    
    row = 2
    For i = 1 To LDSMax
        at = MulValMat(1 / SGMW(row), SGMWX(row))
        dat = MulValMat(LDS(i, 2), at)
        Ut = MinusMat(LDS(i, 3), dat)
        U = PlusMat(U, Ut)

        i1 = MulValMat(1 / SGMW(row), SGMWXX(row))
        i2 = MulMat(at, TMat(at))
        IIt = MulValMat(LDS(i, 2), MinusMat(i1, i2))
        II = PlusMat(II, IIt)
        
        row = row + LDS(i, 1)
    Next
    invI = InvMat(II)
    UpdateBeta = PlusMat(beta, MulMat(invI, U))
End Function

Function checkBeta(oldBeta(), newBeta())
    Dim i As Long
    For i = 1 To UBound(newBeta)
        If Abs(oldBeta(i, 1) - newBeta(i, 1)) > 1E-06 Then
            checkBeta = False
            Exit Function
        End If
    Next
    checkBeta = True
End Function

Sub calBeta(data(), dataInf(), xkey, beta, invI)
    Dim oldBeta(), newBeta()
    Dim LDSMax As Long, matSize As Long
    Dim i As Long
    
    If IsArray(xkey) Then: ReDim oldBeta(1 To UBound(xkey), 1 To 1)
    If Not IsArray(xkey) Then: ReDim oldBeta(1 To dataInf(2) - 2, 1 To 1)
    
    Call setLDS(data, dataInf, LDSMax, xkey)
    For i = 1 To 10
        Call setSGMs(data, dataInf, oldBeta, xkey)
        newBeta = UpdateBeta(LDSMax, invI, oldBeta)
        If checkBeta(oldBeta, newBeta) Then Exit For
        oldBeta = newBeta
    Next
    beta = newBeta
    Call printMat("beta", beta)
End Sub

Sub calXX(beta, invI, XX())
    Dim i As Long
    ReDim XX(1 To UBound(beta))
    Debug.Print ("-----------------")
    Debug.Print ("X^2")
    For i = 1 To UBound(XX)
        XX(i) = beta(i, 1) * beta(i, 1) / invI(i, i)
        Debug.Print (XX(i))
    Next
    Debug.Print ("")
End Sub
