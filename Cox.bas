Attribute VB_Name = "Cox"
Option Explicit

Sub BasicData(data(), dataInf(), result1(), result2())
    Dim dNum As Long, ave As Double, ave2 As Double
    Dim s2 As Double, max As Double, min As Double
    Dim rowData()
    Dim i As Long, j As Long
    
    'result1
    ReDim result1(1 To 3, 1 To 5)
    For i = 2 To dataInf(1)
        If data(i, 2) = 1 Then dNum = dNum + 1
    Next
    rowData = Array("有効サンプル数", "発生数", "打ち切り数")
    Call SAV.AddRow(result1, rowData, 1, 1)
    rowData = Array(dataInf(1) - 1, dNum, dataInf(1) - dNum - 1)
    Call SAV.AddRow(result1, rowData, 2, 1)
    
    'result2
    ReDim result2(1 To dataInf(2), 1 To 10)
    rowData = Array("共変数", "平均", "分散", "標準偏差", "最大値", "最小値")
    Call SAV.AddRow(result2, rowData, 1, 1)
    For i = 3 To dataInf(2)
        ave = 0
        ave2 = 0
        For j = 2 To dataInf(1)
            If j = 2 Then
                max = data(j, i)
                min = data(j, i)
            End If
            
            ave = ave + data(j, i)
            ave2 = ave2 + (data(j, i) * data(j, i))
            If max < data(j, i) Then max = data(j, i)
            If min > data(j, i) Then min = data(j, i)
        Next
        ave = ave / (dataInf(1) - 1)
        ave2 = ave2 / (dataInf(1) - 1)
        s2 = ave2 - (ave * ave)
        rowData = Array(data(1, i), ave, s2, Sqr(s2), max, min)
        Call SAV.AddRow(result2, rowData, i - 1, 1)
    Next
End Sub

Sub MainTable(res(), data(), beta(), invI(), XX())
    Dim rowData(), P As Double, i As Long
    ReDim res(1 To UBound(beta) + 2, 1 To 10)
    
    rowData = Array("共変数", "係数", "標準誤差", "カイ二乗値", "P値", "ハザード比")
    Call SAV.AddRow(res, rowData, 1, 1)
    For i = 1 To UBound(beta)
        P = 1 - WorksheetFunction.ChiSq_Dist(XX(i), 1, True)
        rowData = Array(data(1, i + 2), beta(i, 1), Sqr(invI(i, i)), XX(i), P, Exp(beta(i, 1)))
        Call SAV.AddRow(res, rowData, i + 1, 1)
    Next
End Sub

Sub simple(data(), dataInf(), res())
    Dim beta(), invI(), XX()
    Dim LDSMax As Long
    
    Call calBeta(data, dataInf, 0, beta, invI)
    Call calXX(beta, invI, XX)
    Call MainTable(res, data, beta, invI, XX)
End Sub

Function selectX(data(), dataInf(), xkey(), Pin, Pout, SlctTable(), BetaTable())
    Dim LDSMax As Long
    Dim beta(), invI(), XX()
    Dim state() As Boolean, keyNum As Long
    Dim minPi, maxPi
    Dim step As Long, endFlag As Boolean
    Dim rRow As Long, rData()
    
    ReDim state(1 To dataInf(2) - 2)
    step = 1
    rRow = 1
    ReDim SlctTable(1 To dataInf(2) * dataInf(2) * 2, 1 To 10)
    rData = Array("", "", "スコア検定", "Wald検定")
    Call SAV.AddRow(SlctTable, rData, rRow, 1)
    rData = Array("ステップ", "共変数", "投入のΧ^2", "除去のΧ^2", "P値")
    Call SAV.AddRow(SlctTable, rData, rRow, 1)
    
    Do While True
        SlctTable(rRow, 1) = "ステップ" & step
'***投入***
        If step = 1 Then
            minPi = AddX(data, dataInf, state, 0, SlctTable, rRow)
        Else
            minPi = AddX(data, dataInf, state, beta, SlctTable, rRow)
        End If
        If minPi(1) <= Pin Then
            state(minPi(2)) = True
            endFlag = False
        Else
            If endFlag Then Exit Do Else endFlag = True
        End If
'***除去***
        maxPi = RemoveX(data, dataInf, state, beta, invI, XX, SlctTable, rRow)
        If maxPi(1) >= Pout Then
            state(maxPi(2)) = False
            endFlag = False
            'xkey, beta
            Call GetXKey(state, keyNum, xkey)
            Call calBeta(data, dataInf, xkey, beta, invI)
        Else
            If endFlag Then Exit Do Else endFlag = True
        End If
        
        If step = dataInf(2) * dataInf(2) * 2 Then Exit Do
        step = step + 1
    Loop
    
    If IsArrayEx(beta) = 1 Then
        Call MainTable(BetaTable, data, beta, invI, XX)
    Else
        ReDim BetaTable(1 To 1, 1 To 1)
        BetaTable(1, 1) = "条件を満たす共変数はありません"
    End If
    
    selectX = 1
    Exit Function
ERR:
    MsgBox ("エラーC：入力シートを確認してください。")
    selectX = -1
End Function

Sub GetXKey(state, keyNum, xkey)
    keyNum = 0
    Dim i As Long
    For i = 1 To UBound(state)
        If state(i) Then
            keyNum = keyNum + 1
            ReDim Preserve xkey(1 To keyNum)
            xkey(keyNum) = i + 2
        End If
    Next
End Sub

Function AddX(data(), dataInf(), state, bestBeta, result(), rRow)
    Dim tmpBeta(), invI()
    Dim xkey() As Long, keyNum As Long
    Dim XX(), P As Double, minPi(1 To 2) As Double
    Dim LDSMax As Long, rData()
    Dim i As Long, j As Long
    
    minPi(1) = 1
    For i = 1 To UBound(state)
        If state(i) = False Then
'tmpBeta
            If IsArray(bestBeta) Then
                ReDim tmpBeta(1 To UBound(bestBeta, 1) + 1, 1 To 1)
                For j = 1 To UBound(bestBeta, 1)
                    tmpBeta(j, 1) = bestBeta(j, 1)
                Next
            Else
                ReDim tmpBeta(1 To 1, 1 To 1)
            End If
'xkey
            Call GetXKey(state, keyNum, xkey)
            ReDim Preserve xkey(1 To keyNum + 1)
            xkey(keyNum + 1) = i + 2
'beta
            Call setLDS(data, dataInf, LDSMax, xkey)
            Call setSGMs(data, dataInf, tmpBeta, xkey)
            tmpBeta = UpdateBeta(LDSMax, invI, tmpBeta)
            Call printMat("ADD", tmpBeta)
'XX, P
            Call calXX(tmpBeta, invI, XX)
            P = 1 - WorksheetFunction.ChiSq_Dist(XX(UBound(XX)), 1, True)
            If minPi(1) > P Then
                minPi(1) = P: minPi(2) = i
            End If
'Add
            rData = Array(data(1, i + 2), XX(UBound(XX)), "", P)
            Call SAV.AddRow(result, rData, rRow, 2)
        End If
    Next
    AddX = minPi
End Function

Function RemoveX(data(), dataInf(), state, beta(), invI(), XX(), result(), rRow)
    Dim xkey() As Long, keyNum As Long
    Dim P As Double, maxPi(1 To 2) As Double
    Dim LDSMax As Long, rData()
    Dim i As Long, j As Long
        
'xkey, beta, XX, P, Add
    Call GetXKey(state, keyNum, xkey)
    If keyNum <> 0 Then
        Call calBeta(data, dataInf, xkey, beta, invI)
        Call calXX(beta, invI, XX)
        j = 1
        For i = 1 To UBound(state)
            If state(i) = True Then
                P = 1 - WorksheetFunction.ChiSq_Dist(XX(j), 1, True)
                If maxPi(1) < P Then
                    maxPi(1) = P: maxPi(2) = i
                End If
                rData = Array(data(1, i + 2), "", XX(j), P)
                Call SAV.AddRow(result, rData, rRow, 2)
                j = j + 1
            End If
        Next
    End If
    RemoveX = maxPi
End Function
