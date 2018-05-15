Attribute VB_Name = "KMLR"
Option Explicit

Function BasicArray(data1(), data2(), step)
    Dim dataInf1() As Long, dataInf2() As Long
    Dim res(), row As Long
    Dim time As Double, tMax As Double
    Dim L(2) As Long, D(2) As Long, W(2) As Long
    Dim i As Long, j As Long
    
    Call SAV.GetArrayInfo(data1, dataInf1)
    Call SAV.GetArrayInfo(data2, dataInf2)
    ReDim res(1 To dataInf1(1) + dataInf2(1), 1 To 10): res(1, 1) = data1(1, 1)
    If step <> 0 Then
        For i = 2 To dataInf1(1)
            data1(i, 1) = WorksheetFunction.RoundUp(data1(i, 1) / step, 0)
        Next
        For i = 2 To dataInf2(1)
            data2(i, 1) = WorksheetFunction.RoundUp(data2(i, 1) / step, 0)
        Next
    End If
    
    tMax = data1(dataInf1(1), 1)
    If tMax < data2(dataInf2(1), 1) Then tMax = data2(dataInf2(1), 1)
    L(1) = dataInf1(1) - 1: L(2) = dataInf2(1) - 1
    row = 2: i = 2: j = 2
    
    Do While time < tMax
        time = IIf(data1(i, 1) < data2(j, 1), data1(i, 1), data2(j, 1))
        If i > dataInf1(1) Then time = data2(j, 1)
        If j > dataInf2(1) Then time = data1(i, 1)
        L(1) = L(1) - D(1) - W(1): D(1) = 0: W(1) = 0
        L(2) = L(2) - D(2) - W(2): D(2) = 0: W(2) = 0
        
        Do While data1(i, 1) = time
            Select Case data1(i, 2)
                Case 0: W(1) = W(1) + 1
                Case 1: D(1) = D(1) + 1
            End Select
            i = i + 1
            If i > dataInf1(1) Then Exit Do
        Loop
        Do While data2(j, 1) = time
            Select Case data2(j, 2)
                Case 0: W(2) = W(2) + 1
                Case 1: D(2) = D(2) + 1
            End Select
            j = j + 1
            If j > dataInf2(1) Then Exit Do
        Loop
        
        Call SAV.AddRow(res, Array(time, L(1), D(1), W(1), L(2), D(2), W(2)), row, 1)
    Loop
    BasicArray = res
End Function

Function KMLR(data, aName, bName, step, res1(), res2(), res3(), res4()) As Long
'    On Error GoTo ERR
    
'Declare
    Dim dataInf() As Long
    Dim row1 As Long, row2 As Long, row3 As Long, row4 As Long
    Dim rcData As Variant
        
    Dim start As Long, fin As Long               '開始/終了
    Dim L(3) As Long, D(3) As Long, W(3) As Long, N(3) As Long
    Dim DSum(2) As Long, WSum(2) As Long
    Dim inc(2) As Double                            '期間発生率
    Dim S(2) As Double, SE_pre(2) As Double, SE(2) As Double, H(2) As Variant
    Dim ave(2) As Double, med(2) As Variant    '平均/メディアン生存時間
    Dim TR(2, 2) As Double, a As Double, b As Double    '95%信頼区間(下/上)
    Dim E(2) As Double, V As Double, k As Double
    Dim XX_PP, P_PP, five_PP As String, one_PP As String
    Dim XX_CMH, P_CMH, five_CMH As String, one_CMH As String
    Dim i As Long
    
'Initialize
    Call SAV.GetArrayInfo(data, dataInf)
    ReDim res1(1 To 5, 1 To 10)
    ReDim res2(1 To 5, 1 To 10)
    ReDim res3(1 To dataInf(1) + 5, 1 To 30)
    ReDim res4(1 To dataInf(1) * 2 + 5, 1 To 5)
    row1 = 1: row2 = 1: row3 = 1: row4 = 1
    S(1) = 1: S(2) = 1
    H(1) = 0: H(2) = 0
    med(1) = "-": med(2) = "-"
    XX_PP = "-": P_PP = "-"
    XX_CMH = "-": P_CMH = "-"
    
'Table Lable
    rcData = Array("データ", "総サンプル数", "総発生数", "総打ち切り数", _
                     "平均生存時間", "メディアン生存時間")
    Call SAV.AddRow(res1, rcData, row1, 1)
    rcData = Array("検定", "手法", "カイ二乗値", "P値", "1%検定", "5%検定")
    Call SAV.AddRow(res2, rcData, row2, 1)
    rcData = Array("", "", "", _
                    aName, "", "", "", "", "", "", "", "", _
                    bName, "", "", "", "", "", "", "", "", _
                    "2群合計", "", "")
    Call SAV.AddRow(res3, rcData, row3, 1)
    rcData = Array("開始時点", "終了時点", "期間間隔", _
    "生存数", "発生数", "打ち切り数", "期間発生率", "累積生存率", _
    "標準誤差", "累積ハザード関数", "95%信頼区間(下限)", "95%信頼区間(上限)", _
    "生存数", "発生数", "打ち切り数", "期間発生率", "累積生存率", _
    "標準誤差", "累積ハザード関数", "95%信頼区間(下限)", "95%信頼区間(上限)", _
    "生存数", "発生数", "打ち切り数")
    Call SAV.AddRow(res3, rcData, row3, 1)
    rcData = Array("", "累積生存率", "", "累積ハザード関数", "")
    Call SAV.AddRow(res4, rcData, row4, 1)
    rcData = Array(data(1, 1), aName, bName, aName, bName)
    Call SAV.AddRow(res4, rcData, row4, 1)
    
'Process
    For i = 2 To dataInf(1)
        'Before
        start = fin: fin = data(i, 1)
        L(1) = data(i, 2): D(1) = data(i, 3): W(1) = data(i, 4)
        L(2) = data(i, 5): D(2) = data(i, 6): W(2) = data(i, 7)
        L(3) = L(1) + L(2): D(3) = D(1) + D(2): W(3) = W(1) + W(2)
        N(1) = L(1) - D(1): N(2) = L(2) - D(2): N(3) = L(3) - D(3):
        DSum(1) = DSum(1) + D(1): DSum(2) = DSum(2) + D(2)
        WSum(1) = WSum(1) + W(1): WSum(2) = WSum(2) + W(2)
        If L(1) = 0 Then inc(1) = 0 Else inc(1) = D(1) / L(1)
        If L(2) = 0 Then inc(2) = 0 Else inc(2) = D(2) / L(2)
        If L(1) <> 0 Then ave(1) = ave(1) + S(1) * (fin - start)
        If L(2) <> 0 Then ave(2) = ave(2) + S(2) * (fin - start)
        If S(1) <= 0.5 And med(1) = "-" Then med(1) = start
        If S(2) <= 0.5 And med(2) = "-" Then med(2) = start
        E(1) = E(1) + (D(3) * L(1) / L(3))
        E(2) = E(2) + (D(3) * L(2) / L(3))
        If L(3) <> 1 Then
            V = V + (N(3) * D(3) / (L(3) ^ 2) * L(1) * L(2) / (L(3) - 1))
            k = k + (N(1) - N(3) * L(1) / L(3))
        End If
        
        'Add
        rcData = Array(start, fin, fin - start, _
                    L(1), D(1), W(1), inc(1), S(1), SE(1), H(1), TR(1, 1), TR(1, 2), _
                    L(2), D(2), W(2), inc(2), S(2), SE(2), H(2), TR(2, 1), TR(2, 2), _
                    L(3), D(3), W(3))
        Call SAV.AddRow(res3, rcData, row3, 1)
        rcData = Array(start, S(1), S(2), H(1), H(2))
        Call SAV.AddRow(res4, rcData, row4, 1)
        rcData = Array(fin, S(1), S(2), H(1), H(2))
        Call SAV.AddRow(res4, rcData, row4, 1)
        
        'After
        S(1) = S(1) * (1 - inc(1))
        S(2) = S(2) * (1 - inc(2))
        If (L(1) - D(1)) <> 0 Then
            SE_pre(1) = SE_pre(1) + inc(1) / (L(1) - D(1))
            SE(1) = S(1) * Sqr(SE_pre(1))
        End If
        If (L(2) - D(2)) <> 0 Then
            SE_pre(2) = SE_pre(2) + inc(2) / (L(2) - D(2))
            SE(2) = S(2) * Sqr(SE_pre(2))
        End If
        H(1) = IIf(S(1) = 0, "", -Log(S(1)))
        H(2) = IIf(S(2) = 0, "", -Log(S(2)))
        If S(1) <> 1 And S(1) <> 0 Then
            a = Exp((-1.96 * SE(1)) / (S(1) * Log(S(1))))
            b = Exp((1.96 * SE(1)) / (S(1) * Log(S(1))))
            TR(1, 1) = S(1) ^ a: TR(1, 2) = S(1) ^ b
        Else
            TR(1, 1) = 0: TR(1, 2) = 0
        End If
        If S(2) <> 1 And S(2) <> 0 Then
            a = Exp((-1.96 * SE(2)) / (S(2) * Log(S(2))))
            b = Exp((1.96 * SE(2)) / (S(2) * Log(S(2))))
            TR(2, 1) = S(2) ^ a: TR(2, 2) = S(2) ^ b
        Else
            TR(2, 1) = 0: TR(2, 2) = 0
        End If
    Next
    
    If E(1) * E(2) <> 0 Then
        XX_PP = (DSum(1) - E(1)) ^ 2 * ((1 / E(1)) + (1 / E(2)))
        P_PP = 1 - WorksheetFunction.ChiSq_Dist(XX_PP, 1, True)
        five_PP = IIf(P_PP <= 0.05, "有意性あり", "有意性なし")
        one_PP = IIf(P_PP <= 0.01, "有意性あり", "有意性なし")
    End If
    
    If V <> 0 Then
        XX_CMH = (DSum(1) - E(1)) ^ 2 / V
        P_CMH = 1 - WorksheetFunction.ChiSq_Dist(XX_CMH, 1, True)
        five_CMH = IIf(P_CMH <= 0.05, "有意性あり", "有意性なし")
        one_CMH = IIf(P_CMH <= 0.01, "有意性あり", "有意性なし")
    End If
    
'Add
    rcData = Array(aName, DSum(1) + WSum(1), DSum(1), WSum(1), ave(1), med(1))
    Call SAV.AddRow(res1, rcData, row1, 1)
    rcData = Array(bName, DSum(2) + WSum(2), DSum(2), WSum(2), ave(2), med(2))
    Call SAV.AddRow(res1, rcData, row1, 1)
    rcData = Array("2群合計", data(2, 2) + data(2, 5), DSum(1) + DSum(2), _
                    WSum(1) + WSum(2), "", "")
    Call SAV.AddRow(res1, rcData, row1, 1)
    rcData = Array("log-rank", "Peto-Peto", XX_PP, P_PP, five_PP, one_PP)
    Call SAV.AddRow(res2, rcData, row2, 1)
    rcData = Array("", "Cochran-Mantel-Haenszel", XX_CMH, P_CMH, five_CMH, one_CMH)
    Call SAV.AddRow(res2, rcData, row2, 1)
    rcData = Array(fin, "", "", _
                    0, "", "", "", S(1), SE(1), H(1), TR(1, 1), TR(1, 2), _
                    0, "", "", "", S(2), SE(2), H(2), TR(2, 1), TR(2, 2), _
                    0, 0, 0)
    Call SAV.AddRow(res3, rcData, row3, 1)
    rcData = Array(fin, S(1), S(2), H(1), H(2))
    Call SAV.AddRow(res4, rcData, row4, 1)
    
'刻み幅後処理（開始時点，終了時点，表時間）
    If step <> 0 Then
        For i = 3 To row3 - 1
            res3(i, 1) = res3(i, 1) * step
        Next
        For i = 3 To row3 - 2
            res3(i, 2) = res3(i, 2) * step
        Next
        For i = 3 To row4 - 1
            res4(i, 1) = res4(i, 1) * step
        Next
    End If
    
    KMLR = 1
    Exit Function
ERR:
    MsgBox ("エラーL：入力シートを確認してください。")
    KMLR = -1
End Function

Sub Graph(BookName, SheetName, serNum, ser(), xName, yName, yMax, xPos, yPos)
    Dim xyRange(1 To 5, 1 To 2) As Range    '1:SeriesNumber 2:x/y
    Dim rMax As Long
    Dim serName(1 To 5) As String
    Dim chartObj As ChartObject
    Dim i As Long
    
    With Workbooks(BookName).Sheets(SheetName)
        For i = 1 To serNum
            rMax = .Cells(Rows.Count, ser(i, 1, 2)).End(xlUp).row
            Set xyRange(i, 1) = Range(.Cells(ser(i, 1, 1) + 1, ser(i, 1, 2)), .Cells(rMax, ser(i, 1, 2)))
            Set xyRange(i, 2) = Range(.Cells(ser(i, 2, 1) + 1, ser(i, 2, 2)), .Cells(rMax, ser(i, 2, 2)))
            serName(i) = .Cells(ser(i, 2, 1), ser(i, 2, 2))
        Next
        Set chartObj = .ChartObjects.Add(xPos, yPos, 425, 275)
    End With
    
    With chartObj.Chart
        'Graph Options
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = False 'グラフタイトル
        .HasLegend = True '凡例
        .Legend.Position = xlLegendPositionBottom '凡例位置（下）
        With .Axes(xlCategory, xlPrimary) 'X軸
            .HasTitle = True '軸名
            .AxisTitle.Characters.Text = xName
            .TickLabels.NumberFormatLocal = "G/標準" '表示形式
        End With
        With .Axes(xlValue, xlPrimary) 'Y軸
            .HasTitle = True '軸名
            .AxisTitle.Characters.Text = yName
            .TickLabels.NumberFormatLocal = "G/標準" '表示形式
        End With
        
        If (yMax <> "") Then
            .Axes(xlValue).MaximumScale = yMax
        End If
        'Add Series
        For i = 1 To serNum
            .SeriesCollection.NewSeries
            With .SeriesCollection(i)
                .name = serName(i)
                .XValues = xyRange(i, 1)
                .Values = xyRange(i, 2)
            End With
        Next
    End With
End Sub

