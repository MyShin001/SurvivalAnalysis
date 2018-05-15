Attribute VB_Name = "UserForm"
Option Explicit

Private Sub CoxCmbFile_Change()
    Dim ws As Worksheet
    Dim BookName As String
    Dim maxSize As Long

    If CoxCmbFile.ListIndex <> -1 Then
        BookName = CoxCmbFile.Text
        CoxCmbSheet.Clear
        For Each ws In Workbooks(BookName).Sheets
            CoxCmbSheet.AddItem ws.name
            If ws.name = ActiveSheet.name Then
                CoxCmbSheet.ListIndex = CoxCmbSheet.ListCount - 1
            End If
            If maxSize < LenB(ws.name) / 2 Then maxSize = LenB(ws.name) / 2
        Next
        CoxCmbSheet.ColumnWidths = maxSize * CoxCmbSheet.Font.size
    End If
End Sub

Private Sub CoxCmbSheet_Change()
    Dim BookName As String
    Dim SheetName As String
    Dim obj(), str As String, maxSize As Long
    Dim i As Long

    If CoxCmbSheet.ListIndex <> -1 Then
        BookName = CoxCmbFile.Text
        SheetName = CoxCmbSheet.Text
        obj = Array(CoxCmbTime, CoxCmbEvent, CoxLstX1, CoxLstX2)
        Call ClearItems(obj)
        With Workbooks(BookName).Worksheets(SheetName)
            obj = Array(CoxCmbTime, CoxCmbEvent, CoxLstX1)
            For i = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
                str = CNumAlp(i) & ". " & .Cells(1, i)
                Call AddItems(obj, str)
                If maxSize < LenB(str) / 2 Then maxSize = LenB(str) / 2
            Next
            obj = Array(CoxCmbTime, CoxCmbEvent, CoxLstX1, CoxLstX2)
            Call SetWidth(obj, maxSize * CoxLstX2.Font.size)
        End With
    End If
End Sub

Private Sub CoxCmdStart_Click()
    Select Case CoxCmbFile.ListIndex
        Case -1: MsgBox ("ファイル1を選択してください")
    Case Else
    Select Case CoxCmbSheet.ListIndex
        Case -1: MsgBox ("シート1を選択してください")
    Case Else
    Select Case CoxCmbTime.ListIndex
         Case -1: MsgBox ("時間列1を指定してください")
    Case Else
    Select Case CoxCmbEvent.ListIndex
         Case -1: MsgBox ("イベント判定列1を指定してください")
    Case Else
    Select Case CoxLstX2.ListCount
        Case 0: MsgBox ("説明変数を１つ以上選択してください")
    Case Else
    Select Case (CoxChSlct.value And (Not IsNumeric(CoxTxtPin.Text) Or Not IsNumeric(CoxTxtPout.Text)))
        Case True: MsgBox ("P値には数値のみを入力してください")
    Case Else
    Select Case (CoxChSlct.value And (val(CoxTxtPin.Text) < 0 Or val(CoxTxtPin.Text) >= 1))
        Case True: MsgBox ("P値は0以上1未満の数値で入力してください")
    Case Else
    Select Case (CoxChSlct.value And (val(CoxTxtPout.Text) < 0 Or val(CoxTxtPout.Text) >= 1))
        Case True: MsgBox ("P値は0以上1未満の数値で入力してください")
    Case Else
    Select Case val(CoxTxtPin.Text) > val(CoxTxtPout.Text)
        Case True: MsgBox ("追加P値には除去P値以下の数値を設定してください")
    Case Else
    Call Others.Focus(False)

'宣言
    Dim BookName As String, SheetName As String
    Dim data(), dataInf(), beta()
    Dim maxRow As Long, maxClm As Long
    Dim tkey As Long, ekey As Long, xkey()
    Dim Pin As Double, Pout As Double
    Dim TotalTable(), BasicTable(), StandTable(), SlctTable(), BetaTable()
    Dim rs As Worksheet, rRow As Long, tmpRow As Long
    Dim i As Long

'初期化
    BookName = CoxCmbFile.value
    SheetName = CoxCmbSheet.value
    With Workbooks(BookName).Sheets(SheetName)
        maxRow = .Cells(Rows.Count, 1).End(xlUp).row
        maxClm = .Cells(1, Columns.Count).End(xlToLeft).Column
        data = SAV.Sheet2Array(BookName, SheetName, 1, 1, maxRow + 3, maxClm + 3)
    End With
    tkey = CoxCmbTime.ListIndex + 1
    ekey = CoxCmbEvent.ListIndex + 1
    For i = 1 To CoxLstX2.ListCount
        ReDim Preserve xkey(1 To i)
        xkey(UBound(xkey)) = CNumAlp(Mid(CoxLstX2.List(i - 1), 1, 1))
    Next
    rRow = 2

'処理
    data = SAV.SetArray(data, tkey, ekey, xkey)
    Call SAV.GetArrayInfo(data, dataInf)

    If dataInf(1) < 3 Then
        MsgBox ("有効なサンプル数が1以下のため終了します。")
        Call Others.Focus(True): Exit Sub
    End If

    Call Cox.BasicData(data, dataInf, TotalTable, BasicTable)

    If CoxChSlct.value Then '変数選択あり
        Pin = val(CoxTxtPin.Text)
        Pout = val(CoxTxtPout.Text)
        ReDim StandTable(1 To 2, 1 To 2)
        StandTable(1, 1) = "投入の基準P値": StandTable(1, 2) = "除去の基準P値"
        StandTable(2, 1) = Pin: StandTable(2, 2) = Pout
        i = Cox.selectX(data, dataInf, xkey, Pin, Pout, SlctTable, BetaTable)
    Else '変数選択なし
        Call Cox.simple(data, dataInf, BetaTable)
    End If

'出力
    Set rs = Workbooks(BookName).Sheets.Add(After:=Sheets(Sheets.Count))
    rs.name = Others.CheckName(BookName, "Cox")

    Call SAV.Array2Sheet(TotalTable, BookName, rs.name, rRow, 1)
    Call SAV.Array2Sheet(BasicTable, BookName, rs.name, rRow, 1)
    tmpRow = rRow
    If CoxChSlct.value Then
        Call SAV.Array2Sheet(StandTable, BookName, rs.name, rRow, 1)
        Call SAV.Array2Sheet(SlctTable, BookName, rs.name, rRow, 1)
    End If
    Call SAV.Array2Sheet(BetaTable, BookName, rs.name, rRow, 1)

'表示設定（小数点以下4桁まで，整数は小数点なし，絶対値0,000005未満は0）
    With Range(Cells(1, 1), Cells(rRow, 6))
        .NumberFormatLocal = "[>=0.00005]#,##0.0###;[<=-0.00005]-#,##0.0###;0"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=INDIRECT(ADDRESS(ROW(),COLUMN()))=TRUNC(INDIRECT(ADDRESS(ROW(),COLUMN())))"
        .FormatConditions(1).NumberFormat = "#,##0"
'        .EntireColumn.ColumnWidth = 10
'        .EntireColumn.AutoFit
    End With

    Call Others.Focus(True)
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
End Sub

Private Sub CoxLstX1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long, flag As Boolean
    flag = True
    For i = 0 To CoxLstX2.ListCount - 1
        If CoxLstX1.value = CoxLstX2.List(i) Then flag = False
    Next
    If flag Then CoxLstX2.AddItem CoxLstX1.value

End Sub

Private Sub CoxLstX2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CoxLstX2.RemoveItem CoxLstX2.ListIndex
End Sub

Private Sub KMLRCmbEvent1_Change()
    Dim index As Long
    index = KMLRCmbEvent1.ListIndex
    If KMLRCmbEvent2.ListCount >= index + 1 Then
        KMLRCmbEvent2.ListIndex = index
    End If
End Sub

Private Sub KMLRCmbFile1_Change()
    Dim index As Long
    Dim ws As Worksheet
    Dim BookName As String
    Dim maxSize As Long
    index = KMLRCmbFile1.ListIndex
    If index <> -1 Then
        BookName = KMLRCmbFile1.Text
        KMLRCmbFile2.ListIndex = index
        KMLRCmbSheet1.Clear
        For Each ws In Workbooks(BookName).Sheets
            KMLRCmbSheet1.AddItem ws.name
            If ws.name = ActiveSheet.name Then
                KMLRCmbSheet1.ListIndex = KMLRCmbSheet1.ListCount - 1
            End If
            If maxSize < LenB(ws.name) / 2 Then maxSize = LenB(ws.name) / 2
        Next
        KMLRCmbSheet1.ColumnWidths = maxSize * KMLRCmbSheet1.Font.size
    End If
End Sub

Private Sub KMLRCmbFile2_Change()
    Dim ws As Worksheet
    Dim BookName As String
    Dim maxSize As Long

    If KMLRCmbFile2.ListIndex <> -1 Then
        BookName = KMLRCmbFile2.Text
        KMLRCmbSheet2.Clear
        For Each ws In Workbooks(BookName).Sheets
            KMLRCmbSheet2.AddItem ws.name
            If ws.name = ActiveSheet.name Then
                KMLRCmbSheet2.ListIndex = KMLRCmbSheet2.ListCount - 1
            End If
            If maxSize < LenB(ws.name) / 2 Then maxSize = LenB(ws.name) / 2
        Next
        KMLRCmbSheet2.ColumnWidths = maxSize * KMLRCmbSheet2.Font.size
    End If
End Sub

Private Sub KMLRCmbSheet1_Change()
    Dim BookName As String
    Dim SheetName As String
    Dim obj(), str As String, maxSize As Long

    If KMLRCmbSheet1.ListIndex <> -1 Then
        BookName = KMLRCmbFile1.Text
        SheetName = KMLRCmbSheet1.Text
        obj = Array(KMLRCmbTime1, KMLRCmbEvent1)
        Call ClearItems(obj)
        With Workbooks(BookName).Worksheets(SheetName)
            Dim i As Long
            For i = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
                str = CNumAlp(i) & ". " & .Cells(1, i)
                Call AddItems(obj, str)
                If maxSize < LenB(str) / 2 Then maxSize = LenB(str) / 2
            Next
            Call SetWidth(obj, maxSize * KMLRCmbEvent1.Font.size)
        End With
    End If
End Sub

Private Sub KMLRCmbSheet2_Change()
    Dim BookName As String
    Dim SheetName As String
    Dim obj(), str As String, maxSize As Long

    If KMLRCmbSheet2.ListIndex <> -1 Then
        BookName = KMLRCmbFile2.Text
        SheetName = KMLRCmbSheet2.Text
        obj = Array(KMLRCmbTime2, KMLRCmbEvent2)
        Call ClearItems(obj)
        With Workbooks(BookName).Worksheets(SheetName)
            Dim i As Long
            For i = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
                str = CNumAlp(i) & ". " & .Cells(1, i)
                Call AddItems(obj, str)
                If maxSize < LenB(str) / 2 Then maxSize = LenB(str) / 2
            Next
            Call SetWidth(obj, maxSize * KMLRCmbEvent2.Font.size)
        End With
    End If
End Sub

Private Sub KMLRCmbTime1_Change()
    Dim index As Long
    index = KMLRCmbTime1.ListIndex
    If KMLRCmbTime2.ListCount >= index + 1 Then
        KMLRCmbTime2.ListIndex = index
    End If
End Sub

Private Sub KMLRCmdStart_Click()
    Select Case KMLRCmbFile1.ListIndex
        Case -1: MsgBox ("ファイル1を選択してください")
    Case Else
    Select Case KMLRCmbSheet1.ListIndex
        Case -1: MsgBox ("シート1を選択してください")
    Case Else
    Select Case KMLRCmbTime1.ListIndex
         Case -1: MsgBox ("時間列1を指定してください")
    Case Else
    Select Case KMLRCmbEvent1.ListIndex
         Case -1: MsgBox ("イベント判定列1を指定してください")
    Case Else
    Select Case KMLRCmbFile2.ListIndex
        Case -1: MsgBox ("ファイル2を選択してください")
    Case Else
    Select Case KMLRCmbSheet2.ListIndex
        Case -1: MsgBox ("シート2を選択してください")
    Case Else
    Select Case KMLRCmbTime2.ListIndex
         Case -1: MsgBox ("時間列2を指定してください")
    Case Else
    Select Case KMLRCmbEvent2.ListIndex
         Case -1: MsgBox ("イベント判定2を指定してください")
    Case Else
    Select Case (KMLRChStep.value And KMLRTxtStep.value = "")
        Case True: MsgBox ("刻み幅が入力されていません")
    Case Else
    Call Others.Focus(False)

'宣言
    Dim BookName1 As String, BookName2 As String
    Dim SheetName1 As String, SheetName2 As String
    Dim maxRow As Long, maxClm As Long
    Dim data1(), data2(), dataInf1(), dataInf2()
    Dim res1(), res2(), res3(), res4()
    Dim tkey1 As Long, tkey2 As Long
    Dim ekey1 As Long, ekey2 As Long
    Dim rs As Worksheet, row As Long, lRow As Long, gRow As Long
    Dim step As Double
    Dim ser(2, 2, 2), hight As Double
    Dim i As Long

'初期化
    BookName1 = KMLRCmbFile1.Text
    SheetName1 = KMLRCmbSheet1.Text
    With Workbooks(BookName1).Sheets(SheetName1)
        maxRow = .Cells(Rows.Count, 1).End(xlUp).row
        maxClm = .Cells(1, Columns.Count).End(xlToLeft).Column
        data1 = SAV.Sheet2Array(BookName1, SheetName1, 1, 1, maxRow + 3, maxClm + 3)
    End With

    BookName2 = KMLRCmbFile2.Text
    SheetName2 = KMLRCmbSheet2.Text
    With Workbooks(BookName2).Sheets(SheetName2)
        maxRow = .Cells(Rows.Count, 1).End(xlUp).row
        maxClm = .Cells(1, Columns.Count).End(xlToLeft).Column
        data2 = SAV.Sheet2Array(BookName2, SheetName2, 1, 1, maxRow + 3, maxClm + 3)
    End With

    tkey1 = KMLRCmbTime1.ListIndex + 1: tkey2 = KMLRCmbTime2.ListIndex + 1
    ekey1 = KMLRCmbEvent1.ListIndex + 1: ekey2 = KMLRCmbEvent2.ListIndex + 1

    data1 = SAV.SetArray(data1, tkey1, ekey1, 0)
    data2 = SAV.SetArray(data2, tkey2, ekey2, 0)

    Call SAV.GetArrayInfo(data1, dataInf1)
    Call SAV.GetArrayInfo(data2, dataInf2)

    If (dataInf1(1) < 3 Or dataInf2(1) < 3) Then
        MsgBox ("有効なサンプル数が1以下です")
        Call Others.Focus(True): Exit Sub
    End If

'刻み幅前処理
    If KMLRChStep.value Then
        step = val(KMLRTxtStep.value)
        If step >= data1(dataInf1(1), 1) Or step >= data2(dataInf2(1), 1) Then
            MsgBox ("刻み幅が大きすぎます")
            Call Others.Focus(True): Exit Sub
        End If
    Else
        step = 0
    End If

    Call KMLR.KMLR(KMLR.BasicArray(data1, data2, step), SheetName1, SheetName2, _
                    step, res1, res2, res3, res4)

'出力
    Set rs = Workbooks(BookName1).Sheets.Add(After:=Sheets(Sheets.Count))
    With rs

    hight = .Rows(1).RowHeight
    .name = Others.CheckName(BookName1, "KMLR")
    row = 2
    Call SAV.Array2Sheet(res1, BookName1, .name, row, 1)
    Call SAV.Array2Sheet(res2, BookName1, .name, row, 1)
    row = row + 275 / hight + 2: lRow = row
    Call SAV.Array2Sheet(res3, BookName1, .name, row, 1)
    gRow = row
    Call SAV.Array2Sheet(res4, BookName1, .name, row, 1)
'表示設定（小数点以下4桁まで，整数は小数点なし，絶対値0,000005未満は0）
    With Range(Cells(1, 1), Cells(row, 24))
        .NumberFormatLocal = "[>=0.00005]#,##0.0###;[<=-0.00005]-#,##0.0###;0"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=INDIRECT(ADDRESS(ROW(),COLUMN()))=TRUNC(INDIRECT(ADDRESS(ROW(),COLUMN())))"
        .FormatConditions(1).NumberFormat = "#,##0"
'        .EntireColumn.ColumnWidth = 10
'        .EntireColumn.AutoFit
    End With

'枠線追加
    .Range(.Cells(2, 1), .Cells(5, 1)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(2, 6), .Cells(5, 6)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(1, 1), .Cells(1, 6)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(2, 1), .Cells(2, 6)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(4, 1), .Cells(4, 6)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(5, 1), .Cells(5, 6)).Borders(xlEdgeBottom).LineStyle = True

    .Range(.Cells(7, 1), .Cells(9, 1)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(7, 6), .Cells(9, 6)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(6, 1), .Cells(6, 6)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(7, 1), .Cells(7, 6)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(9, 1), .Cells(9, 6)).Borders(xlEdgeBottom).LineStyle = True

    .Range(.Cells(lRow, 3), .Cells(gRow - 2, 3)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(lRow, 12), .Cells(gRow - 2, 12)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(lRow, 21), .Cells(gRow - 2, 21)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(lRow, 24), .Cells(gRow - 2, 24)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(lRow - 1, 1), .Cells(lRow - 1, 24)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(lRow, 1), .Cells(lRow, 24)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(lRow + 1, 1), .Cells(lRow + 1, 24)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(gRow - 2, 1), .Cells(gRow - 2, 24)).Borders(xlEdgeBottom).LineStyle = True

    .Range(.Cells(gRow, 1), .Cells(row - 2, 1)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(gRow, 3), .Cells(row - 2, 3)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(gRow, 5), .Cells(row - 2, 5)).Borders(xlEdgeRight).LineStyle = True
    .Range(.Cells(gRow - 1, 1), .Cells(gRow - 1, 5)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(gRow, 1), .Cells(gRow, 5)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(gRow + 1, 1), .Cells(gRow + 1, 5)).Borders(xlEdgeBottom).LineStyle = True
    .Range(.Cells(row - 2, 1), .Cells(row - 2, 5)).Borders(xlEdgeBottom).LineStyle = True

'グラフ
    ser(1, 1, 1) = gRow + 1: ser(1, 2, 1) = gRow + 1
    ser(1, 1, 2) = 1: ser(1, 2, 2) = 2
    ser(2, 1, 1) = gRow + 1: ser(2, 2, 1) = gRow + 1
    ser(2, 1, 2) = 1: ser(2, 2, 2) = 3
    Call Graph(BookName1, rs.name, 2, ser, _
                    rs.Cells(ser(1, 1, 1), 1), "累計生存率", 1, 25, hight * 10.5)
    ser(1, 1, 2) = 1: ser(1, 2, 2) = 4
    ser(2, 1, 2) = 1: ser(2, 2, 2) = 5
    Call Graph(BookName1, rs.name, 2, ser, _
                    rs.Cells(ser(1, 1, 1), 1), "累計ハザード関数", "", 500, hight * 10.5)

    End With
    Call Others.Focus(True)
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
    End Select
End Sub

Private Sub MultiPage_Change()
    Dim obj(), item()
    Dim wb As Workbook
    Dim maxSize As Long

    obj = Array(SmpCmbFile, KMLRCmbFile1, KMLRCmbFile2, CoxCmbFile)
    Call ClearItems(obj)
    For Each wb In Workbooks
        Call AddItems(obj, wb.name)
        If wb.name = ActiveWorkbook.name Then
            Call SetIndex(obj, SmpCmbFile.ListCount - 1)
        End If
        If maxSize < LenB(wb.name) / 2 Then maxSize = LenB(wb.name) / 2
    Next
    Call SetWidth(obj, maxSize * CoxCmbFile.Font.size)
End Sub

Private Sub SmpCmbFile_Change()
    Dim BookName As String
    Dim ws As Worksheet
    Dim maxSize As Long

    If SmpCmbFile.ListIndex <> -1 Then
        BookName = SmpCmbFile.Text
        SmpCmbSheet.Clear
        For Each ws In Workbooks(BookName).Sheets
            SmpCmbSheet.AddItem ws.name
            If ws.name = ActiveSheet.name Then
                SmpCmbSheet.ListIndex = SmpCmbSheet.ListCount - 1
            End If
            If maxSize < LenB(ws.name) / 2 Then maxSize = LenB(ws.name) / 2
        Next
        SmpCmbSheet.ColumnWidths = maxSize * SmpCmbSheet.Font.size
    End If
End Sub

Private Sub SmpCmbSheet_Change()
    Dim BookName As String
    Dim SheetName As String
    Dim obj(), str As String
    Dim maxSize As Long
    Dim i As Long

    If SmpCmbSheet.ListIndex <> -1 Then
        BookName = SmpCmbFile.Text
        SheetName = SmpCmbSheet.Text
        obj = Array(SmpCmbIs1, SmpCmbIs2, SmpCmbIs3, SmpCmbIs4, SmpCmbIs5)
        Call SetIndex(obj, 0)
        obj = Array(SmpTxtVal1, SmpTxtVal2, SmpTxtVal3, SmpTxtVal4, SmpTxtVal5)
        Call SetValue(obj, "")
        obj = Array(SmpCmbClm1, SmpCmbClm2, SmpCmbClm3, SmpCmbClm4, SmpCmbClm5)
        Call ClearItems(obj)
        Call AddItems(obj, "")
        With Workbooks(BookName).Worksheets(SheetName)
            For i = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
                str = CNumAlp(i) & ". " & .Cells(1, i)
                Call AddItems(obj, str)
                If maxSize < LenB(str) / 2 Then maxSize = LenB(str) / 2
            Next
            Call SetWidth(obj, maxSize * SmpCmbClm5.Font.size)
        End With
    End If
End Sub

Private Sub SmpCmdReset_Click()
    Dim obj()
    obj = Array(SmpCmbClm1, SmpCmbClm2, SmpCmbClm3, SmpCmbClm4, SmpCmbClm5)
    Call SetIndex(obj, -1)
    obj = Array(SmpCmbIs1, SmpCmbIs2, SmpCmbIs3, SmpCmbIs4, SmpCmbIs5)
    Call SetIndex(obj, 0)
    obj = Array(SmpTxtVal1, SmpTxtVal2, SmpTxtVal3, SmpTxtVal4, SmpTxtVal5)
    Call SetValue(obj, "")
End Sub

Private Sub SmpCmdStart_Click()
    Select Case SmpCmbFile.ListIndex
        Case -1: MsgBox ("ファイルを選択してください")
    Case Else
    Select Case SmpCmbSheet.ListIndex
        Case -1: MsgBox ("シートを選択してください")
    Case Else
    Select Case SmpCmbAO.ListIndex
        Case -1: MsgBox ("抽出方法を選択してください")
    Case Else
    Call Others.Focus(False)

'宣言
    Dim BookName As String
    Dim SheetName As String
    Dim operator As Variant
    Dim terms(1 To 5, 1 To 3) As Variant
    Dim data, dataInf() As Long
    Dim maxRow As Long, maxClm As Long
    Dim clmData As Variant, row As Long
    Dim i As Long

'初期化
    BookName = SmpCmbFile.Text
    SheetName = SmpCmbSheet.Text
    operator = SmpCmbAO.ListIndex

    For i = 1 To 5
        terms(i, 1) = Me.Controls("SmpCmbClm" & i).ListIndex
        terms(i, 2) = Me.Controls("SmpCmbIs" & i).ListIndex
        terms(i, 3) = Me.Controls("SmpTxtVal" & i).Text
    Next

    With Workbooks(BookName).Sheets(SheetName)
        maxRow = .Cells(Rows.Count, 1).End(xlUp).row
        maxClm = .Cells(1, Columns.Count).End(xlToLeft).Column
        data = .Range(.Cells(1, 1), .Cells(maxRow, maxClm))
    End With

    ReDim clmData(5)
    clmData(0) = "抽出条件（" & SmpCmbAO.Text & "）"
    row = 1
    For i = 1 To 5
        If terms(i, 2) = -1 Then
            MsgBox ("等号・不等号を指定してください。")
            Call Others.Focus(True): Exit Sub
        End If

        If terms(i, 1) > 0 Then
            Select Case terms(i, 2)
            Case 0: clmData(row) = data(1, terms(i, 1)) & " ＝ " & terms(i, 3)
            Case 1: clmData(row) = data(1, terms(i, 1)) & " ≠ " & terms(i, 3)
            Case 2: clmData(row) = data(1, terms(i, 1)) & " < " & terms(i, 3)
            Case 3: clmData(row) = data(1, terms(i, 1)) & " > " & terms(i, 3)
            Case 4: clmData(row) = data(1, terms(i, 1)) & " ≦ " & terms(i, 3)
            Case 5: clmData(row) = data(1, terms(i, 1)) & " ≧ " & terms(i, 3)
            End Select
            row = row + 1
        End If
Next

'処理
    data = SmpMain(data, operator, terms)
    If Not IsArray(data) Then Call Others.Focus(True): Exit Sub
    Call SAV.AddClm(data, clmData, 2, UBound(data, 2) - 1)

'出力
    Dim rs As Worksheet
    With Workbooks(BookName)
        Set rs = .Sheets.Add(After:=.Worksheets(.Worksheets.Count))
    End With
    rs.name = Others.CheckName(BookName, "Sampling")
    rs.Range(rs.Cells(1, 1), rs.Cells(UBound(data, 1), UBound(data, 2))) = data

    Call Others.Focus(True)
    End Select
    End Select
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim obj(), item()
    Dim wb As Workbook
    Dim maxSize As Long

    Call AddItems(Array(SmpCmbAO), Array("AND", "OR", "NAND", "NOR"))
    SmpCmbAO.ListIndex = 0

    obj = Array(SmpCmbIs1, SmpCmbIs2, SmpCmbIs3, SmpCmbIs4, SmpCmbIs5)
    item = Array("＝", "≠", "<", ">", "≦", "≧")
    Call AddItems(obj, item)
    Call SetIndex(obj, 0)

    obj = Array(SmpCmbFile, KMLRCmbFile1, KMLRCmbFile2, CoxCmbFile)
    For Each wb In Workbooks
        Call AddItems(obj, wb.name)
        If wb.name = ActiveWorkbook.name Then
            Call SetIndex(obj, SmpCmbFile.ListCount - 1)
        End If
        If maxSize < LenB(wb.name) / 2 Then maxSize = LenB(wb.name) / 2
    Next
    Call SetWidth(obj, maxSize * CoxCmbFile.Font.size)
End Sub

Private Sub UserForm_Activate()
    Call FormResize
End Sub