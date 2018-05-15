Attribute VB_Name = "SAV"
Option Explicit

Public Function Sheet2Array(BookName, SheetName, fRow, fClm, eRow, eClm) As Variant
    With Workbooks(BookName).Sheets(SheetName)
        Sheet2Array = .Range(.Cells(fRow, fClm), .Cells(eRow, eClm))
    End With
End Function

Public Sub Array2Sheet(data, BookName, SheetName, fRow, fClm)
    Dim dataInf() As Variant
    Call GetArrayInfo(data, dataInf)
    With Workbooks(BookName).Sheets(SheetName)
        .Range(.Cells(fRow, fClm), .Cells(fRow + dataInf(1) - 1, fClm + dataInf(2) - 1)) = data
    End With
    fRow = fRow + dataInf(1) + 1: fClm = fClm + dataInf(2) + 1
End Sub

Public Function GetArrayInfo(data, dataInf)
    Dim i As Long, j As Long
    Dim Temp As Variant, condition As Boolean
    
    'Dimention check
    On Error GoTo DtErr
    For i = 1 To 3
        Temp = UBound(data, i) '3次元までのUBoundをチェック
    Next
DtErr:
    GetArrayInfo = i - 1 '(エラーが発生したi)-1が次元数
    On Error GoTo 0 'Error check ON
    
    '1D Array
    If GetArrayInfo = 1 Then
        ReDim Items(1 To 1)
        For i = 1 To UBound(data, 1)
            If data(i) <> Empty Then dataInf(1) = i '空欄でなければ要素数を増やす
        Next
        If dataInf(1) = Empty Then GetArrayInfo = -1 'データがなければ-1を返す
    End If
    
    '2D Array
    If GetArrayInfo = 2 Then
        ReDim dataInf(1 To 2)
        'dataInf(1)
        For i = UBound(data, 1) To 1 Step -1
            Select Case UBound(data, 2)
            Case 1: condition = data(i, 1) <> Empty
            Case Else: condition = data(i, 1) <> Empty Or data(i, 2) <> Empty
            End Select
            If condition Then Exit For
        Next
        dataInf(1) = i
        If dataInf(1) = 0 Then GetArrayInfo = -1
        'dataInf(2)
        For i = UBound(data, 2) To 1 Step -1
            Select Case UBound(data, 1)
            Case 1: condition = data(1, i) <> Empty
            Case Else: condition = data(1, i) <> Empty Or data(2, i) <> Empty
            End Select
            If condition Then Exit For
        Next
        dataInf(2) = i
        If dataInf(2) = 0 Then GetArrayInfo = -1
    End If
End Function

Sub AddRow(data(), rowData, fRow, fClm)
    Dim i As Long
    For i = fClm To fClm + UBound(rowData)
        data(fRow, i) = rowData(i - fClm)
    Next
    fRow = fRow + 1
End Sub

Sub AddClm(data, clmData, fRow, fClm)
    Dim i As Long
    For i = fRow To fRow + UBound(clmData)
        data(i, fClm) = clmData(i - fRow)
    Next
    fClm = fClm + 1
End Sub

Function SetArray(data(), tkey, ekey, xkey)
    Dim result(), rInf()
    Dim i As Long, j As Long
    
    ReDim result(1 To UBound(data, 1), 1 To 2)
    For i = 1 To UBound(data, 1)
        result(i, 1) = data(i, tkey)
        result(i, 2) = data(i, ekey)
    Next
    
    If IsArray(xkey) Then
        For j = 1 To UBound(xkey)
            ReDim Preserve result(1 To UBound(result, 1), 1 To UBound(result, 2) + 1)
            For i = 1 To UBound(data, 1)
                result(i, j + 2) = data(i, xkey(j))
            Next
        Next
    End If
    
    Call SAV.GetArrayInfo(result, rInf)
    result = CutStr(result, rInf)
    Call SAV.GetArrayInfo(result, rInf)
    If UBound(result, 1) > 5 Then
        Call Others.QuickSort(result, 2, UBound(result, 1) - 3, 1)
    End If
    
    SetArray = result
End Function

Function CutStr(data, dataInf())
    Dim isOut() As Boolean
    Dim result(), rNum As Long, row As Long
    Dim i As Long, j As Long
    
    ReDim isOut(1 To dataInf(1))
    rNum = dataInf(1)
    
    For i = 2 To dataInf(1)
        For j = 1 To dataInf(2)
            If Not isOut(i) Then
                If Not IsNumeric(data(i, j)) Or IsEmpty(data(i, j)) Then
                    isOut(i) = True
                    rNum = rNum - 1
                End If
            End If
        Next
    Next
    
    ReDim result(1 To rNum + 3, 1 To dataInf(2) + 3)
    row = 1
    For i = 1 To UBound(isOut)
        If Not isOut(i) Then
            For j = 1 To dataInf(2)
                result(row, j) = data(i, j)
            Next
            row = row + 1
        End If
    Next
    CutStr = result
End Function

