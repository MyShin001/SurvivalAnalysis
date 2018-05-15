Attribute VB_Name = "Sampling"
Option Explicit

Function SmpMain(data, operator, terms())
On ERR GoTo ERR
    
    Dim dataInf() As Long, isIn() As Boolean
    Dim res(), UR As Long, row As Long
    Dim i As Long, j As Long
    
    Call SAV.GetArrayInfo(data, dataInf)
    ReDim isIn(1 To dataInf(1))
    isIn(1) = True
    
    Select Case operator
    Case 0, 2: UR = SmpAnd(data, dataInf, terms, isIn)
    Case 1, 3: UR = SmpOr(data, dataInf, terms, isIn)
    End Select
    
    Select Case operator
    Case 0, 1: ReDim res(1 To UR + 10, 1 To dataInf(2) + 4)
    Case 2, 3: ReDim res(1 To dataInf(1) - UR + 11, 1 To dataInf(2) + 4)
    End Select
    
    row = 1
    For i = 1 To UBound(isIn)
        If (i = 1) Or (operator <= 1 And isIn(i)) Or (operator > 1 And Not isIn(i)) Then
            For j = 1 To dataInf(2)
                res(row, j) = data(i, j)
            Next
            row = row + 1
        End If
    Next
    
    Call SAV.AddClm(res, Array("抽出数", row - 2), 2, UBound(res, 2) - 2)
    
    SmpMain = res
    Exit Function
ERR:
    MsgBox ("エラーS：抽出条件を確認してください。")
End Function

Function SmpAnd(data, dataInf, terms(), isIn() As Boolean)
    Dim condition As Boolean
    Dim resNum As Long
    Dim i As Long, j As Long
    
    resNum = dataInf(1)
    For i = 2 To UBound(isIn)
        isIn(i) = True
    Next
    
    For i = 1 To UBound(terms, 1)
        For j = 2 To dataInf(1)
            If isIn(j) And terms(i, 1) > 0 Then
                Select Case terms(i, 2)
                Case 0: condition = (CStr(data(j, terms(i, 1))) = terms(i, 3))
                Case 1: condition = (CStr(data(j, terms(i, 1))) <> terms(i, 3))
                Case 2: condition = (data(j, terms(i, 1)) < val(terms(i, 3)))
                Case 3: condition = (data(j, terms(i, 1)) > val(terms(i, 3)))
                Case 4: condition = (data(j, terms(i, 1)) <= val(terms(i, 3)))
                Case 5: condition = (data(j, terms(i, 1)) >= val(terms(i, 3)))
                End Select
                If Not condition Then
                    isIn(j) = False
                    resNum = resNum - 1
                End If
            End If
        Next
    Next
    
    SmpAnd = resNum
End Function

Function SmpOr(data, dataInf, terms(), isIn() As Boolean)
    Dim condition As Boolean
    Dim resNum As Long
    Dim i As Long, j As Long
    
    resNum = 1
    For i = 2 To UBound(isIn)
        isIn(i) = False
    Next
    
    For i = 1 To UBound(terms, 1)
        For j = 2 To dataInf(1)
            If Not isIn(j) And terms(i, 1) > 0 Then
                Select Case terms(i, 2)
                Case 0: condition = (CStr(data(j, terms(i, 1))) = terms(i, 3))
                Case 1: condition = (CStr(data(j, terms(i, 1))) <> terms(i, 3))
                Case 2: condition = (data(j, terms(i, 1)) < val(terms(i, 3)))
                Case 3: condition = (data(j, terms(i, 1)) > val(terms(i, 3)))
                Case 4: condition = (data(j, terms(i, 1)) <= val(terms(i, 3)))
                Case 5: condition = (data(j, terms(i, 1)) >= val(terms(i, 3)))
                End Select
                If condition Then
                    isIn(j) = True
                    resNum = resNum + 1
                End If
            End If
        Next
    Next
    
    SmpOr = resNum
End Function
