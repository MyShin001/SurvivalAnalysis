Attribute VB_Name = "Others"
Option Explicit

'// Win32API用定数
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
'// Win32API参照宣言
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Sub SAフォーム()
    UserForm.Show
End Sub

Public Sub FormResize()
    Dim hwnd As Long
    Dim style As Long
    
    hwnd = GetActiveWindow()
    style = GetWindowLong(hwnd, GWL_STYLE)
    style = style Or WS_MINIMIZEBOX
    Call SetWindowLong(hwnd, GWL_STYLE, style)
End Sub

Sub Focus(flag As Boolean)
    With Application
        .EnableEvents = flag
        .ScreenUpdating = flag
        .Calculation = IIf(flag, xlCalculationAutomatic, xlCalculationManual)
    End With
End Sub

Sub QuickSort(data(), min, max, key)
    'Declaration
    Dim r As Variant 'Reference
    Dim buf As Variant
    Dim i As Long, j As Long, k As Long
    
    'Initialize
    r = data(Int((min + max) / 2), key)
    i = min: j = max
    
    Do
        Do While data(i, key) < r
            i = i + 1
        Loop
        Do While data(j, key) > r
            j = j - 1
        Loop
        
        If i >= j Then Exit Do
        
        'Exchange Rows
        For k = LBound(data, 2) To UBound(data, 2)
            buf = data(i, k)
            data(i, k) = data(j, k)
            data(j, k) = buf
        Next
        
        i = i + 1
        j = j - 1
    Loop
    
    'Recall
    If (min < i - 1) Then
        Call QuickSort(data, min, i - 1, key)
    End If
    If (max > j + 1) Then
        Call QuickSort(data, j + 1, max, key)
    End If
End Sub

Function CheckName(BookName, SheetName) As String
'    On Error GoTo Err1
    
    'Declaration
    Dim ws As Worksheet
    Dim flag As Boolean: flag = False
    Dim flag2 As Boolean: flag2 = True
    Dim btn As Long
    Dim tmpName As String
    
    '同じ名前のシートがないかチェック
    For Each ws In Workbooks(BookName).Worksheets
        If ws.name = SheetName Then flag = True
    Next ws
    
    If flag = True Then
        btn = MsgBox("既に出力シートが存在します。上書きしますか？" & vbCr & _
                        "(「いいえ」を選択した場合は別名のシートに出力します)", vbYesNo + vbQuestion)
        If btn = vbYes Then
            Application.DisplayAlerts = False
            Workbooks(BookName).Worksheets(SheetName).Delete
            Application.DisplayAlerts = True
        Else
            Dim i As Long: i = 1
            Do While flag2 = True
                flag2 = False
                tmpName = SheetName & "(" & i & ")"
                For Each ws In Worksheets
                    If ws.name = tmpName Then flag2 = True
                Next ws
                i = i + 1
            Loop
            SheetName = tmpName
        End If
    End If
    CheckName = SheetName
End Function

Function IsArrayEx(varArray)
    On Error GoTo ERR

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function
ERR:
    If ERR.Number = 9 Then IsArrayEx = 0
End Function
