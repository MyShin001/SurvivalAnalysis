Attribute VB_Name = "Form"
Option Explicit

Function CNumAlp(val As Variant) As Variant
    Dim al As String
    If IsNumeric(val) = True Then
        al = Cells(1, val).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        CNumAlp = Left(al, Len(al) - 1)
    Else 'If Alfabet
        CNumAlp = Range(val & "1").Column
    End If
End Function

Sub AddItems(obj, item)
    Dim i As Long, j As Long
    For i = 0 To UBound(obj)
        If IsArray(item) Then
            For j = 0 To UBound(item)
                obj(i).AddItem item(j)
            Next
        Else
            obj(i).AddItem item
        End If
    Next
End Sub

Sub ClearItems(obj)
    Dim i As Long
    For i = 0 To UBound(obj)
        obj(i).Clear
    Next
End Sub

Sub SetIndex(obj, num)
    Dim i As Long
    For i = 0 To UBound(obj)
        obj(i).ListIndex = num
    Next
End Sub

Sub SetValue(obj, value)
    Dim i As Long
    For i = 0 To UBound(obj)
        obj(i).value = value
    Next
End Sub

Sub SetWidth(obj, size)
    Dim i As Long
    For i = 0 To UBound(obj)
        obj(i).ColumnWidths = size
    Next
End Sub
