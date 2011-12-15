Attribute VB_Name = "module_array"
Option Explicit

Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Sub array_merge_bytes(ByRef stack() As Byte, push() As Byte)
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim size As Long
    
    count = 0
    j = 0
    
    If SafeArrayGetDim(stack) > 0 Then
        count = count + UBound(stack) + 1
        j = UBound(stack) + 1
    End If
    
    If SafeArrayGetDim(push) > 0 Then
        count = count + UBound(push) + 1
    End If
    
    If count > 0 Then
    
        size = count - 1
    
        ReDim Preserve stack(size) As Byte
                
        i = 0
        Do While j <= size
            stack(j) = push(i)
            j = j + 1
            i = i + 1
        Loop
    End If
    
End Sub

