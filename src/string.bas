Attribute VB_Name = "module_string"
Option Explicit

' Uber trim
Public Function string_trim(ByVal text As String) As String
    On Error Resume Next
    
    Dim text_length As Long
    Dim new_length As Long

    text = Replace(text, vbTab, " ")
    text = Trim$(text)
    
    text = string_undouble(text, " ")
    
    string_trim = text
End Function

' Remove doubled character
Public Function string_undouble(ByVal text As String, ByVal undouble As String) As String
    On Error Resume Next
    
    Dim text_length As Long
    Dim new_length As Long

    text_length = Len(text)
    
    Do While text_length <> new_length
    
        text_length = Len(text)
    
        text = Replace(text, undouble & undouble, undouble)
        
        new_length = Len(text)
    
    Loop
    
    string_undouble = text
End Function

' String contains
Public Function string_contains(ByVal haystack As String, needle As String) As Boolean
    On Error GoTo error:
    
    If (Len(needle) > Len(haystack)) Then
        string_contains = False
    Else
    
        string_contains = (InStr(haystack, needle) > 0)
        
    End If
    
    Exit Function
error:
    string_contains = False
End Function

' String count
Public Function string_count(ByVal haystack As String, ByVal needle As String, Optional ByVal method As VbCompareMethod = vbBinaryCompare) As Long
    Dim position As Long
    Dim count As Long
    
    If Len(haystack) = 0 Then
        Exit Function
    End If
    
    If Len(needle) = 0 Then
        Exit Function
    End If
    
    On Error GoTo error:
    
    position = 1
    
    Do
        position = InStr(position, haystack, needle, method)
        
        If position > 0 Then
        
            count = count + 1
            position = position + Len(needle)
            
        End If
        
    Loop Until position = 0
    
    string_count = count
    
    Exit Function
error:
    string_count = 0
End Function

Public Function string_escape(text As String) As String
    string_escape = Replace(text, "'", "''")
End Function


