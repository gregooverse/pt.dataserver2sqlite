Attribute VB_Name = "module_path"
Option Explicit

' Path to processed log file
Function path_now() As String
    On Error GoTo error:
    
    path_now = datetime_formating("yyyy-mm-dd-hh-nn-ss") & ".ptsqlite"

    Exit Function
error:
    Log "[path_now] Error while creating 'now' path"
    path_now = "database_now_error.ptsqlite"
End Function

' Adding a antislash at the end of the path
Public Function path_slash(ByVal path As String, Optional ByVal trailing As Boolean = True) As String
    On Error GoTo error:
    
    path = Replace(path, "/", "\")
    path = string_undouble(path, "\")
    
    If trailing Then
        If Right$(path, 1) <> "\" Then
        
            path = path & "\"
            
        End If
    Else
        If Right$(path, 1) = "\" Then
        
            path = Left$(path, Len(path) - 1)
            
        End If
    End If
    
    path_slash = path
    
    Exit Function
error:
    Log "[path_slash] Error while cleaning path"
    If trailing Then
        path_slash = "\"
    Else
        path_slash = ""
    End If
End Function

Public Function path_basename(path As String, Optional separator = "\") As String
    Dim index As Long
    
    index = InStrRev(path, separator, , vbTextCompare)
    
    If index <> 0 Then
        path_basename = Mid(path, index + 1)
    Else
        path_basename = path
    End If
End Function
