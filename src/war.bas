Attribute VB_Name = "module_war"
Option Explicit

Public Sub war_process(path As String, handler_id As Long)
    On Error GoTo error:
    
    Dim content() As Byte
    Dim offset As Long
    Dim id As String

    id = path_basename(path)
    id = Left$(id, Len(id) - 4)

    content = byte_read(path)
    
    content = war_deflate(content)
    
    If SafeArrayGetDim(content) > 0 Then
        offset = -1
        
        Do While offset < UBound(content) And byte_dword(content, offset)
        
            offset = offset + &HF0
            
            dat_store id, "", "Warehouse", path, content, offset, handler_id
            
            offset = offset + &H224
        Loop
    End If
    
    Exit Sub
error:
    Log "[war_process] Error while processing: " & path
End Sub

Public Function war_deflate(content() As Byte) As Variant
    On Error GoTo error:
    
    Dim offset As Long

    offset = &H30

    war_deflate = dat_deflate_item(content, offset)
    
    Exit Function
error:
    Log "[war_deflate] Error while deflating"
    war_deflate = StrConv("", vbFromUnicode)
End Function

