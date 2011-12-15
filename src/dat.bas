Attribute VB_Name = "module_dat"
Option Explicit

Public Sub dat_process(path As String, handler_id As Long)
    On Error GoTo error:
    
    Dim content() As Byte
    Dim offset As Long
    Dim id As String
    Dim character As String

    content = byte_read(path)
    
    id = byte_string(content, &H2C0, &H15)
    character = byte_string(content, &H10, &H20)
    
    content = dat_deflate(content)
    
    If SafeArrayGetDim(content) > 0 Then
        offset = -1
        
        Do While offset < UBound(content)
            
            dat_store id, character, "Character", path, content, offset, handler_id
            
            offset = offset + &H224
        Loop
    End If
    
    Exit Sub
error:
    Log "[dat_process] Error while processing: " & path
End Sub

Public Sub dat_store(id As String, character As String, location As String, path As String, content() As Byte, offset As Long, handler_id As Long)
    On Error GoTo error:
    
        Dim query_string As String
        Dim error_string As String
        Dim return_code As Integer
        
        query_string = "INSERT INTO items VALUES("
        query_string = query_string & "'" & string_escape(id) & "',"
        query_string = query_string & "'" & string_escape(character) & "',"
        query_string = query_string & "'" & string_escape(location) & "',"
        query_string = query_string & "'" & string_escape(path) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H14) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H18) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H1C) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H20) & "',"
        query_string = query_string & "'" & string_escape(byte_string(content, offset + &H2C, &H20)) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H4C) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H50) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H24) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H26) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H5C) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H5E) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H60) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H62) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H64) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H66) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H68) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H6A) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H74) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H76) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H78) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H7C) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H80) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H84) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H88) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H8C) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H90) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H94) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H98) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H9C) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HA0) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HA4) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HA8) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HAC) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HB0) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &HB4) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HB8) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HBC) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HC0) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HC4) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HC8) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HCC) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &HF0) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H108) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H10C) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H110) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H114) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H118) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H11C) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H120) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H124) & "',"
        query_string = query_string & "'" & byte_dword(content, offset + &H150) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H156) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H158) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H15C) & "',"
        query_string = query_string & "'" & byte_float(content, offset + &H160) & "',"
        query_string = query_string & "'" & byte_word(content, offset + &H1EC) & "'"
        query_string = query_string & ")"
        
        sqlite_get_table handler_id, query_string, error_string, return_code
        
        If return_code <> SQLITE_OK Then
            Log "[dat_store] (" & return_code & ") " & error_string
        End If
        
        Main.items = Main.items + 1
error:
End Sub

Public Function dat_deflate(content() As Byte) As Variant
    On Error GoTo error:
    
    Dim deflate() As Byte
    Dim result() As Byte
    Dim offset As Long
    Dim total As Integer
    Dim i As Integer
    
    offset = &H6E4
    total = byte_dword(content, offset)
    
    offset = &H6F0
    
    For i = 1 To total
        result = dat_deflate_item(content, offset)
        array_merge_bytes deflate, result
        
    Next
    
    dat_deflate = deflate
    
    Exit Function
error:
    Log "[dat_deflate] Error while deflating"
    dat_deflate = StrConv("", vbFromUnicode)
End Function

Public Function dat_deflate_item(content() As Byte, ByRef offset As Long) As Variant
    On Error GoTo error:
    
    Dim deflate() As Byte
    Dim result() As Byte
    Dim size As Integer
    Dim start As Long
      
    size = byte_dword(content, offset)
    start = offset
    
    offset = offset + &H4
    
    Do While (offset - start) < size
    
        result = dat_deflate_element(content, offset)
        array_merge_bytes deflate, result
        
    Loop
    
    dat_deflate_item = deflate
    
    Exit Function
error:
    Log "[dat_deflate_item] Error while deflating"
    dat_deflate_item = StrConv("", vbFromUnicode)
End Function

Public Function dat_deflate_element(content() As Byte, ByRef offset As Long) As Variant
    On Error GoTo error:
    
    Dim deflate() As Byte
    Dim e As Integer
    Dim i As Integer
    Dim j As Integer
    
    e = byte_byte(content, offset)
    j = 0
    
    If (e And &H80) <> 0 Then
    
        offset = offset + 1
        
        For i = 1 To (e And &H7F)
        
            ReDim Preserve deflate(j) As Byte
            deflate(j) = 0
            j = j + 1
        Next
    Else
    
        offset = offset + 1
        
        For i = 1 To (e And &H7F)
        
            ReDim Preserve deflate(j) As Byte
            offset = offset + 1
            deflate(j) = content(offset)
            j = j + 1
        Next
    End If
    
    dat_deflate_element = deflate
    
    Exit Function
error:
    Log "[dat_deflate_element] Error while deflating"
    dat_deflate_element = StrConv("", vbFromUnicode)
End Function
