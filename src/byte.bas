Attribute VB_Name = "module_byte"
Option Explicit

Public Function byte_read(path As String) As Variant
    On Error GoTo error:
    
    Dim descriptor As Long
    Dim bytes() As Byte
    
    descriptor = FreeFile
    
    Open path For Binary Access Read As descriptor
    
    ReDim bytes(1 To LOF(descriptor)) As Byte
    
    ' read blob
    Get descriptor, , bytes
    
    ' close file
    Close descriptor
    
    byte_read = bytes
    
    Exit Function
error:
    Log "[byte_read] Error while reading : " & path
    byte_read = StrConv("", vbFromUnicode)
End Function

Public Function byte_byte(content() As Byte, start As Long) As Integer
    On Error GoTo error:
    
    Dim chunck() As Byte
    
    chunck = byte_chunck(content, start, 1)
    
    byte_byte = ByteArrayToLong(chunck)
    
    Exit Function
error:
    Log "[byte_byte] Error while getting byte at : " & start
    byte_byte = 0
End Function

Public Function byte_word(content() As Byte, start As Long) As Double
    On Error GoTo error:
    
    Dim chunck() As Byte
    
    chunck = byte_chunck(content, start, 2)
    
    byte_word = ByteArrayToLong(chunck)
    
    Exit Function
error:
    Log "[byte_word] Error while getting word at : " & start
    byte_word = 0
End Function

Public Function byte_dword(content() As Byte, start As Long) As Double
    On Error GoTo error:
    
    Dim chunck() As Byte
    
    chunck = byte_chunck(content, start, 4)
    
    byte_dword = ByteArrayToLong(chunck)
    
    Exit Function
error:
    Log "[byte_dword] Error while getting dword at : " & start
    byte_dword = 0
End Function

Public Function byte_chunck(content() As Byte, start As Long, Length As Integer) As Variant
    On Error GoTo error:
    
    Dim chunck() As Byte
    Dim offset As Long
    Dim i As Integer
    
    ReDim chunck(Length - 1) As Byte
    
    offset = start + 1
    
    If UBound(content) >= offset + Length Then
    
        For i = 0 To Length - 1
        
            chunck(i) = content(offset)
                    
            offset = offset + 1
        Next
     End If
             
     byte_chunck = chunck
   
    Exit Function
error:
    Log "[byte_chunck] Error while getting chunck of size " & Length & " at : " & start
    byte_chunck = StrConv("", vbFromUnicode)
End Function

Public Function byte_string(content() As Byte, start As Long, Length As Integer) As String
    On Error GoTo error:
    
    Dim char As Byte
    Dim offset As Long
    Dim i As Integer
    
    offset = start + 1
    i = 0
    
    If UBound(content) >= offset + Length Then
        For i = 0 To Length - 1
            char = content(offset + i)
            
            If char = 0 Then
                Exit For
            End If
            
            byte_string = byte_string & Chr(char)
        
        Next
    End If
    
    Exit Function
error:
    Log "[byte_string] Error while getting string of size " & Length & " at : " & start
    byte_string = ""
End Function

Public Function byte_float(content() As Byte, start As Long) As Single
    On Error GoTo error:
    
    Dim chunck() As Byte
    
    chunck = byte_chunck(content, start, 4)
    
    byte_float = ByteArrayToFloat(chunck)
    
    Exit Function
error:
    Log "[byte_float] Error while getting float at : " & start
    byte_float = 0
End Function
