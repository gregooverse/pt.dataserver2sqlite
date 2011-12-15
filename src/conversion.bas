Attribute VB_Name = "module_conversion"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function ByteArrayToLong(byte_array() As Byte) As Double
    On Error GoTo error:
        Dim i As Integer
        
        ByteArrayToLong = 0
        i = LBound(byte_array)
        
        Do While i <= UBound(byte_array)
        
            ByteArrayToLong = ByteArrayToLong + (byte_array(i) * (256 ^ i))
            
            i = i + 1
            
        Loop
            
    Exit Function
error:
    Log "[ByteArrayToLong] Error while converting"
    ByteArrayToLong = 0
End Function

Public Function ByteArrayToFloat(byte_array() As Byte) As Single
    On Error GoTo error:

    CopyMemory ByteArrayToFloat, byte_array(0), 4
    
    ByteArrayToFloat = Round(ByteArrayToFloat, 1)
    
    Exit Function
error:
    Log "[ByteArrayToFloat] Error while converting"
    ByteArrayToFloat = 0
End Function
