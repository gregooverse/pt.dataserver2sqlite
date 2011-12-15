Attribute VB_Name = "module_registry"
Option Explicit

Const REG_KEY = "PT2SQLite"

' Opens a key in the registry
Public Function registry_open(ByVal key As String, Optional ByVal default As String = "") As String
    On Error Resume Next
    
    registry_open = GetSetting(REG_KEY, "Configuration", key, default)
End Function

' Saves a key in the registry
Public Sub registry_save(ByVal key As String, ByVal value As String)
    On Error Resume Next
    
    SaveSetting REG_KEY, "Configuration", key, value
End Sub

' Delete a key in the registry
Public Sub registry_delete(ByVal key As String)
    On Error Resume Next
    
    If registry_open(key) <> "" Then
        DeleteSetting REG_KEY, "Configuration", key
    End If
End Sub
