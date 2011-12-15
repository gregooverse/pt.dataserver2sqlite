Attribute VB_Name = "module_log"
Option Explicit

Private Declare Function send_message Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7

Public Sub Log(message As String)
    Main.LogBox.text = Main.LogBox.text & Date & " " & Time & " -- " & message & vbCrLf
    
    send_message Main.LogBox.hwnd, WM_VSCROLL, SB_BOTTOM, 0
    
    DoEvents
End Sub
