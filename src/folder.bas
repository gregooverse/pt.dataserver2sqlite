Attribute VB_Name = "module_folder"
Option Explicit
 
Public Type BrowseInfo
hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
 
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
 
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
 
Public Function folder_browser(hwndOwner As Long, sPrompt As String) As String
 
    'declare variables to be used
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
     
    'initialise variables
    With udtBI
    .hwndOwner = hwndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
    End With
     
    'Call the browse for folder API
    lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
    If lpIDList Then
    sPath = string$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then sPath = Left$(sPath, iNull - 1)
    End If
     
    'If cancel was pressed, sPath = ""
    folder_browser = sPath
 
End Function

Function folder_exists(path As String) As Boolean
    On Error GoTo error:
    
     folder_exists = GetAttr(path) And vbDirectory
     
error:
End Function
