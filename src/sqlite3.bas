Attribute VB_Name = "module_sqlite3"
Option Explicit

' Handle a small selection of useful and relevant errors for us to handle
Public Const SQLITE_OK = 0
Public Const SQLITE_ERROR = 1
Public Const SQLITE_ABORT = 4
Public Const SQLITE_BUSY = 5
Public Const SQLITE_LOCKED = 6
Public Const SQLITE_INTERRUPT = 9
Public Const SQLITE_NOTFOUND = 12
Public Const SQLITE_AUTH = 23
Public Const SQLITE_NOTADB = 26
Public Const SQLITE_DONE = 101

' SQLite functions
Public Declare Function sqlite3_open Lib "sqlite3.dll" (ByVal FileName As String, _
    ByRef DB_Handle As Long) As Integer
Public Declare Function sqlite3_close Lib "sqlite3.dll" (ByRef DB_Handle As Long) As Integer

Public Declare Function sqlite_get_table Lib "sqlite3.dll" ( _
    ByVal DB_Handle As Long, _
    ByVal SQLString As String, _
    ByRef SQLErrorString As String, _
    ByRef SQLReturnCode As Integer) As Variant
    

'Row counter functs - may be changed later or merged as one
Public Declare Function sqlite_query_rowcount Lib "sqlite3.dll" () As Long

' Toolbox VB5+ functions
Public Declare Function GetArrayDimensions Lib "sqlite3.dll" (ByRef V As Variant) As Long
Public Declare Function GetArrayRows Lib "sqlite3.dll" (ByRef V As Variant) As Long
