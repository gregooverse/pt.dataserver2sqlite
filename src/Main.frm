VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PT2SQLite by Gregoo"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressWar 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   255
   End
   Begin ComctlLib.ProgressBar ProgressDat 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   255
   End
   Begin VB.TextBox LogBox 
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
   End
   Begin VB.CommandButton ScanStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton ScanStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame FolderBox 
      Caption         =   "Dataserver path (contains userdata && warehouse folders)"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton FolderSelect 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox FolderPath 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VERSION = "0.1"

Public dats As Long
Public wars As Long
Public items As Double
Public run As Byte

Private Sub FolderPath_LostFocus()
    FolderPath.text = path_slash(FolderPath.text)
    registry_save "FolderPath", FolderPath.text
End Sub

Private Sub FolderSelect_Click()
    Dim path As String
    
    path = folder_browser(Me.hwnd, "Select the dataserver folder")
    
    If path <> "" Then
        FolderPath.text = path_slash(path)
        registry_save "FolderPath", FolderPath.text
        
        If Not folder_exists(FolderPath.text & "userdata") Then
            Log "userdata folder not found"
        End If
        
        If Not folder_exists(FolderPath.text & "warehouse") Then
            Log "warehouse folder not found"
        End If
    End If
End Sub

Private Sub Form_Load()
    Main.Caption = Main.Caption & " - Version: " & VERSION
    
    FolderPath.text = path_slash(registry_open("FolderPath"))
End Sub


Private Sub ScanStart_Click()
    scan_start
End Sub

Private Sub ScanStop_Click()
    run = 0
End Sub
