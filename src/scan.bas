Attribute VB_Name = "module_scan"
Option Explicit

Public Sub scan_start()
    On Error GoTo error:
    
    Dim i As Integer
    Dim File As String
    Dim path As String
    Dim Path2 As String
    Dim Path3 As String
    Dim handler_id As Long
    
    Main.ScanStart.Enabled = False
    Main.ScanStop.Enabled = True
    Main.ProgressDat.value = 0
    Main.ProgressWar.value = 0
    
    DoEvents
    
    Main.run = 1
    Main.dats = 0
    Main.wars = 0
    Main.items = 0
    
    Log "Scan started"
    
    Log "Opening database file"
    
    handler_id = database_open(path_now)
    
    If handler_id <> 0 Then
    
        Log "Creating item table"
        
        database_create_item handler_id
    
        path = path_slash(Main.FolderPath.text & "userdata")
        
        If folder_exists(path) Then
            
            Log "Processing userdata"
            
            For i = 0 To 255
                Main.ProgressDat.value = i
                DoEvents
                
                If Main.run = 0 Then
                    Log "Scan aborted"
                    Main.ScanStart.Enabled = True
                    Main.ScanStop.Enabled = False
                    Exit Sub
                End If
                
                Path2 = path_slash(path & i)
                
                If folder_exists(Path2) Then
                
                    Path3 = Path2 & "*.dat "
                
                    File = Dir$(Path3)
                    
                    Do While Len(File)
                        If Right$(File, 4) = ".dat" Then
                            dat_process Path2 & File, handler_id
                            Main.dats = Main.dats + 1
                        End If
                        
                        File = Dir$
                    Loop
                Else
                    
                    Log Path2 & " not found"
                    
                End If
            Next
        Else
            
            Log path & " not found"
            
        End If
        
        path = path_slash(Main.FolderPath.text & "warehouse")
        
        If folder_exists(path) Then
            
            Log "Processing warehouse"
            
            For i = 0 To 255
                Main.ProgressWar.value = i
                DoEvents
                
                If Main.run = 0 Then
                    Log "Scan aborted"
                    Main.ScanStart.Enabled = True
                    Main.ScanStop.Enabled = False
                    Exit Sub
                End If
                
                Path2 = path_slash(path & i)
                
                If folder_exists(Path2) Then
                
                    Path3 = Path2 & "*.war"
                
                    File = Dir$(Path3)
                    
                    Do While Len(File)
                        If Right$(File, 4) = ".war" Then
                            war_process Path2 & File, handler_id
                            Main.wars = Main.wars + 1
                        End If
                        
                        File = Dir$
                    Loop
                Else
                    
                    Log Path2 & " not found"
                    
                End If
            Next
        Else
            
            Log path & " not found"
            
        End If
            
        sqlite3_close handler_id
        
        Log "Scan completed"
        Log "`- .dat: " & Main.dats
        Log "`- .war: " & Main.wars
        Log "`- Items: " & Main.items
        
        Main.ScanStart.Enabled = True
        Main.ScanStop.Enabled = False
        
    End If
    
    Exit Sub
error:
    Log "[scan_start] An error has occured while scanning"
    Main.ScanStart.Enabled = True
    Main.ScanStop.Enabled = False
End Sub
