Attribute VB_Name = "module_database"
Option Explicit

Public Function database_open(path As String) As Long
    On Error GoTo error:
    
    Dim handler_id As Long
    
    If (sqlite3_open(path, handler_id) <> SQLITE_OK) Then
        Log "An error occured while opening/creating the database : " & path
        
        handler_id = 0
    End If
    
    database_open = handler_id
    
    Exit Function
error:
    database_open = 0
End Function
    
Public Function database_create_item(handler_id As Long) As Integer
    On Error GoTo error:
    
    Dim query_string As String
    Dim error_string As String
    Dim return_code As Integer
    
    query_string = "CREATE TABLE items("
    query_string = query_string & "Id TEXT,"
    query_string = query_string & "Character TEXT,"
    query_string = query_string & "Location TEXT,"
    query_string = query_string & "Path TEXT,"
    query_string = query_string & "Head NUMERIC,"
    query_string = query_string & "Version NUMERIC,"
    query_string = query_string & "Time NUMERIC,"
    query_string = query_string & "Checksum NUMERIC,"
    query_string = query_string & "Name TEXT,"
    query_string = query_string & "Weight NUMERIC,"
    query_string = query_string & "Price NUMERIC,"
    query_string = query_string & "Integrity0 NUMERIC,"
    query_string = query_string & "Integrity1 NUMERIC,"
    query_string = query_string & "Resistance0 NUMERIC,"
    query_string = query_string & "Resistance1 NUMERIC,"
    query_string = query_string & "Resistance2 NUMERIC,"
    query_string = query_string & "Resistance3 NUMERIC,"
    query_string = query_string & "Resistance4 NUMERIC,"
    query_string = query_string & "Resistance5 NUMERIC,"
    query_string = query_string & "Resistance6 NUMERIC,"
    query_string = query_string & "Resistance7 NUMERIC,"
    query_string = query_string & "Damage0 NUMERIC,"
    query_string = query_string & "Damage1 NUMERIC,"
    query_string = query_string & "Range NUMERIC,"
    query_string = query_string & "AttackSpeed NUMERIC,"
    query_string = query_string & "AttackRating NUMERIC,"
    query_string = query_string & "Critical NUMERIC,"
    query_string = query_string & "Absorb NUMERIC,"
    query_string = query_string & "Defence NUMERIC,"
    query_string = query_string & "Block NUMERIC,"
    query_string = query_string & "Speed NUMERIC,"
    query_string = query_string & "Potions NUMERIC,"
    query_string = query_string & "MagicMastery NUMERIC,"
    query_string = query_string & "ManaRegen NUMERIC,"
    query_string = query_string & "LifeRegen NUMERIC,"
    query_string = query_string & "StaminaRegen NUMERIC,"
    query_string = query_string & "IncreaseLife NUMERIC,"
    query_string = query_string & "IncreaseMana NUMERIC,"
    query_string = query_string & "IncreaseStamina NUMERIC,"
    query_string = query_string & "Level NUMERIC,"
    query_string = query_string & "Strength NUMERIC,"
    query_string = query_string & "Spirit NUMERIC,"
    query_string = query_string & "Talent NUMERIC,"
    query_string = query_string & "Dexterity NUMERIC,"
    query_string = query_string & "Health NUMERIC,"
    query_string = query_string & "UniqueItem NUMERIC,"
    query_string = query_string & "SpecAbsorb NUMERIC,"
    query_string = query_string & "SpecDefense NUMERIC,"
    query_string = query_string & "SpecSpeed NUMERIC,"
    query_string = query_string & "SpecBlock NUMERIC,"
    query_string = query_string & "SpecAttackSpeed NUMERIC,"
    query_string = query_string & "SpecCritical NUMERIC,"
    query_string = query_string & "SpecRange NUMERIC,"
    query_string = query_string & "SpecMagicMastery NUMERIC,"
    query_string = query_string & "SpecAttackRating NUMERIC,"
    query_string = query_string & "SpecDamage NUMERIC,"
    query_string = query_string & "SpecManaRegen NUMERIC,"
    query_string = query_string & "SpecLifeRegen NUMERIC,"
    query_string = query_string & "SpecStaminaRegen NUMERIC,"
    query_string = query_string & "Age NUMERIC"
    query_string = query_string & ")"
        
    sqlite_get_table handler_id, query_string, error_string, return_code
    
    If return_code <> SQLITE_OK Then
        Log "[database_create_item] (" & return_code & ") " & error_string
    End If
    
    database_create_item = return_code
    
    Exit Function
error:
    database_create_item = SQLITE_ERROR
End Function

Public Function database_query(handler_id As Long, query_string As String) As Variant
    On Error GoTo error:
    
    Dim error_string As String
    Dim return_code As Integer
    
    database_query = sqlite_get_table(handler_id, query_string, error_string, return_code)
    
    If handler_id <> SQLITE_OK Then
        Log "[database_query] (" & return_code & ") " & error_string
        
        database_query = Empty
    End If
    
    Exit Function
error:
    database_query = Empty
End Function
