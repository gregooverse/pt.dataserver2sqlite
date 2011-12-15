Attribute VB_Name = "module_datetime"
Option Explicit

' Building format
Public Function datetime_formating(ByVal format As String) As String
    On Error GoTo error:
    
    Dim day_raw, month_raw, year_raw, hour_raw, minute_raw, second_raw As String
    Dim day_zero, month_zero, year_two, hour_zero, minute_zero, second_zero As String
        
    second_raw = Second(Now)
    minute_raw = Minute(Now)
    hour_raw = Hour(Now)
    day_raw = Day(Now)
    month_raw = Month(Now)
    year_raw = Year(Now)
     
    second_zero = IIf(Len(second_raw) = 2, second_raw, "0" & second_raw)
    minute_zero = IIf(Len(minute_raw) = 2, minute_raw, "0" & minute_raw)
    hour_zero = IIf(Len(hour_raw) = 2, hour_raw, "0" & hour_raw)
    day_zero = IIf(Len(day_raw) = 2, day_raw, "0" & day_raw)
    month_zero = IIf(Len(month_raw) = 2, month_raw, "0" & month_raw)
    year_two = Right$(year_raw, 2)
    
    format = Replace(format, "yyyy", year_raw)
    format = Replace(format, "yy", year_two)
    format = Replace(format, "mm", month_zero)
    format = Replace(format, "m", month_raw)
    format = Replace(format, "dd", day_zero)
    format = Replace(format, "d", day_raw)
    format = Replace(format, "hh", hour_zero)
    format = Replace(format, "h", hour_raw)
    format = Replace(format, "nn", minute_zero)
    format = Replace(format, "n", minute_raw)
    format = Replace(format, "ss", second_zero)
    format = Replace(format, "s", second_raw)
    
    datetime_formating = format

    Exit Function
error:
    Log "[datetime_formating] Error while formating date: " & format
    datetime_formating = "datetime_error"
End Function
