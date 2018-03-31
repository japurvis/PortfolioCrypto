Attribute VB_Name = "UTC"
Option Explicit

'---------------------------------------------------------------------
' Declarations must be at the top -- see below
'---------------------------------------------------------------------
Public Declare Function SystemTimeToFileTime Lib _
  "Kernel32" (lpSystemTime As SYSTEMTIME, _
  lpFileTime As FILETIME) As Long

Public Declare Function FileTimeToLocalFileTime Lib _
  "Kernel32" (lpLocalFileTime As FILETIME, _
  lpFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib _
  "Kernel32" (lpFileTime As FILETIME, lpSystemTime _
  As SYSTEMTIME) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'---------------------------------------------------------------------
' Convert ISO8601 dateTimes to Excel Dates
'---------------------------------------------------------------------
Public Function ISODATE(iso As String)
    ' Find location of delimiters in input string
    Dim tPos As Integer: tPos = InStr(iso, "T")
    If tPos = 0 Then tPos = Len(iso) + 1
    Dim zPos As Integer: zPos = InStr(iso, "Z")
    If zPos = 0 Then zPos = InStr(iso, "+")
    If zPos = 0 Then zPos = InStr(tPos, iso, "-")
    If zPos = 0 Then zPos = Len(iso) + 1
    If zPos = tPos Then zPos = tPos + 1

    ' Get the relevant parts out
    Dim datePart As String: datePart = Mid(iso, 1, tPos - 1)
    Dim timePart As String: timePart = Mid(iso, tPos + 1, zPos - tPos - 1)
    Dim dotPos As Integer: dotPos = InStr(timePart, ".")
    If dotPos = 0 Then dotPos = Len(timePart) + 1
    timePart = Left(timePart, dotPos - 1)

    ' Have them parsed separately by Excel
    Dim d As Date: d = DateValue(datePart)
    Dim t As Date: If timePart <> "" Then t = TimeValue(timePart)
    Dim dt As Date: dt = d + t

    ' Add the timezone
    Dim tz As String: tz = Mid(iso, zPos)
    If tz <> "" And Left(tz, 1) <> "Z" Then
        Dim colonPos As Integer: colonPos = InStr(tz, ":")
        If colonPos = 0 Then colonPos = Len(tz) + 1

        Dim minutes As Integer: minutes = CInt(Mid(tz, 2, colonPos - 2)) * 60 + CInt(Mid(tz, colonPos + 1))
        If Left(tz, 1) = "+" Then minutes = -minutes
        dt = DateAdd("n", minutes, dt)
    End If

    ' Return value is the ISO8601 date in the local time zone
    dt = ToUTC(dt)
    ISODATE = dt
End Function

'---------------------------------------------------------------------
' Got this function to convert local date to UTC date from
' http://excel.tips.net/Pages/T002185_Automatically_Converting_to_GMT.html
'---------------------------------------------------------------------
Public Function ToUTC(dteTime As Date) As Date
    Dim infile As FILETIME
    Dim outfile As FILETIME
    Dim insys As SYSTEMTIME
    Dim outsys As SYSTEMTIME

    insys.wYear = CInt(Year(dteTime))
    insys.wMonth = CInt(Month(dteTime))
    insys.wDay = CInt(Day(dteTime))
    insys.wHour = CInt(Hour(dteTime))
    insys.wMinute = CInt(Minute(dteTime))
    insys.wSecond = CInt(Second(dteTime))

    Call SystemTimeToFileTime(insys, infile)
    Call FileTimeToLocalFileTime(infile, outfile)
    Call FileTimeToSystemTime(outfile, outsys)

    ToUTC = CDate(outsys.wMonth & "/" & _
      outsys.wDay & "/" & _
      outsys.wYear & " " & _
      outsys.wHour & ":" & _
      outsys.wMinute & ":" & _
      outsys.wSecond)
End Function

'---------------------------------------------------------------------
' Got this function to convert local date to UTC date from
' http://excel.tips.net/Pages/T002185_Automatically_Converting_to_GMT.html
'---------------------------------------------------------------------
Public Function FromUTC(d As Date) As Date
    Dim dateLocalFile As FILETIME
    Dim dateFile As FILETIME
    Dim dateLocalSystem As SYSTEMTIME
    Dim dateSystem As SYSTEMTIME

    dateLocalSystem.wYear = CInt(Year(d))
    dateLocalSystem.wMonth = CInt(Month(d))
    dateLocalSystem.wDay = CInt(Day(d))
    dateLocalSystem.wHour = CInt(Hour(d))
    dateLocalSystem.wMinute = CInt(Minute(d))
    dateLocalSystem.wSecond = CInt(Second(d))

    Call SystemTimeToFileTime(dateLocalSystem, dateLocalFile)
    Call FileTimeToLocalFileTime(dateLocalFile, dateFile)
    Call FileTimeToSystemTime(dateFile, dateSystem)

    FromUTC = CDate(dateSystem.wMonth & "/" & _
      dateSystem.wDay & "/" & _
      dateSystem.wYear & " " & _
      dateSystem.wHour & ":" & _
      dateSystem.wMinute & ":" & _
      dateSystem.wSecond)
End Function

Function DateToUnixTime(dt) As Long
    DateToUnixTime = DateDiff("s", "1/1/1970", dt)
End Function

Function UnixTimeToDate(ts As Long) As Date
    'http://www.vbforums.com/showthread.php?513727-RESOLVED-Convert-Unix-Time-to-Date&p=3168062&viewfull=1#post3168062
    Dim intDays As Integer, intHours As Integer, intMins As Integer, intSecs As Integer
    
    intDays = ts \ 86400
    intHours = (ts Mod 86400) \ 3600
    intMins = (ts Mod 3600) \ 60
    intSecs = ts Mod 60
    
    UnixTimeToDate = DateSerial(1970, 1, intDays + 1) + TimeSerial(intHours, intMins, intSecs)
End Function


