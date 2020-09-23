Attribute VB_Name = "IdleTime"
'this module returns the amount of time the computer has been idle
'for since the GetTimeIdle function has been called last.

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Public Declare Function TrackIdleTime Lib "IdleTime" () As Integer
'Public Declare Sub EndTracking Lib "IdleTime" ()
'Public Declare Function GetIdleTime Lib "IdleTime" () As Long

Private Type PointAPI
        X As Long
        Y As Long
End Type

Private LastPos As PointAPI
Private LastKey(256) As Long
Private StartTick As Long

Public Function GetTimeIdle() As String
'this function returns the amount of time the computer has been
'idle for in Time format (hh:mm:ss)

Dim Counter As Long
Dim NowPos As PointAPI
Dim Result As Long
Dim Ticks As Long
Dim KeyVal(256) As Long

'get the current mouse position
Result = GetCursorPos(NowPos)

'check if the mouse has moved
If (NowPos.X <> LastPos.X) Or (NowPos.Y <> LastPos.Y) Then
    'reset idletime
    'call endtracking
    'call trackidletime
    Ticks = 0 'GetTickCount / 1000
    
    'set the lastposition
    LastPos.X = NowPos.X
    LastPos.Y = NowPos.Y
Else
    'check if a key was pressed
    Ticks = GetIdleTime / 1000
End If

GetTimeIdle = SecToTime(Ticks)
End Function

Public Function SecToTime(Seconds As Long) As String
'This function will convert the time in seconds to a time format
'(hh:mm:ss)

SecToTime = Format((Seconds \ 3600), "00") & ":" & Format((Seconds \ 60) Mod 60, "00") & ":" & Format(Seconds Mod 60, "00")
End Function

Public Function TimeToSec(Time As String) As Long
'This function will convert the time into the number of seconds in
'that time.

If Time <> "" Then
    TimeToSec = (Hour(Time) * 3600) + (Minute(Time) * 60) + (Second(Time))
End If
End Function

Public Function PredictIdle(TotalIdleSec As Long) As String
'This function will predict the time the computer should shutdown
'at after the specified period of idle time (in 24 hour format)

Dim SecToGo As Long
Dim PHour As Long
Dim PMin As Long
Dim PSec As Long
Dim CurTime As String

SecToGo = IdleTimeInSec - TimeToSec(GetTimeIdle)

'we don't want the time to change while we are working with it
CurTime = Time

PSec = Abs((SecToGo + Second(CurTime)) Mod 60)
PMin = Abs(((SecToGo / 60) + Minute(CurTime)) Mod 60)
PHour = Abs(((SecToGo / 3600) + Hour(CurTime)) Mod 24)


PredictIdle = FormatTime(PHour, PMin, PSec) 'Format(PHour, "00") & ":" & Format(PMin, "00") & ":" & Format(PSec, "00")
'Debug.Print PredictIdle
End Function
Public Function InvalidIdleTime(Seconds As Long) As Long
'This will return a valid idle time in seconds. The idle time is
'only invalid if the number of idle seconds is less than one minute.
'An invalid idle time will return a time of one hour (3600 seconds).

If Seconds < 60 Then
    'invalid time
    InvalidIdleTime = 3600 'one hour
Else
    InvalidIdleTime = Seconds
End If
End Function

Public Function FormatTime(fHours As Long, fMinutes As Long, fSeconds As Long) As String
'This will force the time to a specified format
FormatTime = Format(fHours, "00") & ":" & Format(fMinutes, "00") & ":" & Format(fSeconds, "00")
End Function
