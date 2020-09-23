Attribute VB_Name = "SystemTime"
'This module will let the programmer set the time on the
'local machine.
'
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================

Private Type SystemTime
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Declare Function SetLocalTime Lib "kernel32.dll" (lpSystemTime As SystemTime) As Long

Public Sub SetNewTime(NewHour As Integer, NewMinute As Integer, NewSecond As Integer)
' Set the system time to
' March 26, 1987 20:45:00 Greenwich Time
Dim SetTime As SystemTime
Dim RetVal As Long

SetTime.wHour = NewHour
SetTime.wMinute = NewMinute
SetTime.wSecond = NewSecond
SetTime.wMilliseconds = 0
SetTime.wDay = Day(Date)
SetTime.wMonth = Month(Date)
SetTime.wYear = Year(Date)

' Set time and date.
RetVal = SetLocalTime(SetTime)
End Sub

