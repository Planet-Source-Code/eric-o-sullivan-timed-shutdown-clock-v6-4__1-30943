VERSION 5.00
Begin VB.Form frmShut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2490
   Icon            =   "frmShut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2490
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer TimClose 
      Interval        =   1
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer timAlarm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblShut 
      BackStyle       =   0  'Transparent
      Caption         =   "Shut down the computer ?"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmShut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartSec As Byte
Dim StartMin As Byte
Dim StartHour As Byte


Private Sub cmdDisable_Click()
DisableCtrlAltDel
End Sub

Private Sub cmdEnable_Click()
EnableCtrlAltDel
End Sub

Private Sub cmdForce_Click()
WINForceClose
'screen.
End Sub

Private Sub cmdRestart_Click()
WINReboot
End Sub

Private Sub cmdShut_Click()
WINShutdown
End Sub

Public Function Start(ByVal Sec As Byte) As Byte
Const Delay = 15
Sec = Sec + Delay
Sec = Sec Mod 60
Start = Sec
End Function

Private Sub Alarm()
If StartHour = 0 Then
    StartSec = Second(Time)
    StartHour = Hour(Time)
End If

End Sub

Private Sub cmdNo_Click()
frmShut.Visible = False
StartSec = 0
StartHour = 0
timAlarm.Enabled = False
End Sub

Private Sub cmdYes_Click()
WINShutdown
End Sub

Private Sub timAlarm_Timer()

If Second(Time) = Start(StartSec) Then
    'Call MsgBox("Shut down ?", vbExclamation + vbYesNo, "Close")
    WINShutdown
    timAlarm.Enabled = False
    TimClose.Enabled = False
    End
End If
End Sub

Private Sub TimClose_Timer()
Const CloseHour = 3
Const CloseMinute = 0
Const CloseSecond = 0
Dim Response As Integer

If (Hour(Time) = CloseHour) And (Minute(Time) = CloseMinute) And (Second(Time) = CloseSecond) Then
    Alarm
    timAlarm.Enabled = True
    frmShut.Visible = True
    Beep
End If
End Sub
