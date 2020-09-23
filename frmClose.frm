VERSION 5.00
Object = "{5A1DEC06-02A4-11D5-B786-978568376651}#8.0#0"; "ProgressControl.ocx"
Begin VB.Form frmShut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "frmClose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer timIdle 
      Interval        =   50
      Left            =   1920
      Top             =   480
   End
   Begin VB.Frame framAuto 
      Caption         =   "Automatic Shutdown In"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      Begin DetailedProgressBar.ProgressBar dpbShut 
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "0 Seconds"
      End
   End
   Begin VB.Timer timClose 
      Interval        =   50
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer timAlarm 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   480
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
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblShut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shut down the computer ?"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmShut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================
Option Explicit
Option Base 1

Const HeightForBar = 1950
Const HeightOtherwise = 1230

Dim StartSec As Integer
Dim StartMin As Integer
Dim StartHour As Integer

Dim StartTick As Long

Dim CloseHour As Integer
Dim CloseMin As Integer
Dim CloseSec As Integer

Dim Delay As Integer

Dim ShutWin As Boolean
Dim CancelShut As Boolean

'Colour variables
Dim ColHour As Variant
Dim ColMin As Variant
Dim ColSec As Variant
Dim ColDot As Variant
Dim CAnBak As Variant
Dim CTmFon As Variant   'time font
Dim CTmBak As Variant   'time back ground
Dim CDyFon As Variant   'day font
Dim CDyBak As Variant   'day ...
Dim CDtFon As Variant   'date ..
Dim CDtBak As Variant   'date ..


Private Sub cmdNo_Click()
'don't shut-down the computer. Disable form
Call NotOnTop(Me)
Call EnableForms
frmShut.Visible = False
timAlarm.Enabled = False
timClose.Enabled = Week(Weekday(Date)).ShutWin
timIdle.Enabled = IdleShut
StartTick = 0
End Sub

Private Sub cmdYes_Click()
'shut down the computer
timClose.Enabled = False
Call frmHandsClk.SaveStatus
Call DoShutMethod(Method)
End
End Sub

Private Sub Form_Activate()
Dim MaxVal As Variant
Dim X As Integer
Dim Y As Integer

'hide the other forms
Call DisableForms
CancelShut = False


cmdYes.SetFocus

Call ShowProg

If (Week(Weekday(Date)).DelayOn) And (StartTick = 0) Then
    'set progress bar values and starting time
    StartTick = GetTickCount
    
    'set progressbar values
    dpbShut.Min = 0
    dpbShut.Value = 0
    dpbShut.Caption = Week(Weekday(Date)).DelayTime & " Seconds"
    
    'here when the computer is performing calculatons
    'it returns only a type of integer. Any number
    'produced that is greater will cause an error.
    MaxVal = CDbl(Week(Weekday(Date)).DelayTime) * 1000 'TypeName(55 * 1000))
    
    'convert second time to millisecond tick
    dpbShut.Max = MaxVal
End If

'make sure the box is seen
Call StayOnTop(Me)

'convert the position of the mouse in pixles
X = ((cmdYes.Width / 2) + frmShut.Left + cmdYes.Left)
Y = ((cmdYes.Height / 2) + frmShut.Top + cmdYes.Top + (cmdYes.Height))

'move the mouse to the "Yes" button
Call MoveMouseTo(X, Y)
End Sub

Private Sub Form_Load()
'stop the owner search if active
'Searching = False

'if no shutdown method is specified, then
'default to Shut Down
If Method = "" Then
    Method = "Shut Down"
End If

'centre label and display shut down method.
lblShut.Caption = Method & " the computer?"
lblShut.Left = (frmShut.Width / 2) - (lblShut.Width / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'PROPERTIES OF THE "CANCEL" VARIABLE
'==================================================
'Cancel                 An integer. Setting this
'                       argument to any value other
'                       than 0 stops the QueryUnload
'                       event in all loaded forms
'                       and stops the form and
'                       application from closing.
'==================================================

'CAUSES AND VALUES OF "UNLOADING" USED BY
'THE "UNLOADMODE" PARAMETER
'==================================================
'vbFormControlMenu      0   The user chose the Close
'                           command from the Control
'                           menu on the form.
'vbFormCode             1   The Unload statement is
'                           invoked from code.
'vbAppWindows           2   The current Microsoft
'                           Windows operating
'                           environment session is
'                           ending.
'vbAppTaskManager       3   The Microsoft Windows
'                           Task Manager is closing
'                           the application.
'vbFormMDIForm          4   An MDI child form is
'                           closing because the MDI parent
'                           form is closing.
'==================================================

Const CancelUnload = 1

'This form may not be unloaded for the simple reason
'that the user should not be able to exit without
'answering 'yes' or 'no'.
'Hence the "Cancel" variable to set to stop the
'unloading.
If (UnloadMode <> vbFormCode) And (UnloadMode <> vbAppWindows) Then
    Cancel = CancelUnload
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'This procedure is called when the program starts to
'unload.

If CancelShut Then
    Call EnableForms
End If
End Sub

Private Sub timAlarm_Timer()
'if "No" wasn't pressed before the delay time runs
'out, then shut down the computer.

Static SecondsLeft As Integer

If ((GetTickCount - StartTick) >= (CDbl(Week(Weekday(Date)).DelayTime) * 1000)) Then
    'time ran out...
    dpbShut.Caption = "0 Seconds"
    dpbShut.Value = dpbShut.Max
    dpbShut.Refresh
    
    timAlarm.Enabled = False
    timClose.Enabled = False
    
    Call NotOnTop(frmShut)
    frmShut.Visible = False
    
    'Unload Me
    'Call frmHandsClk.SaveStatus
    Call DoShutMethod(Method)
    End
    Stop '<-- used for testing.
Else
    'show remaining time in progres bar
    SecondsLeft = (Week(Weekday(Date)).DelayTime - ((GetTickCount - StartTick) \ 1000))
    If SecondsLeft <> 1 Then
        dpbShut.Caption = SecondsLeft & " Seconds"
    Else
        dpbShut.Caption = SecondsLeft & " Second"
    End If
    dpbShut.Value = (GetTickCount - StartTick) + 1
End If
End Sub

Private Sub timClose_Timer()
Static DayNum As Integer

If DayNum = 0 Then
    DayNum = Weekday(Date)
End If

'if the time matches the time to shut down at today, or the computer
'has been idle for X seconds, then
If ((Hour(Time) = Week(DayNum).CloseHour) And (Minute(Time) = Week(DayNum).CloseMin) And (Second(Time) = Week(DayNum).CloseSec)) Then
    If Week(DayNum).DelayOn Then
        timIdle.Enabled = False
        timAlarm.Enabled = True
        Call StayOnTop(Me)

        'audio warning
        Beep
    End If
    
    DoEvents
    'visible warning
    frmShut.Visible = True
    
    timClose.Enabled = False
End If
End Sub

Public Sub ShowProg()
'shows or hides the progress bar

If Week(Weekday(Date)).DelayOn Then
    frmShut.Height = HeightForBar
    Searching = False
Else
    frmShut.Height = HeightOtherwise
End If
End Sub

Private Sub DisableForms()
frmShut.Visible = True

'this hides the other forms whil this form is active
Unload frmAbout
Unload frmBack
Unload frmOptions
Unload frmScheme


frmHandsClk.Visible = False
frmHandsClk.timHand.Enabled = False
frmHandsClk.timDigital.Enabled = False
End Sub

Private Sub EnableForms()
'This shows only the main form
frmShut.Visible = False

frmHandsClk.Visible = True
frmHandsClk.timHand.Enabled = True
frmHandsClk.timDigital.Enabled = True
frmHandsClk.Show
End Sub

Private Sub timIdle_Timer()
'Dim ComputerIdleTime As String
Dim Difference As Long

Const SecInDay = 86400 ' The no. of seconds in a day

ComputerIdleTime = PSTime
If ComputerIdleTime = "" Then
    PSTime = SecToTime(0)
    'ComputerIdleTime = SecToTime(0)
    'PSTime = ComputerIdleTime
End If

'if the predicted time is less than the currentv
Difference = DateDiff("s", Time, ComputerIdleTime)
If (Difference < 0) And (Difference > -5) Then
    'find the new predicted idle time
    PSTime = PredictIdle(IdleTimeInSec - TimeToSec(GetTimeIdle))
    'ComputerIdleTime = PSTime
End If

'get the difference in seconds between the current time and
'the predicted idle time
Difference = DateDiff("s", Time, PSTime)
If ((Difference < 0) And (Difference > -5)) And (IdleShut) Then
    'find the new predicted idle time and go through this timer again
    
    'save settings
    Call frmHandsClk.SaveStatus
    
    timClose.Enabled = False
    timAlarm.Enabled = True
    Call StayOnTop(Me)

    'audio warning
    Beep
    
    DoEvents
    'visible warning
    frmShut.Visible = True
    
    timIdle.Enabled = False
End If
End Sub
