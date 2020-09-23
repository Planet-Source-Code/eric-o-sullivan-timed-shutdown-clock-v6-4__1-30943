VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timed Shutdown Options"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   FillColor       =   &H80000006&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framDaily 
      Caption         =   "Daily Options"
      Height          =   3495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3375
      Begin VB.Frame framShut 
         Caption         =   "Timed Shutdown At"
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3135
         Begin VB.TextBox txtHour 
            Height          =   285
            Left            =   600
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "0"
            ToolTipText     =   "Hours"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "0"
            ToolTipText     =   "Minutes"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtSec 
            Height          =   285
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "0"
            ToolTipText     =   "Seconds"
            Top             =   960
            Width           =   375
         End
         Begin VB.ComboBox cmbWeek 
            Height          =   315
            ItemData        =   "frmOptions.frx":030A
            Left            =   360
            List            =   "frmOptions.frx":0326
            TabIndex        =   0
            Text            =   "Weekday"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Image imgTick 
            Height          =   1065
            Left            =   2040
            Picture         =   "frmOptions.frx":0371
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Time 24H"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Time in 24 hours"
            Top             =   960
            Width           =   375
         End
      End
      Begin VB.Frame framDelay 
         Caption         =   "Default Shutdown Delay"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   3135
         Begin VB.CheckBox chkOn 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirm Shut Down"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1755
         End
         Begin VB.TextBox txtDelay 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1520
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "15"
            ToolTipText     =   "Delay time in seconds"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CheckBox chkShut 
            Alignment       =   1  'Right Justify
            Caption         =   "Shut Down On/Off"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1755
         End
         Begin VB.Image imgTock 
            Height          =   1095
            Left            =   2040
            Picture         =   "frmOptions.frx":067B
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblSeconds 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Wait For"
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s"
            Height          =   195
            Left            =   1905
            TabIndex        =   13
            Top             =   1140
            Width           =   75
         End
      End
   End
   Begin VB.Frame framMethod 
      Caption         =   "Shutdown Details"
      Height          =   3495
      Left            =   3480
      TabIndex        =   8
      Top             =   0
      Width           =   2295
      Begin VB.CheckBox chkPrev 
         Alignment       =   1  'Right Justify
         Caption         =   "Prevent Other Apps Closing Windows"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.Frame framIdle 
         Caption         =   "Idle Time"
         Height          =   1335
         Left            =   120
         TabIndex        =   17
         Tag             =   " at all. Compiling a large database is not feasable for a program this small, hence the option."
         Top             =   2040
         Width           =   2055
         Begin VB.CheckBox chkIdle 
            Alignment       =   1  'Right Justify
            Caption         =   "Idle Shutdown On"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbMin 
            Height          =   315
            ItemData        =   "frmOptions.frx":0985
            Left            =   1200
            List            =   "frmOptions.frx":0A3D
            TabIndex        =   20
            Text            =   "00"
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cmbHour 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B31
            Left            =   360
            List            =   "frmOptions.frx":0B7D
            TabIndex        =   19
            Text            =   "01"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M"
            Height          =   195
            Left            =   1050
            TabIndex        =   22
            Top             =   840
            Width           =   135
         End
         Begin VB.Label lblIdleHour 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H"
            Height          =   195
            Left            =   225
            TabIndex        =   21
            Top             =   840
            Width           =   120
         End
         Begin VB.Label lblIdle 
            BackStyle       =   0  'Transparent
            Caption         =   "Shutdown After;"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.ComboBox cmbMethod 
         Height          =   315
         ItemData        =   "frmOptions.frx":0BE1
         Left            =   720
         List            =   "frmOptions.frx":0BF4
         TabIndex        =   7
         Text            =   "Shut Down"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image imgIdle 
         Height          =   735
         Left            =   600
         Picture         =   "frmOptions.frx":0C2E
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblExplanation 
         BackColor       =   &H0080FFFF&
         Caption         =   $"frmOptions.frx":1070
         Height          =   60
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Dim FormLoading As Boolean

Private Sub chkIdle_Click()
'wether or not to shutdown the computer after a specified period
'of idleness.

If FormLoading Then
    'don't do anything if the form is just setting the initial parameters
    Exit Sub
End If

If chkIdle.Value = 1 Then
    IdleShut = True
    'call trackidletime
Else
    IdleShut = False
    'call endtracking
End If

frmHandsClk.mnuFileIdle.Checked = IdleShut

'save new change
'Call SaveStatus
End Sub

Private Sub chkOn_Click()
'Delay on. If checked, then a box will appear at
'appionted time to confirm shutdown. Only waits for
'an answer for a set number of seconds.

'If off then the computer forces shutdown at
'appointed time.
If chkOn.Value = 1 Then
    frmOptions.Tag = "On"
    lblSeconds.Enabled = True
    txtDelay.Enabled = True
Else
    frmOptions.Tag = "Off"
    lblSeconds.Enabled = False
    txtDelay.Enabled = False
End If

'save changes
Call SaveWeek((cmbWeek.ListIndex), "DelayOn", chkOn.Value)
'Call SaveStatus
End Sub

Private Sub chkPrev_Click()
'stop other apps from closing windows?

If FormLoading Then
    Exit Sub
End If

PreventShut = Not PreventShut
frmHandsClk.mnuFileAdvPrev.Checked = PreventShut
'Call frmHandsClk.SaveStatus
End Sub

Private Sub chkShut_Click()
'shut down computer ? yes/no

ShutWin = Not ShutWin
Call SaveWeek(cmbWeek.ListIndex, "ShutWin", GetVal(chkShut.Value))
'Call SaveStatus
End Sub

Private Sub cmbHour_Click()

'if the total time is 0 then, the minute value is 1
If (Val(cmbHour.Text) = 0) And (Val(cmbMin.Text) = 0) Then
    cmbMin.ListIndex = 1
End If

IdleTimeInSec = (Val(cmbHour.Text) * 3600) + (Val(cmbMin.Text) * 60)
'Call SaveStatus

'find the new predicted idle time
PSTime = PredictIdle(IdleTimeInSec)
ComputerIdleTime = PSTime
End Sub

Private Sub cmbHour_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cmbHour_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbMin_Click()

'if the total time is 0 then, the minute value is 1
If (Val(cmbHour.Text) = 0) And (Val(cmbMin.Text) = 0) Then
    cmbMin.ListIndex = 1
End If

IdleTimeInSec = (Val(cmbHour.Text) * 3600) + (Val(cmbMin.Text) * 60)
'Call SaveStatus

'find the new predicted idle time
PSTime = PredictIdle(IdleTimeInSec)
ComputerIdleTime = PSTime
End Sub

Private Sub cmbMin_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cmbMin_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbMethod_Click()
Method = cmbMethod.Text
'Call SaveStatus
End Sub

Private Sub cmbMethod_KeyDown(KeyCode As Integer, Shift As Integer)
'stop the user from changing the data
Const Delete = 46

If KeyCode = Delete Then KeyCode = 0
End Sub

Private Sub cmbMethod_KeyPress(KeyAscii As Integer)
'stop the user from changing the data

KeyAscii = 0
End Sub

Private Sub cmbWeek_Click()
'select a particular days details and display them

Dim Day As Integer
Dim Hour As Integer
Dim Min As Integer
Dim Sec As Integer
Dim DelTim As Integer
Dim DelOn As Boolean
Dim Shut As Boolean

'day 1 means the settings are for all days
'day 2 = Sunday, day 8 = saturday
If (cmbWeek.ListIndex + 1) <> 1 Then
    'display settings form selected day
    Week(cmbWeek.ListIndex).Day = cmbWeek.ListIndex
    txtHour.Text = Week(cmbWeek.ListIndex).CloseHour
    txtMin.Text = Week(cmbWeek.ListIndex).CloseMin
    txtSec.Text = Week(cmbWeek.ListIndex).CloseSec
    chkOn.Value = GetBool(Week(cmbWeek.ListIndex).DelayOn)    'convert to true/false true = -1, false = 0
    txtDelay.Text = Week(cmbWeek.ListIndex).DelayTime
    chkShut.Value = GetBool(Week(cmbWeek.ListIndex).ShutWin)  'convert to true/false true = -1, false = 0
Else
    'use current settings for all days.
    Hour = Val(txtHour.Text)
    Min = Val(txtMin.Text)
    Sec = Val(txtSec.Text)
    DelTim = Val(txtDelay.Text)
    DelOn = GetBool(chkOn.Value)
    Shut = GetBool(chkShut.Value)
    
    For Day = 1 To 7
        Week(Day).CloseHour = Hour
        Week(Day).CloseMin = Min
        Week(Day).CloseSec = Sec
        Week(Day).DelayTime = DelTim
        Week(Day).DelayOn = DelOn
        Week(Day).ShutWin = Shut
    Next Day
End If

'Call SaveStatus
End Sub

Private Sub cmbWeek_KeyPress(KeyAscii As Integer)
'if user presses any key except tab, then ignore
'ASCII #9 = Tab key
If KeyAscii <> 9 Then
    KeyAscii = 0
End If
End Sub

Private Sub Form_Activate()
'starting program
Dim Day As Integer

'set current day
Day = Weekday(Date)

'display results and show current day in combo-box
cmbWeek.ListIndex = Day
If Method = "" Then
    cmbMethod.ListIndex = GetIndex(Method)
Else
    cmbMethod.Text = Method
End If

'time to shut down
txtHour.Text = Week(Day).CloseHour
txtMin.Text = Week(Day).CloseMin
txtSec.Text = Week(Day).CloseSec

'delay time to ask for confirmation
txtDelay.Text = Week(Day).DelayTime

'close at set time
chkOn.Value = GetBool(Week(Day).DelayOn)

'idle values
chkIdle.Value = GetBool(IdleShut)
cmbHour.Text = Format(Hour(SecToTime(IdleTimeInSec)), "00")
cmbMin.Text = Format(Minute(SecToTime(IdleTimeInSec)), "00")

'prevent other apps from closing windows on/off
chkPrev.Value = GetBool(PreventShut)

'let other procedures know that the form has finished loading
FormLoading = False
End Sub

Private Sub Form_Load()
'let other procedures know that the form is loading
FormLoading = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'this procedure is called when the prorgam receives
'a request to shut down. The program is not actually
'shutting down yet. For more information, please see
'the Form_QueryUnload in the frmShut screen.


'only save the changes if the form is unloading because of
'code we have written - otherwise, don't risk data corruption.
If (UnloadMode = vbFormCode) Or (UnloadMode = vbFormControlMenu) Then
    Call frmHandsClk.SaveStatus
End If
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
'data validation for delay seconds
KeyAscii = EnterNum(txtDelay, KeyAscii, "Delay")
'Call SaveStatus
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
'advanced data validation for hour
txtHour.SelText = ""

KeyAscii = GetNum(KeyAscii)
If KeyAscii <> 0 Then
    'if the new value is greater than 24 then
    If ((Val(Left(txtHour.Text, 1)) >= 2) And (KeyAscii > 51)) Or ((Val(txtHour.Text) > 2) And (KeyAscii >= 48)) Then
        'reset to 23
        txtHour.Text = "23"
        txtHour.Refresh
        
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    Else
        'display new value
        If (Len(txtHour.Text) = 1) And (txtHour.Text <> "0") And (KeyAscii = 8) Then
            txtHour.Text = "0"
        Else
            If txtHour.Text = "0" Then txtHour = ""
        End If
        
        txtHour.Text = txtHour.Text & Trim(Str(KeyAscii - 48))
        
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End If

txtHour.SelStart = Len(txtHour)

'save the changes
Call SaveWeek((cmbWeek.ListIndex), "Hour", Val(txtHour.Text))
CloseHour = Val(txtHour.Text)

'Call SaveStatus
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
'data validation for minutes
KeyAscii = EnterNum(txtMin, KeyAscii, "Min")
'Call SaveStatus
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
'data validation for seconds
KeyAscii = EnterNum(txtSec, KeyAscii, "Sec")
'Call SaveStatus
End Sub

Public Sub SaveStatus()
'This saves the current settings. Obviously :P

'testing if this procedure can be removed to prevent data
'replication.
'Call frmHandsClk.SaveStatus
End Sub

Public Function EnterNum(txtEnter As TextBox, KeyAscii As Integer, NewVal As String) As Integer
'This function controls the data validation
txtEnter.SelText = ""

KeyAscii = GetNum(KeyAscii)
If KeyAscii <> 0 Then
    'if new value if greater than 59 then
    If (Val(Left(txtEnter.Text, 1)) >= 6) And (KeyAscii >= 48) Then
        txtEnter.Text = "59"
        txtEnter.Refresh
        
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    Else
        'display new value
        If (Len(txtEnter.Text) = 1) And (txtEnter.Text <> "0") And (KeyAscii = 8) Then
            'if the last number is deleted, display
            '"0"
            txtEnter.Text = "0"
        Else
            'remove leading zero
            If txtEnter.Text = "0" Then txtEnter = ""
        End If

        txtEnter.Text = txtEnter.Text & Right(Str(KeyAscii - 48), 1)
        
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End If

txtEnter.SelStart = Len(txtEnter)

'save changes depending on which box was validated
Select Case LCase(NewVal)
Case "delay"
    Delay = Val(txtEnter.Text)
    Call SaveWeek((cmbWeek.ListIndex), "DelayTime", Val(txtEnter.Text))
Case "min"
    Call SaveWeek((cmbWeek.ListIndex), "Min", Val(txtEnter.Text))
    CloseMin = Val(txtEnter.Text)
Case "sec"
    Call SaveWeek((cmbWeek.ListIndex), "Sec", Val(txtEnter.Text))
    CloseSec = Val(txtEnter.Text)
End Select

'save settings
'frmHandsClk.Tag = Delay & "," & CloseHour & "," & CloseMin & "," & CloseSec & "," & ShutWin & "," & ColHour & "," & ColMin & "," & ColSec & "," & ColDot & "," & CAnBak & "," & CTmFon & "," & CTmBak & "," & CDyFon & "," & CDyBak & "," & CDtFon & "," & CDtBak

EnterNum = KeyAscii
End Function

Public Function GetVal(Number As Integer) As Boolean
'for the checkboxes. Convert the numeric value
'to boolean
If Number = 0 Then
    GetVal = False
Else
    GetVal = True
End If
End Function

Public Function GetBool(Value As Boolean) As Integer
'This converts the boolean value to numeric.
If Value Then
    GetBool = 1
Else
    GetBool = 0
End If
End Function

Public Sub SaveWeek(Element As Integer, Where As String, Value As Integer)
'save any or all days setting for shut-down
If Element <> 0 Then
    'element "0" is "[All]" in the combo box
    Select Case LCase(Where)
    Case "hour"
        Week(Element).CloseHour = Value
    Case "min"
        Week(Element).CloseMin = Value
    Case "sec"
        Week(Element).CloseSec = Value
    Case "delayon"
        Week(Element).DelayOn = GetVal(Value)
    Case "delaytime"
        Week(Element).DelayTime = Value
    Case "shutwin"
        Week(Element).ShutWin = GetVal(Value)
        frmHandsClk.mnuFileTim.Checked = GetVal(Value)
        
        'see frmShut.timClose_Timer()
        If GetVal(Value) Then
            frmOptions.Tag = "On"
        Else
            frmOptions.Tag = "Off"
        End If
    End Select
Else
    'save settings for all days
    For Element = 1 To 7
        Select Case LCase(Where)
        Case "hour"
            Week(Element).CloseHour = Value
        Case "min"
            Week(Element).CloseMin = Value
        Case "sec"
            Week(Element).CloseSec = Value
        Case "delayon"
            Week(Element).DelayOn = GetVal(Value)
        Case "delaytime"
            Week(Element).DelayTime = Value
        Case "shutwin"
            frmHandsClk.mnuFileTim.Checked = ShutWin
            Week(Element).ShutWin = GetVal(Value)
        End Select
    Next Element
End If
End Sub

Private Function GetIndex(ShutMethod As String)
'This procedure returns the index of the named
'shut down method for the list box

Dim Index As Integer

'0 - Shut Down
'1 - Power Down
'2 - Force Close
'3 - Restart
'4 - Log Off

Select Case LCase(ShutMethod)
Case ""
    Index = 0
Case "shut down"
    Index = 0
Case "power down"
    Index = 1
Case "force close"
    Index = 2
Case "restart"
    Index = 3
Case "log off"
    Index = 4
End Select

GetIndex = Index
End Function
