VERSION 5.00
Begin VB.Form frmSetTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set System Time"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmSetTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framTime 
      Caption         =   "Change Setting"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   2655
      Begin VB.Timer timCurrent 
         Interval        =   100
         Left            =   120
         Top             =   420
      End
      Begin VB.TextBox txtHour 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   600
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "0"
         ToolTipText     =   "Hours"
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   960
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "Minutes"
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox txtSec 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Seconds"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblBreak1 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   ":"
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblBreak1 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   ":"
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   300
         Width           =   135
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time 24H"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Time in 24 hours"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblBack 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   570
         TabIndex        =   9
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Time"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GotFocus As Boolean

Private Sub cmdSet_Click()
Dim SetHour As Integer
Dim SetMin As Integer
Dim SetSec As Integer

SetHour = Val(txtHour.Text)
SetMin = Val(txtMin.Text)
SetSec = Val(txtSec.Text)

Call SetNewTime(SetHour, SetMin, SetSec)
End Sub

Private Sub Form_Activate()
txtHour.SelStart = 0
txtHour.SelLength = Len(txtHour.Text)
End Sub

Private Sub Form_Load()
'display the current time in the text boxes
txtHour.Text = Hour(Time)
txtMin.Text = Minute(Time)
txtSec.Text = Second(Time)
GotFocus = False

lblCurrent.Width = frmSetTime.ScaleWidth
End Sub

Private Sub timCurrent_Timer()
'display the current time
Static CurrTime As String
Static OldSecond As Integer

If OldSecond <> Val(Second(Time)) Then
    If Not GotFocus Then
        'display the current time in the text boxes
        txtHour.Text = Hour(Time)
        txtMin.Text = Minute(Time)
        txtSec.Text = Second(Time)
        GotFocus = False
    End If
    
    'display time
    OldSecond = Second(Time)
    CurrTime = Time & "     " & Format(Date, "Long date") '"Current Time: " &
    lblCurrent.Caption = CurrTime
End If
End Sub

Private Sub txtHour_Change()
GotFocus = True
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

EnterNum = KeyAscii
End Function

Private Sub txtMin_Change()
GotFocus = True
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
'data validation for minutes
KeyAscii = EnterNum(txtMin, KeyAscii, "Min")
End Sub

Private Sub txtSec_Change()
GotFocus = True
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
'data validation for seconds
KeyAscii = EnterNum(txtSec, KeyAscii, "Min")
End Sub
