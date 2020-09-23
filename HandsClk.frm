VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmHandsClk 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock"
   ClientHeight    =   3000
   ClientLeft      =   10050
   ClientTop       =   5535
   ClientWidth     =   1920
   Icon            =   "HandsClk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picTest 
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer timDetectDrag 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   960
   End
   Begin VB.Timer timSnapWindow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   960
   End
   Begin SysInfoLib.SysInfo SysInfoClock 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picHook 
      AutoSize        =   -1  'True
      Height          =   535
      Left            =   1320
      Picture         =   "HandsClk.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   535
   End
   Begin MSComDlg.CommonDialog cmndlgClock 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timHand 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timDigital 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   2400
   End
   Begin VB.Image imgLogo 
      Height          =   255
      Left            =   840
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line lnSecond 
      BorderColor     =   &H008080FF&
      X1              =   720
      X2              =   360
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line lnHour 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   960
      X2              =   1320
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line lnMinute 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   960
      X2              =   840
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Label lblShowDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblShowDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblShowTime 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblShowHands 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileAna 
         Caption         =   "&Analogue"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileHour 
         Caption         =   "24 &Hour"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileAboutbreak 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileTimebreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileIdle 
         Caption         =   "&Idle Shut Down"
      End
      Begin VB.Menu mnuFileTim 
         Caption         =   "&Timed Shut Down"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&Shut Down Options..."
      End
      Begin VB.Menu mnuFileColorBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSnap 
         Caption         =   "S&nap Window"
      End
      Begin VB.Menu mnuFileBackOn 
         Caption         =   "&Background On/Off"
      End
      Begin VB.Menu mnuFileBackOpt 
         Caption         =   "Back&ground Options..."
      End
      Begin VB.Menu mnuFileScheme 
         Caption         =   "C&olour Schemes"
      End
      Begin VB.Menu mnuFileBackBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileColor 
         Caption         =   "Set &Colours..."
         Visible         =   0   'False
         Begin VB.Menu mnuFileColorHour 
            Caption         =   "&Hour Hand"
         End
         Begin VB.Menu mnuFileColorMin 
            Caption         =   "&Minute Hand"
         End
         Begin VB.Menu mnuFileColorSec 
            Caption         =   "&Second Hand"
         End
         Begin VB.Menu mnuFileColorBreakAnaBack 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileColorDot 
            Caption         =   "Minute &Dots"
         End
         Begin VB.Menu mnuFileColorAnaBack 
            Caption         =   "A&nalogue Background"
         End
         Begin VB.Menu mnuFileColorBreakDigTime 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileColorTimeFont 
            Caption         =   "&Time Font"
         End
         Begin VB.Menu mnuFileColorTimeBack 
            Caption         =   "T&ime Background"
         End
         Begin VB.Menu mnuFileColorBreakTime 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileColorDayFont 
            Caption         =   "&Day Font"
         End
         Begin VB.Menu mnuFileColorDayback 
            Caption         =   "Da&y Background"
         End
         Begin VB.Menu mnuFileColorBreakDay 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileColorDateFont 
            Caption         =   "Dat&e Font"
         End
         Begin VB.Menu mnuFileColorDateBack 
            Caption         =   "Date &Background"
         End
      End
      Begin VB.Menu mnuFilePass 
         Caption         =   "&Password Options"
         Begin VB.Menu mnuFilePassOpt 
            Caption         =   "Enter/Change &Password"
         End
         Begin VB.Menu mnuFilePassOn 
            Caption         =   "Password &Enabled"
         End
         Begin VB.Menu mnuFilePassLok 
            Caption         =   "Loc&k Menu"
         End
      End
      Begin VB.Menu mnuFileAdv 
         Caption         =   "A&dvanced Options"
         Begin VB.Menu mnuFileAdvOnTop 
            Caption         =   "&Always On Top"
         End
         Begin VB.Menu mnuFileAdvPrev 
            Caption         =   "Pre&vent Shut Down"
         End
         Begin VB.Menu mnuFileAdvStartup 
            Caption         =   "R&un At Startup"
         End
         Begin VB.Menu mnuFileAdvStartMin 
            Caption         =   "Start Minimi&zed"
         End
         Begin VB.Menu mnuFileAdvBreak 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdvSysTime 
            Caption         =   "&Change System Time..."
         End
         Begin VB.Menu mnuFileLoad 
            Caption         =   "Re-&Load System Tray Icon"
         End
         Begin VB.Menu mnuFileAdvStartBreak 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdvShut 
            Caption         =   "&Shut Down Computer"
         End
         Begin VB.Menu mnuFileAdvRestart 
            Caption         =   "&Re-Start Computer"
         End
         Begin VB.Menu mnuFileAdvPower 
            Caption         =   "&Power Down Computer"
         End
         Begin VB.Menu mnuFileAdvForce 
            Caption         =   "&Force Close Windows"
         End
         Begin VB.Menu mnuFileAdvLog 
            Caption         =   "Log-&Off User"
         End
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "A&bout"
      End
      Begin VB.Menu mnuFileReloadBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysShow 
         Caption         =   "&Show Clock"
      End
      Begin VB.Menu mnuSysBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysQuit 
         Caption         =   "&Quit Program"
      End
   End
End
Attribute VB_Name = "frmHandsClk"
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

'details about shut down for particular day.
Private Type ShutDown
    CloseHour As Integer
    CloseMin As Integer
    CloseSec As Integer
    DelayTime As Integer
    DelayOn As Boolean
    ShutWin As Boolean
End Type

'more shut down details. Throw-back from previous version
Dim CloseHour As Integer
Dim CloseMin As Integer
Dim CloseSec As Integer

Dim Delay As Integer

Dim ShutWin As Boolean

'the clock hand variables
Dim SecondAngle As Integer
Dim LastSecond As Integer
Dim LastMinute As Integer
Dim LastHour As Integer
Dim MinuteAngle As Integer
Dim HourAngle As Integer

'digital and analogue clock variables. (Display)
Dim ProperHour As Integer
Dim ProperTime As String
Dim TFHour As Boolean   'TF = twenty four

'general variables.
Dim Counter As Integer
Dim Saved As Boolean

'Colour variables
Dim ColHour As Long
Dim ColMin As Long
Dim ColSec As Long
Dim ColDot As Long
Dim CAnBak As Long
Dim CTmFon As Long   'time font
Dim CTmBak As Long   'time back ground
Dim CDyFon As Long   'day font
Dim CDyBak As Long   'day ...
Dim CDtFon As Long   'date ..
Dim CDtBak As Long   'date ..

'SysTray Icon stuff
'----------------------------------------
Private Type NotifyIconData
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    UCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const WM_MOUSEMOVE = &H200
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NotifyIconData) As Boolean

Dim SysTrayDetails As NotifyIconData
'icon stuff finished
'------------------------------------------

Private Sub Form_Activate()
'This procedure is loaded up when focus is given to
'the window.

'activate the appropiate menus
Call SetMenus

Saved = False

If StillLoading Then
    'set the area where the time will be displayed
    Call SetTimeDimensions
        
    'only do this during program startup
    If lblShowHands.Visible Then
        timHand.Enabled = True
    End If
    
    Call LoadPictureOntoForm(frmHandsClk)
    
    timDigital.Enabled = True
    
    StillLoading = False
    
    'predict the idle shutdown time if active
    If IdleShut Then
        'predict time
        PSTime = PredictIdle(IdleTimeInSec)
    End If
End If

'if the shut down screen is active, then dont go
'through procedure
If frmShut.Visible Then
    Exit Sub
Else
    If frmHandsClk.Visible Then
        'draw the dots if you can see the clock
        Call DrawDots
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Keyboard activation of popup menu

'if space pressed, then show menu
If KeyAscii = 32 Then
    Me.PopupMenu mnuFile, mnuFileExit
    
    'clear text box
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
'set some values used just within this form

'set this flag to let other procedures know what to do
StillLoading = True

'Create Seconds, minutes, hours And angles
LastSecond = Second(Time) - 1
LastMinute = Minute(Time) - 1
LastHour = Hour(Time) - 1
SecondAngle = (Second(Time) * 6) - 90
MinuteAngle = (Minute(Time) * 6) - 90


'get a 12 hour time value
If Hour(Time) > 12 Then
    ProperHour = Hour(Time) - 12
Else
    ProperHour = Hour(Time)
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a popup menu if right click
Call PopMenu(Button)
End Sub

Private Sub Form_Paint()
'redraw the appropiate parts of the screen when necessary
Call DrawDots
Call ShowDigitalValues(True)
End Sub

Private Sub Form_Resize()
Static LastWindowState As Integer
Static Loaded As Boolean

'set clock picture and position if the clock has been minmized or
'restored.
If Not Loaded Then
    'only do this once
    LastWindowState = frmHandsClk.WindowState
    Loaded = True
Else
    If (frmHandsClk.WindowState <> LastWindowState) Then
        LastWindowState = frmHandsClk.WindowState
        
        If frmHandsClk.WindowState = vbNormal Then
            'load clock picture (if applicable)
            Call LoadPictureOntoForm(frmHandsClk)
            Call MoveClock
        End If
    End If
End If
End Sub

Private Sub Form_Terminate()
'save check
Call Form_Unload(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveCheck
If ShutWin Then
    'cancel unload
    Cancel = 1
    frmHandsClk.Visible = False
Else
    'get rid of sys-tray icon
    Call UnloadIcon
    End
End If
End Sub

Private Sub lblShowDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a popup menu if right click
Call PopMenu(Button)
End Sub

Private Sub lblShowDay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a popup menu if right click
Call PopMenu(Button)
End Sub

Private Sub timHand_Timer()
'This Procedure moves the hands of the clock
DoEvents

If (Second(Time) <> LastSecond) Then
    
    'second hand angle = to the second * 6 degrees (360/60) - 90 degrees so that 12 o' clock is parrlell to the sides of the window
    SecondAngle = (Second(Time) * 6) - 90
    '950 is the starting points of the line
    'cos(secondangle)*3.14 /180 = to a point on the circle
    '700 is the distance from the centre point
    lnSecond.X1 = Centre + (Cos(SecondAngle * 3.14 / 180) * 300)
    lnSecond.Y1 = Centre + (Sin(SecondAngle * 3.14 / 180) * 300)
    lnSecond.X2 = Centre + (Cos(SecondAngle * 3.14 / 180) * 860)
    lnSecond.Y2 = Centre + (Sin(SecondAngle * 3.14 / 180) * 860)
    LastSecond = Second(Time)
    
    'redraw the dot the second hand is pointing to
    Call DrawDots((LastSecond + 45)) ' Mod 360)
    
    'change minute hand
    If LastMinute <> Minute(Time) Then
        MinuteAngle = (Minute(Time) * 6) - 90
        lnMinute.X1 = Centre - (Cos(MinuteAngle * 3.14 / 180) * 50)
        lnMinute.Y1 = Centre - (Sin(MinuteAngle * 3.14 / 180) * 50)
        lnMinute.X2 = Centre + (Cos(MinuteAngle * 3.14 / 180) * 800)
        lnMinute.Y2 = Centre + (Sin(MinuteAngle * 3.14 / 180) * 800)
        LastMinute = Minute(Time)
    
        'change hour
        If Hour(Time) > 12 Then
            ProperHour = Hour(Time) - 12
        Else
            ProperHour = Hour(Time)
        End If
        HourAngle = ((ProperHour * 30) - 90) + (Minute(Time) / 2)
        lnHour.X1 = Centre - (Cos(HourAngle * 3.14 / 180) * 50)
        lnHour.Y1 = Centre - (Sin(HourAngle * 3.14 / 180) * 50)
        lnHour.X2 = Centre + (Cos(HourAngle * 3.14 / 180) * 580)
        lnHour.Y2 = Centre + (Sin(HourAngle * 3.14 / 180) * 580)
        LastHour = Hour(Time)
        
    End If
End If
    
End Sub

Private Sub lblShowHands_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a popup menu if right click
Call PopMenu(Button)
End Sub

Private Sub timDigital_Timer()
Call ShowDigitalValues
End Sub

Private Sub lblShowTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a popup menu if right click
Call PopMenu(Button)
End Sub

Private Sub mnuFileAbout_Click()
'show program details (see frmAbout)
DoEvents
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuFileAdvForce_Click()
Call WINForceClose
End
End Sub

Private Sub mnuFileAdvLog_Click()
Call WINLogUserOff
End
End Sub

Private Sub mnuFileAdvOnTop_Click()
'put the form on top or not on top

IsOnTop = Not IsOnTop
PutOnTop
End Sub

Private Sub mnuFileAdvPower_Click()
Call WINPowerDown
End
End Sub

Private Sub mnuFileAdvRestart_Click()
Call WINReboot
End
End Sub

Private Sub mnuFileAdvShut_Click()
Call WINShutdown
End
End Sub

Private Sub mnuFileAdvStartMin_Click()
'change to minimized or not
StartMin = Not StartMin

If StartMin Then
    frmHandsClk.WindowState = 1
Else
    frmHandsClk.WindowState = 0
End If

mnuFileAdvStartMin.Checked = StartMin
Call SaveStatus
End Sub

Private Sub mnuFileAdvStartup_Click()
'put/remove program from startup

StartUp = Not StartUp
mnuFileAdvStartup.Checked = StartUp

Call PutMeInStartup

Call SaveStatus
End Sub

Private Sub mnuFileAdvSysTime_Click()
'call up a form to change the systems' time.
Load frmSetTime
frmSetTime.Show
End Sub

Private Sub mnuFileAna_Click()
'show/hide analogue clock
AnaOn = Not AnaOn
mnuFileAna.Checked = AnaOn
Call HideShow
Call SaveStatus
End Sub

Private Sub mnuFileBackOn_Click()
'turn logo on or off
BackOnOff = Not BackOnOff
mnuFileBackOn.Checked = BackOnOff

Call ShowLogo
Call SaveStatus
End Sub

Private Sub mnuFileBackOpt_Click()
'get picture/logo options
DoEvents
Load frmBack
DoEvents
frmBack.Show
End Sub

Private Sub mnuFileColorAnaBack_Click()
'change colout of the analogue background
Call GetColour("CAnBak")
End Sub

Private Sub mnuFileColorDateBack_Click()
'change colour of date background
Call GetColour("CDtBak")
End Sub

Private Sub mnuFileColorDateFont_Click()
'change colour of date font
Call GetColour("CDtFon")
End Sub

Private Sub mnuFileColorDayback_Click()
'change colour of day background
Call GetColour("CDyBak")
End Sub

Private Sub mnuFileColorDayFont_Click()
'change colour of day font
Call GetColour("CDyFon")
End Sub

Private Sub mnuFileColorDot_Click()
'change colour of minute dots
Call GetColour("ColDot")
End Sub

Private Sub mnuFileColorHour_Click()
'change the colour of the hour hand
Call GetColour("ColHor")
End Sub

Private Sub mnuFileColorMin_Click()
'change the colout of minute hand
Call GetColour("ColMin")
End Sub

Private Sub mnuFileColorSec_Click()
'change the colout of second hand
Call GetColour("ColSec")
End Sub

Private Sub mnuFileColorTimeBack_Click()
'change the colour of the time background
Call GetColour("CTmBak")
End Sub

Private Sub mnuFileColorTimeFont_Click()
'change the colour of time font
Call GetColour("CTmFon")
End Sub

Private Sub mnuFileExit_Click()
'save the current status of the shut-down options
Call SaveStatus

'if the shut-down option is on then...
If ShutWin Then
    'hide clock form (but keep active)
    frmHandsClk.Visible = False
    'disable the display timers
    timDigital.Enabled = False
    timHand.Enabled = False
    
    'minimize to systray
    Call TitleToTray(frmHandsClk)
Else
    'else if "no" then shut the program down
    Call Form_QueryUnload(0, 0)
    End
End If
End Sub

Private Sub mnuFileHour_Click()
'24H option on/off
mnuFileHour.Checked = Not mnuFileHour.Checked

'save changes
Call SaveStatus
End Sub

Public Sub PopMenu(Button As Integer)
'if right-click then display menu
If Button = 2 Then
    Me.PopupMenu mnuFile, mnuFileExit
End If
End Sub

Public Sub CheckStatus(Optional Flag As Integer)
'This procedure loads the details from the .ini
'file and dumps them into variables. If no file
'exists then set defaults and create file.

'===================================================
'Note, .exe buggy, compilation perfect. Cause
'unknown. .ini creation error during startup.
'run-time error 5 - "invalid procedure call or
'argument" is caused by .ini file missing during
'program load.
'1/10/2000
'  ----------------------------------------------
'solution : procedure call "GetAttr()" caused the
'run-time error. I trapped the error before calling
'the "GetAttr()" function in the Form_Load procedure.
'18/10/2000
'--------
'Note :  a possible cause for this could have been the compile
'options for vb. Taking out some of the internal program checks
'vb includes normally in the exe's can cause some unexpected
'errors that are hord to track down.
'27/11/2001
'===================================================

Dim Check As String
Dim ErrorNum As Variant
Dim Day As Integer
Dim FileNum As Integer
Dim TempNum As Integer
Dim TempIdle As String
Dim test As Boolean

'reset error handling
On Error Resume Next

'if file is already being accessed, then pause until
'operation is finished
If Loading Or Saving Or Searching Then
    Exit Sub
End If

'set flag to let other procedures know not to change
'the .ini file.
Loading = True

'error number 53 is "File Not Found"
FileNum = FreeFile
Open FilePath For Input As FileNum
ErrorNum = Err

' ----- No longer used from v6.4
'check to see if 'daynum' is in file.
'(if updating from previous version)
'Day = 0
'
'this will update from previous versions of .ini file
'(versions 3 or below)
'error number 53 is "File Not Found"
'if "no error" then
'If ErrorNum = 0 Then
'    TempNum = FreeFile
'    Open FilePath For Input As TempNum
'        While Not EOF(TempNum)
'            Line Input #TempNum, Check
'            If LCase(GetBefore(Check)) = "daynumber" Then
'                Day = Day + 1
'            End If
'        Wend
'    Close TempNum
'Else
'    Close FileNum
'End If
'
''if day is more than zero, then .ini version is
''current
'If Day > 0 Then
'    Day = 1
'Else
'    ErrorNum = 53 'set default settings (53 = "File Not Found")
'End If
'-----

'check if file was found
If ErrorNum = 0 Then
    
    While Not EOF(FileNum)
        Line Input #FileNum, Check
        
        Select Case LCase(GetBefore(Check))
        'general settings
        Case "appowner"
             Owner = GetAfter(Check)
        
        Case "runatstartup"
            StartUp = GetAfter(Check)
            mnuFileAdvStartup.Checked = StartUp
            
            'remove or add the registry key to start up the
            'program.
            Call PutMeInStartup
            
        Case "startminimized"
            StartMin = GetAfter(Check)
            mnuFileAdvStartMin.Checked = StartMin
            
            If StartMin Then
                'minimize program
                frmHandsClk.WindowState = 1
            End If
            
        Case "shutdownmethod"
            Method = GetAfter(Check)
        
        Case "preventshutdown"
            'whether or not to stop other apps from
            'closing windows.
            PreventShut = GetAfter(Check)
            mnuFileAdvPrev.Checked = PreventShut
        
        Case "idleshutdownon"
            'whether or not the computer should shut down the
            'computer after a specified time
            IdleShut = GetAfter(Check)
            mnuFileIdle.Checked = IdleShut
            If IdleShut Then
                'set the tracking time
                'call trackidletime
            End If
            
        Case "idletimeinsec"
            'the amount of time the program should wait before
            'shutting the computer down (in seconds)
            IdleTimeInSec = Val(GetAfter(Check))
            IdleTimeInSec = InvalidIdleTime(IdleTimeInSec)
            
        Case "alwaysontop"
            IsOnTop = GetAfter(Check)
            PutOnTop
        
        Case "analogue"
            'is the analogue clock on or off
            If LCase(GetAfter(Check)) = "no" Then
                AnaOn = False
            Else
                AnaOn = True
            End If
            mnuFileAna.Checked = AnaOn
    
        Case "24hour"
            If LCase(GetAfter(Check)) = "no" Then
                mnuFileHour.Checked = False
            Else
                mnuFileHour.Checked = True
            End If
            
            'display the time
            ProperHour = Hour(Time) Mod 12
            If ProperHour = 0 Then
                ProperHour = 12
            End If

            'If Not mnuFileHour.Checked Then
                'lblShowTime.Caption = Format(ProperHour, "0") & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
            'End If

        'clock positioning
        Case "snapwindow"
            SnapOn = GetAfter(Check)
            mnuFileSnap.Checked = SnapOn
            Call SetSnap
            
        Case "lastposx"
            LastPos.X = Val(GetAfter(Check))
            
        Case "lastposy"
            LastPos.Y = Val(GetAfter(Check))
        
        'password settings
        Case "password"
            Password = DecryptData(GetAfter(Check))
        
        Case "passwordactive"
            PassActive = GetAfter(Check)
            mnuFilePassOn.Checked = PassActive

        'Daily time settings
        Case "daynumber"
            'get current array element
            Day = Val(GetAfter(Check))
            
        Case "delaytime"
            Week(Day).DelayTime = Val(GetAfter(Check))
        
        Case "closehour"
            Week(Day).CloseHour = Val(GetAfter(Check))
            
        Case "closeminute"
            Week(Day).CloseMin = Val(GetAfter(Check))
        
        Case "closesecond"
            Week(Day).CloseSec = Val(GetAfter(Check))
        
        Case "closewindows"
            Week(Day).ShutWin = GetAfter(Check)
        
        Case "delayon"
            If (LCase(GetAfter(Check)) = "on") Or (LCase(GetAfter(Check)) = "true") Then
                Week(Day).DelayOn = True
            Else
                Week(Day).DelayOn = False
            End If
        
        'colours
        Case "colhour"
            ColHour = Val(GetAfter(Check))
        
        Case "colmin"
            ColMin = Val(GetAfter(Check))
            
        Case "colsec"
            ColSec = Val(GetAfter(Check))
            
        Case "coldots"
            ColDot = Val(GetAfter(Check))
        
        Case "colanaloguebackground"
            CAnBak = Val(GetAfter(Check))
            
        Case "coltimefont"
            CTmFon = Val(GetAfter(Check))
        
        Case "coltimeback"
            CTmBak = Val(GetAfter(Check))
        
        Case "coldayfont"
            CDyFon = Val(GetAfter(Check))
            
        Case "coldayback"
            CDyBak = Val(GetAfter(Check))
            
        Case "coldatefont"
            CDtFon = Val(GetAfter(Check))
        
        Case "coldateback"
            CDtBak = Val(GetAfter(Check))
        
        'background details
        Case "backlogo"
            BackPath = GetAfter(Check)
            
            If Flag <> DontLoadPic Then
                'only go through this if the flag is not set
                
                If BackPath = "0" Then
                    BackPath = ""
                    'clear picture/logo
                    imgLogo.Picture = LoadPicture
                Else
                    'if file exists
                    If Dir(BackPath) <> "" Then
                        imgLogo.Picture = LoadPicture(BackPath)
                    Else
                        'no picture found
                        BackPath = ""
                        imgLogo.Picture = LoadPicture
                        mnuFileBackOn.Checked = False
                    End If
                End If
            End If
            
        Case "backtile"
            'stretch/tile/centre
            StretchTile = GetAfter(Check)
            
        Case "backonoff"
            BackOnOff = GetAfter(Check)
            'activate background if a picture exists
            If (BackOnOff) And (BackPath <> "") Then
                mnuFileBackOn.Checked = True
            Else
                mnuFileBackOn.Checked = False
                BackOnOff = False
            End If

        End Select
        
    Wend

Else
    'file was not found, set default values.
    
    Call SetDefaults
    Call SaveStatus
End If

Close FileNum

Loading = False

'start the idle timer
TempIdle = GetTimeIdle

'get today's values
Day = Weekday(Date)

'if colour values and delay time are zero, then
'assume program saved nulled values during the last
'unloading of the program, so set and use the
'default values. This partly conseals the .exe
'unloading bug along with this codes' cousin in
'the procedure "SaveStatus".
If (Week(Day).DelayTime = 15) And (ColHour = 0) And (ColMin = 0) And (ColSec = 0) And (ColDot = 0) And (CTmFon = 0) And (CTmBak = 0) And (CDyFon = 0) And (CDyBak = 0) And (CDtFon = 0) And (CDtBak = 0) Then
    Call SetDefaults
    
    'save the defaults
    Call SaveStatus
End If

If Week(Day).DelayOn Then
    frmOptions.Tag = "On"
Else
    frmOptions.Tag = "Off"
End If

'set todays shutdown values in array for storage
CloseHour = Week(Day).CloseHour
CloseMin = Week(Day).CloseMin
CloseSec = Week(Day).CloseSec
ShutWin = Week(Day).ShutWin
mnuFileTim.Checked = ShutWin
Delay = Week(Day).DelayTime

'if file not found and default setting hav been set
'then save the default settings
If ErrorNum <> 0 Then
    Call SaveStatus
End If

'do not run the following if the DontLoadPic flag is set
If Flag <> DontLoadPic Then
    Call HideShow
    Call SetMenus
    
    'reduce noticeable flicker by putting
    'the background into the forms' picture
    'property.
    'Call LoadPictureOntoForm(frmhandsclk)
    
    Call ShowLogo
End If

'resume normal error handling
On Error GoTo 0
End Sub

Private Sub mnuFileIdle_Click()
'change whether or not to shut down the computer after a certain
'peroid of idleness.

DoEvents
IdleShut = Not IdleShut

mnuFileIdle.Checked = IdleShut

If IdleShut Then
    'call trackidletime
Else
    'call endtracking
End If

Call SaveStatus
End Sub

Private Sub mnuFileLoad_Click()
're-load the systray icon in case of unexpected
'events.
Call UnloadIcon
Call LoadIcon
End Sub

Private Sub mnuFileOpt_Click()
'set shut-down options
Load frmOptions
frmOptions.Show

'get changed values
'Call GetValues(Delay, CloseHour, CloseMin, CloseSec, ShutWin, ColHour, ColMin, ColSec, ColDot, CAnBak, CTmFon, CTmBak, CDyFon, CDyBak, CDtFon, CDtBak)
End Sub

Private Sub mnuFilePassLok_Click()
'lock the menu
If PassActive Then
    CorrectPass = False
    Call SetMenus
End If
End Sub

Private Sub mnuFilePassOn_Click()
PassActive = Not PassActive
mnuFilePassOn.Checked = PassActive
Call SetMenus
Call SaveStatus
End Sub

Private Sub mnuFilePassOpt_Click()
If CorrectPass Then
    AskOrChange = Change
Else
    AskOrChange = Ask
End If

Load frmPass
frmPass.Show
End Sub

Private Sub mnuFileAdvPrev_Click()
'Turn off/on the option that allows other applications to
'shut down windows.

PreventShut = Not PreventShut
mnuFileAdvPrev.Checked = PreventShut
End Sub

Private Sub mnuFileScheme_Click()
'This loads the screen for colour schemes
Load frmScheme
frmScheme.Show
End Sub

Private Sub mnuFileSnap_Click()
SnapOn = Not SnapOn
Call SetSnap
Call SaveStatus
End Sub

Private Sub mnuFileTim_Click()
'timed shut-down on/off

ShutWin = Not ShutWin
mnuFileTim.Checked = ShutWin
Week(Weekday(Date)).ShutWin = ShutWin

'save change
SaveStatus

If Not ShutWin Then
    Unload frmShut
Else
    Load frmShut
End If
End Sub

Public Sub SaveStatus()
'saves the current values and settings. Obviously :P
'Please note : this procedure is different from the
'SaveStatus procedure in the form frmHandsClk.

Dim Day As Integer

Dim AnaStatus As String
Dim HourStatus As String
Dim DelayTime As String
Dim CloseH As String
Dim CloseM As String
Dim CloseS As String
Dim CloseWin As String
Dim DelayOn As String
Dim ErrFileNum As Integer
Dim FileNum As Integer
Dim CurrentOwner As String

'determine .ini path from .exe path
FilePath = AddFile(App.Path, FileName)

'if file is already being accessed, then pause until
'operation is finished
If (Loading Or Saving) Or (Not CanAccessFile(FilePath, FileOutPut)) Then
    Exit Sub
End If

'set flag to let other procedures know not to change
'the .ini file.
Saving = True

Day = Weekday(Date)

'user the current value of "owner" to save.
CurrentOwner = Owner

'if colour values and delay time are zero, then
'assume program is unloading and don't save the
'reset values.
If (Week(Day).DelayTime = 15) And (ColHour = 0) And (ColMin = 0) And (ColSec = 0) And (ColDot = 0) And (CTmFon = 0) And (CTmBak = 0) And (CDyFon = 0) And (CDyBak = 0) And (CDtFon = 0) And (CDtBak = 0) Then
    Exit Sub
End If

Saved = True

'set strings for saving
AnaStatus = "Analogue="
HourStatus = "24Hour="

'analogue on/off ?
If AnaOn Then
    AnaStatus = AnaStatus + "yes"
Else
    AnaStatus = AnaStatus + "no"
End If

'24H on/off ?
If frmHandsClk.mnuFileHour.Checked Then
    HourStatus = HourStatus + "yes"
Else
    HourStatus = HourStatus + "no"
End If


'Update as many variables as possible before
'saving (this is basically double checking the
'values before saving and stops data corruption)
StartUp = frmHandsClk.mnuFileAdvStartup.Checked
PassActive = frmHandsClk.mnuFilePassOn.Checked
BackOnOff = frmHandsClk.mnuFileBackOn.Checked

If frmHandsClk.WindowState = vbMinimized Then
    StartMin = True
Else
    StartMin = False
    LastPos.X = frmHandsClk.Left
    LastPos.Y = frmHandsClk.Top
End If
mnuFileAdvStartMin.Checked = StartMin

'update the colours before saving
ColHour = lnHour.BorderColor
ColMin = lnMinute.BorderColor
ColSec = lnSecond.BorderColor
ColDot = frmHandsClk.ForeColor
CAnBak = lblShowHands.BackColor
CTmFon = lblShowTime.ForeColor
CTmBak = lblShowTime.BackColor
CDyFon = lblShowDay.ForeColor
CDyBak = lblShowDay.BackColor
CDtFon = lblShowDate.ForeColor
CDtBak = lblShowDate.BackColor

'get an available file number and save values.
FileNum = FreeFile
Open FilePath For Output As #FileNum
    Print #FileNum, "[COMPAPP CLOCK VERSION " & App.Major & "."; App.Minor & "." & App.Revision & "]"
    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[CLOCK VALUES]"
    Print #FileNum, "AppOwner=" & CurrentOwner
    Print #FileNum, "RunAtStartUp=" & StartUp
    Print #FileNum, "ShutDownMethod=" & Method
    Print #FileNum, "PreventShutdown=" & PreventShut
    Print #FileNum, "IdleShutdownOn=" & IdleShut
    Print #FileNum, "IdleTimeInSec=" & IdleTimeInSec
    Print #FileNum, "AlwaysOnTop=" & IsOnTop
    Print #FileNum, "StartMinimized=" & StartMin
    Print #FileNum, "SnapWindow=" & SnapOn
    Print #FileNum, ""
    Print #FileNum, "LastPosX=" & LastPos.X
    Print #FileNum, "LastPosY=" & LastPos.Y
    Print #FileNum, ""
    Print #FileNum, AnaStatus
    Print #FileNum, HourStatus
    Print #FileNum, ""
    
    Print #FileNum, "[PASSWORD SETTINGS]"
    Print #FileNum, "Password=" & EncryptData(Password)
    Print #FileNum, "PasswordActive=" & PassActive
    Print #FileNum, ""
    
    Print #FileNum, "[DAY SETTINGS]"
    For Day = 1 To 7
        'save details for each day
        DelayTime = "DelayTime=" & Week(Day).DelayTime
        CloseH = "CloseHour=" & Week(Day).CloseHour
        CloseM = "CloseMinute=" & Week(Day).CloseMin
        CloseS = "CloseSecond=" & Week(Day).CloseSec
        CloseWin = "CloseWindows=" & Week(Day).ShutWin
        If Week(Day).DelayOn Then
            DelayOn = "DelayOn=" & "On"
        Else
            DelayOn = "DelayOn=" & "Off"
        End If
    
        Print #FileNum, "DayNumber=" & Day
        Print #FileNum, DelayTime
        Print #FileNum, CloseH
        Print #FileNum, CloseM
        Print #FileNum, CloseS
        Print #FileNum, CloseWin
        Print #FileNum, DelayOn
        Print #FileNum, ""
    Next Day
    
    Print #FileNum, "[COLOUR SETTINGS]"
    Print #FileNum, "ColHour=" & ColHour
    Print #FileNum, "ColMin=" & ColMin
    Print #FileNum, "ColSec=" & ColSec
    Print #FileNum, "ColDots=" & ColDot
    Print #FileNum, "ColAnalogueBackground=" & CAnBak
    Print #FileNum, "ColTimeFont=" & CTmFon
    Print #FileNum, "ColTimeBack=" & CTmBak
    Print #FileNum, "ColDayFont=" & CDyFon
    Print #FileNum, "ColDayBack=" & CDyBak
    Print #FileNum, "ColDateFont=" & CDtFon
    Print #FileNum, "ColDateBack=" & CDtBak
    
    Print #FileNum, ""
    Print #FileNum, "[BACKGROUND SETTINGS]"
    Print #FileNum, "BackLogo=" & BackPath      'location of the background picture
    Print #FileNum, "BackTile=" & StretchTile   'picture style
    Print #FileNum, "BackOnOff=" & BackOnOff     'background on or off
    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[DEBUG]"
    Print #FileNum, "LastSave=" & Time & " " & Date
    Print #FileNum, "ShutdownAt=" & PSTime
    
Close #FileNum

'predict the idle shutdown time if active
If IdleShut Then
    'predict time
    PSTime = PredictIdle(IdleTimeInSec)
End If

Saving = False
End Sub

Public Sub SetDefaults()
'This procedure (obviously), sets the default values
'for variables show anything be amiss.

'used to cycle throught the days to set the defaults
Dim DayNum As Integer


For DayNum = 1 To 7
    'the default values are set for all days
    Week(DayNum).CloseHour = 0
    Week(DayNum).CloseMin = 0
    Week(DayNum).CloseSec = 0
    
    'timed shut down on/off
    Week(DayNum).ShutWin = False
    mnuFileTim.Checked = Week(DayNum).ShutWin
    
    'Answer time (in seconds)
    Week(DayNum).DelayTime = 15
    
    'delay on/off (off = waits for answer indefinatly)
    Week(DayNum).DelayOn = False

Next DayNum
    
'analogue time
mnuFileHour.Checked = False
    
'colours
ColHour = &HFF00FF  'light purple
ColMin = &HC000C0   'dark purple
ColSec = &H8080FF   'light red
ColDot = &H80000012 'black
CAnBak = &HFFFF&  'yellow
CTmFon = &H800080 'purple
CTmBak = &HFFFF&  'yellow
CDyFon = &HFFFF&  'yellow
CDyBak = &HC000C0 'purple
CDtFon = &HFFFF&  'yellow
CDtBak = &HC000C0 'purple

'background details
BackPath = AddFile(WindowsDirectory, "Clouds.Bmp")
StretchTile = "Stretch"
BackOnOff = False

'put program in startup
StartUp = True
mnuFileAdvStartup.Checked = StartUp
Call MakeStartUp(AddFile(App.Path, (App.EXEName & ".exe")))

'start minimized
StartMin = False
mnuFileAdvStartMin.Checked = StartMin

'snap window to side of screen
SnapOn = True
mnuFileSnap.Checked = SnapOn

'move the clock to the bottom right of the screen
LastPos.X = Screen.Width
LastPos.Y = Screen.Height
Call MoveClock

'no password, option off
Password = ""
PassActive = False

'get the registered owner
Owner = GetOwnerInReg

'the shut down method
Method = "Shut Down"

'allow other apps to close windows
PreventShut = False

'allow clock to be hidden
IsOnTop = False
PutOnTop    'sub. will take value from "IsOnTop"

'do not close windows after period of
'inactivity
IdleShut = False
IdleTimeInSec = InvalidIdleTime(0)

End Sub

Private Sub mnuSysQuit_Click()
'remove systray icon and exit program
Call SaveStatus
Call UnloadIcon
End
End Sub

Private Sub mnuSysShow_Click()
'menu for system tray. "Show Clock"

If Not frmHandsClk.Visible Then
    'restore from systray
    Call TrayToTitle(frmHandsClk)
End If

'activate the clock if inactive
If Not StillLoading Then
    frmHandsClk.Show
    
    timDigital.Enabled = True
    If frmHandsClk.mnuFileAna.Checked Then
        timHand.Enabled = True
    End If
    timDetectDrag.Enabled = True
End If
End Sub

Private Sub SysInfoClock_DisplayChanged()
'this will re-position the clock to the bottom
'right hand side of the screen after the screen
'resolution is changed.
'Call Form_Activate
Call MoveClock

'set the area where the time will be displayed (in pixels)
LabelRect.Left = lblShowTime.Left / Screen.TwipsPerPixelX
LabelRect.Top = lblShowTime.Top / Screen.TwipsPerPixelY
LabelRect.Bottom = (lblShowTime.Top + lblShowTime.Height) / Screen.TwipsPerPixelY
LabelRect.Right = (lblShowTime.Left + lblShowTime.Width) / Screen.TwipsPerPixelX

End Sub

Private Sub SysInfoClock_TimeChanged()
'reset the systray icon so that it is always
'displayed
Call UnloadIcon
Call LoadIcon
End Sub

Public Sub DrawDots(Optional ByVal SecPoint As Integer = -1)
'Put dots onto the screen
'[optional] draw a single dot at the point given in seconds

Dim XCo As Integer
Dim YCo As Integer

'only draw the dots if the appropiate option is selected
If Not frmHandsClk.mnuFileAna.Checked Then
    Exit Sub
End If

'a specific point was passed. Only draw this dot
If SecPoint >= 0 Then
    'validate the parameter
    SecPoint = SecPoint Mod 60
    
    XCo = ((CentreDot + ((50 + Cos((SecPoint * 6) * 3.14 / 180) * 900))) / Screen.TwipsPerPixelX)
    YCo = ((CentreDot + ((50 + Sin((SecPoint * 6) * 3.14 / 180) * 900))) / Screen.TwipsPerPixelY)
    
    'if the point passed points to a large dot, draw a large dot,
    'otherwise, draw a small one
    If (SecPoint Mod 5) = 0 Then
        Call DrawRect(frmHandsClk.hDc, ColDot, XCo, YCo, XCo + 2, YCo + 2)
    Else
        Call DrawRect(frmHandsClk.hDc, ColDot, XCo, YCo, XCo + 1, YCo + 1)
    End If
    
    'don't draw any other dots
    Exit Sub
End If

'draw all the dots onto the form

'draw small dots
For Counter = 0 To 360 Step 6
    XCo = ((CentreDot + ((50 + Cos(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelX)
    YCo = ((CentreDot + ((50 + Sin(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelY)
    Call DrawRect(frmHandsClk.hDc, ColDot, XCo, YCo, XCo + 1, YCo + 1)
Next Counter
    
'draw big dots
For Counter = 0 To 360 Step 30
    XCo = ((CentreDot + ((50 + Cos(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelX)
    YCo = ((CentreDot + ((50 + Sin(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelY)
    Call DrawRect(frmHandsClk.hDc, ColDot, XCo, YCo, XCo + 2, YCo + 2)
Next Counter

'Call PutDotsOnForm(ColDot)
End Sub

Private Sub timDetectDrag_Timer()
Static FormTop As Integer
Static FormLeft As Integer

DoEvents

If Not SnapOn Then
    'if the Snap Window option is not on, then slow down this
    'timer
    DoEvents
    If timDetectDrag.Interval < 1000 Then
        timDetectDrag.Interval = 1000
    End If
    timSnapWindow.Enabled = False
    Exit Sub
Else
    'if the Snap Window option is turned on, then speed up timer
    'event
    If timDetectDrag.Interval > 1 Then
        timDetectDrag.Interval = 1
    End If
End If

DoEvents

'if just starting then, set the variables
If (FormTop = 0) And (FormLeft = 0) Then
    FormTop = frmHandsClk.Top
    FormLeft = frmHandsClk.Left
    'Exit Sub
End If

If (FormTop <> frmHandsClk.Top) Or (FormLeft <> frmHandsClk.Left) Then
    timSnapWindow.Enabled = True
Else
    timSnapWindow.Enabled = False
End If

DoEvents

Call CheckIfOutsideScreen(frmHandsClk)
End Sub

Private Sub timSnapWindow_Timer()
Const WithinDist = 10   'within a distance of 10 pixels

DoEvents

If frmHandsClk.WindowState <> vbNormal Then
    'an error will occur if the window is not in "normal" mode.
    Exit Sub
End If

LastPos.X = frmHandsClk.Left
LastPos.Y = frmHandsClk.Top

DoEvents

Call SnapWindow(frmHandsClk, WithinDist)
End Sub

Private Sub GetColour(ColObj As String)
'sets the colour for an item and saves the change

Select Case LCase(ColObj)
Case "colhor"
    ColHour = ColWin(ColHour)
    lnHour.BorderColor = Val(ColHour)
Case "colmin"
    ColMin = ColWin(ColMin)
    lnMinute.BorderColor = Val(ColMin)
Case "colsec"
    ColSec = ColWin(ColSec)
    lnSecond.BorderColor = Val(ColSec)
Case "coldot"
    ColDot = ColWin(ColDot)
    frmHandsClk.ForeColor = Val(ColDot)
Case "canbak"
    CAnBak = ColWin(CAnBak)
    lblShowHands.BackColor = CAnBak 'Val(CAnBak)
Case "ctmfon"
    CTmFon = ColWin(CTmFon)
    lblShowTime.ForeColor = Val(CTmFon)
Case "ctmbak"
    CTmBak = ColWin(CTmBak)
    lblShowTime.BackColor = Val(CTmBak)
Case "cdyfon"
    CDyFon = ColWin(CDyFon)
    lblShowDay.ForeColor = Val(CDyFon)
Case "cdybak"
    CDyBak = ColWin(CDyBak)
    lblShowDay.BackColor = Val(CDyBak)
Case "cdtfon"
    CDtFon = ColWin(CDtFon)
    lblShowDate.ForeColor = Val(CDtFon)
Case "cdtbak"
    CDtBak = ColWin(CDtBak)
    lblShowDate.BackColor = Val(CDtBak)
End Select

'save current settings
SaveStatus
End Sub

Public Sub SetColour()
'sets the colours for each item

lnHour.BorderColor = Val(ColHour)
lnMinute.BorderColor = Val(ColMin)
lnSecond.BorderColor = Val(ColSec)
frmHandsClk.ForeColor = Val(ColDot)
lblShowHands.BackColor = Val(CAnBak) 'Val(CAnBak)
lblShowTime.ForeColor = Val(CTmFon)
lblShowTime.BackColor = Val(CTmBak)
lblShowDay.ForeColor = Val(CDyFon)
lblShowDay.BackColor = Val(CDyBak)
lblShowDate.ForeColor = Val(CDtFon)
lblShowDate.BackColor = Val(CDtBak)
End Sub

Private Function ColWin(Colour As Long) As Long
DoEvents

'get new colour from colour dialogue box
cmndlgClock.Color = Colour
cmndlgClock.ShowColor
ColWin = cmndlgClock.Color
End Function

Private Sub PicHook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this picks up events in the system tray
Static MousePos As Long


MousePos = X / Screen.TwipsPerPixelX

Select Case MousePos
'all options are here in case I want to
'cut/copy/paste
    Case WM_LBUTTONDBLCLK
        'show clock
        AppActivate App.Title
        Call mnuSysShow_Click
        frmHandsClk.Show
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
    Case WM_RBUTTONDBLCLK
        'exit
        'Call mnuSysQuit_Click
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
        'show menu
        PopupMenu mnuSysTray
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'this procedure is called when the prorgam receives
'a request to shut down. The program is not actually
'shutting down yet. For more information, please see
'the Form_QueryUnload in the frmShut screen.


If (UnloadMode = vbAppWindows) Or (UnloadMode = vbAppTaskManager) Then
    'program is to be closed by tasklist or if the
    'current session of windows is finishing, then
    'unload the icon and quit windows
    
    If UnloadMode = vbAppWindows Then
        'if turned on, then stop windows from shutting down
        Cancel = PreventShut
    End If
    
    Call PreUnload
    End
Else
    'disable timers to avoid saving null details
    timDigital.Enabled = False
    timHand.Enabled = False

    'if unloading, remove the system tray icon
    If Not frmHandsClk.mnuFileTim.Checked Then
        Call UnloadIcon
        End
    Else
        Call TitleToTray(Me)
    End If

End If

End Sub

Public Sub PreUnload()
'disable timers to avoid saving null details
timDigital.Enabled = False
timHand.Enabled = False

'save the current settings
'Call SaveStatus

'unload icon
Call UnloadIcon

'remove the background bitmap
Call DeleteBitmap(BmpTime.hDcMemory, BmpTime.hDcBitmap, BmpTime.hDcPointer)
End Sub

Public Sub ShowLogo()
'hides or shows the logo or picture on the background
Dim BStyleVal As Integer

'exit sub if there is no logo
If (BackPath = "") Or (Dir(BackPath) = "") Then
    Exit Sub
End If


'convert boolean values to "1" or "0"
BStyleVal = ((Not BackOnOff) * -1)

lblShowTime.BackStyle = BStyleVal
lblShowHands.BackStyle = BStyleVal
lblShowDay.BackStyle = BStyleVal
lblShowDate.BackStyle = BStyleVal

'picText.Visible = False
If Not BackOnOff Then
    'no background picture
    frmHandsClk.Picture = LoadPicture
    Call GetTimeBackground
Else
    'a background picture
    Call LoadPictureOntoForm(frmHandsClk)
End If
End Sub

Public Sub LoadIcon()
'set systray icon details
SysTrayDetails.cbSize = Len(SysTrayDetails)
SysTrayDetails.hwnd = picHook.hwnd
SysTrayDetails.uId = 1&
SysTrayDetails.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
SysTrayDetails.UCallBackMessage = WM_MOUSEMOVE
SysTrayDetails.hIcon = frmHandsClk.Icon
SysTrayDetails.szTip = Format(Date, "Long Date") & Chr$(0)

'send details
Shell_NotifyIcon NIM_ADD, SysTrayDetails

End Sub

Public Sub UnloadIcon()
'remove the icon from the systray
SysTrayDetails.cbSize = Len(SysTrayDetails)
SysTrayDetails.hwnd = picHook.hwnd
SysTrayDetails.uId = 1&
Shell_NotifyIcon NIM_DELETE, SysTrayDetails
End Sub

Public Sub ShowDigitalValues(Optional ByVal Update As Boolean)
'This displays the time in digital mode along with
'the day and date.

'The "Left" value of the text centred in the picture
'box
Dim TextLeft As Integer
Dim TimeFont As FontStruc
Dim TempBmp As BitmapStruc
Dim Result As Long

Static LastTime As String

DoEvents

'calculate and create a string containing time
If ProperHour = 0 Then
    ProperHour = 12
End If
ProperTime = ProperHour & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00") 'Right(Time, Len(Time) - Len(Trim(Str(ProperHour))))

'display time in title if minimized (check p/s)
If (frmHandsClk.WindowState = 1) And (frmHandsClk.Caption <> Time) Then
    'in case of change in shut-down options
    If mnuFileHour.Checked Then
        frmHandsClk.Caption = Time
    Else
        frmHandsClk.Caption = ProperTime
    End If
Else
    If (frmHandsClk.WindowState = 0) And (frmHandsClk.Caption <> "Clock") Then
        frmHandsClk.Caption = "Clock"
    End If
End If

DoEvents

'display time
If (LastTime <> Time) Or (Update) Then
    'redisplay the time
    lblShowTime.Caption = ""
    lblShowTime.Visible = True
    If mnuFileHour.Checked Then
        LastTime = Time
    Else
        LastTime = ProperTime
    End If
    
    'set the font of the time to be displayed
    TimeFont.Bold = lblShowTime.FontBold
    TimeFont.Italic = lblShowTime.FontItalic
    TimeFont.Name = lblShowTime.FontName
    TimeFont.PointSize = lblShowTime.FontSize
    TimeFont.StrikeThru = lblShowTime.FontStrikethru
    TimeFont.Underline = lblShowTime.FontUnderline
    TimeFont.Colour = lblShowTime.ForeColor
    TimeFont.Alignment = vbCentreAlign
    
    'create a new bitmap
    TempBmp.Area = BmpTime.Area
    Call CreateNewBitmap(TempBmp.hDcMemory, TempBmp.hDcBitmap, TempBmp.hDcPointer, TempBmp.Area, frmHandsClk, lblShowTime.BackColor, InPixels)
    
    'copy background onto the new bitmap
    Result = BitBlt(TempBmp.hDcMemory, 0, 0, (TempBmp.Area.Right - TempBmp.Area.Left), (TempBmp.Area.Bottom - TempBmp.Area.Top), BmpTime.hDcMemory, 0, 0, SRCCOPY)

    'draw the time text
    Call MakeText(TempBmp.hDcMemory, LastTime, 0, 0, (TempBmp.Area.Bottom - TempBmp.Area.Top), (TempBmp.Area.Right - TempBmp.Area.Left), TimeFont, InPixels)
    
    'display the time
    Result = BitBlt(frmHandsClk.hDc, TempBmp.Area.Left, TempBmp.Area.Top, (TempBmp.Area.Right - TempBmp.Area.Left), (TempBmp.Area.Bottom - TempBmp.Area.Top), TempBmp.hDcMemory, 0, 0, SRCCOPY)

    Call DeleteBitmap(TempBmp.hDcMemory, TempBmp.hDcBitmap, TempBmp.hDcPointer)
    
    'remember the current time
    LastTime = Time
End If

DoEvents

'show date and update shutdown times for new day
If Format(Date, "d/m/yyyy") <> lblShowDate.Caption Then
    'get name of day
    lblShowDay.Caption = GetDayName
    lblShowDate.Caption = Format(Date, "d/m/yyyy")
    
    'don't load the picture
    Call CheckStatus(1)

    'check for timed shutdown for today
    frmShut.timClose.Enabled = Week(Weekday(Date)).ShutWin
End If
End Sub

Public Function GetDayName() As String
'This function returns the day of the week as a string

GetDayName = WeekdayName(Weekday(Date, vbMonday))
End Function

Private Sub SetMenus()
'This option will enable or disable certain menu
'options depending on whether or not the password
'is enabled.

If (Not CorrectPass) And (PassActive) Then
    'if the password is active and the correct
    'password has not been entered, then;
    mnuFileAdv.Enabled = False
    mnuFileAna.Enabled = False
    mnuFileBackOn.Enabled = False
    mnuFileBackOpt.Enabled = False
    mnuFileColor.Enabled = False
    mnuFileHour.Enabled = False
    mnuFileLoad.Enabled = False
    mnuFileOpt.Enabled = False
    mnuFileTim.Enabled = False
    mnuFileScheme.Enabled = False
    mnuFile.Enabled = False
    mnuFilePassOn.Enabled = False
    mnuFileSnap.Enabled = False
    mnuFileIdle.Enabled = False
Else
    'if the password is not active OR has been
    'entered correctly, then;
    mnuFileAdv.Enabled = True
    mnuFileAna.Enabled = True
    mnuFileBackOn.Enabled = True
    mnuFileBackOpt.Enabled = True
    mnuFileColor.Enabled = True
    mnuFileHour.Enabled = True
    mnuFileLoad.Enabled = True
    mnuFileOpt.Enabled = True
    mnuFileTim.Enabled = True
    mnuFileScheme.Enabled = True
    mnuFile.Enabled = True
    mnuFilePassOn.Enabled = True
    mnuFileSnap.Enabled = True
    mnuFileIdle.Checked = True
End If

'if the password has been entered correctly and is
'active, then;
If PassActive And CorrectPass Then
    mnuFilePassLok.Enabled = True
Else
    mnuFilePassLok.Enabled = False
End If
End Sub

Private Sub ShowPicture()
'set the background properties of all the labels
'to 'Transparent'
Dim TransOn As Integer

'True = -1
'False = 0

TransOn = ((Not BackOnOff) * -1)
lblShowHands.BackStyle = TransOn
lblShowTime.BackStyle = TransOn
lblShowDay.BackStyle = TransOn
lblShowDate.BackStyle = TransOn
End Sub

Private Sub PutMeInStartup()
'Remove or add program to startup when windows starts.
If StartUp Then
    Call MakeStartUp(AddFile(App.Path, (App.EXEName & ".exe")))
Else
    Call DeleteFromStartup(AddFile(App.Path, (App.EXEName & ".exe")))
End If
End Sub

Private Sub CentreText()
'Call LockWindow(frmHandsClk)
'find the width of the text

'lblShowTime.Visible = False

'picTime.Cls
'picTime.Font = lblShowTime.Font
'picTime.FontSize = lblShowTime.FontSize
'picTime.ForeColor = lblShowTime.ForeColor

'picTime.Visible = False

'picText.Width = picTime.Width
'picText.Height = picTime.Height
'picText.Cls
'picText.Font = lblShowTime.Font
'picText.FontSize = lblShowTime.FontSize
'picText.ForeColor = lblShowTime.ForeColor

'this will set the CurrentX property to get the
'width of the text.
'picText.CurrentX = 0
'picText.Print lblShowTime.Caption;

'centre the text and display
'TextLeft = (picText.ScaleWidth / 2) - (picText.CurrentX / 2)
'picTime.Cls
'picTime.CurrentX = TextLeft
'picTime.Print lblShowTime.Caption

'Call UnLockWindow
End Sub

Private Sub SetSnap()
'This will enable or disable the snap window functions

frmHandsClk.timDetectDrag.Enabled = SnapOn
frmHandsClk.mnuFileSnap.Checked = SnapOn

If SnapOn Then
    Call timSnapWindow_Timer
End If
End Sub
