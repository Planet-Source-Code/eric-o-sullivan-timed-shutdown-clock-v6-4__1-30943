VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Background Options"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framPreview 
      Caption         =   "Preview"
      Height          =   3675
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2295
      Begin VB.Timer timPreview 
         Interval        =   100
         Left            =   0
         Top             =   3120
      End
      Begin VB.CommandButton cmdTitleClose 
         Cancel          =   -1  'True
         Caption         =   "X"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   280
         Width           =   240
      End
      Begin VB.Line lnSecond 
         X1              =   1080
         X2              =   600
         Y1              =   1440
         Y2              =   1920
      End
      Begin VB.Line lnMinute 
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   720
         Y2              =   1440
      End
      Begin VB.Line lnHour 
         BorderWidth     =   2
         X1              =   1560
         X2              =   1080
         Y1              =   1680
         Y2              =   1440
      End
      Begin VB.Label lblPrevTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   365
         Left            =   120
         TabIndex        =   6
         Top             =   2450
         Width           =   2055
      End
      Begin VB.Label lblPrevDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2810
         Width           =   2055
      End
      Begin VB.Label lblPrevDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3180
         Width           =   2055
      End
      Begin VB.Image imgTitleIcon 
         Height          =   230
         Left            =   160
         Picture         =   "frmBack.frx":030A
         Stretch         =   -1  'True
         Top             =   280
         Width           =   230
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   295
         Width           =   495
      End
      Begin VB.Shape shpTitleBar 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   300
         Left            =   120
         Top             =   250
         Width           =   2060
      End
      Begin VB.Label lblPreview 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   1890
         Left            =   120
         TabIndex        =   1
         Top             =   555
         Width           =   2055
      End
      Begin MSForms.Image imgPreview 
         Height          =   2990
         Left            =   135
         Top             =   575
         Width           =   2040
         BorderStyle     =   0
         Size            =   "3598;5274"
         PictureAlignment=   0
      End
      Begin MSForms.Image imgBackRaise 
         Height          =   3405
         Left            =   60
         Top             =   195
         Width           =   2160
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "3810;6006"
         VariousPropertyBits=   19
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   315
      ItemData        =   "frmBack.frx":0614
      Left            =   0
      List            =   "frmBack.frx":0621
      TabIndex        =   10
      Text            =   "[None]"
      Top             =   1080
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmndlgBack 
      Left            =   0
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin MSForms.CheckBox chkOnOff 
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
      VariousPropertyBits=   746596371
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      MousePointer    =   1
      Size            =   "2143;873"
      Value           =   "0"
      Caption         =   "Background On/Off"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblAlign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   690
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Picture"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmBack"
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

Const Pi = 3.14159265358979

Dim HourType As Boolean 'if 24h or not

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

'logo path name
Dim GotName As String


Private Sub Stretch(BoolVal As Boolean)
'this stretchs the picture to fit the clock

If BoolVal Then
    'if checked, then stretch, else show normal
    imgPreview.PictureSizeMode = fmPictureSizeModeStretch
    imgPreview.PictureTiling = False
Else
    imgPreview.PictureSizeMode = fmPictureSizeModeClip
End If
End Sub

Private Sub Tile(BoolVal As Boolean)
'this tiles any picture small enough to be tiled
imgPreview.PictureTiling = BoolVal

If BoolVal Then
    'turn off stretch
    imgPreview.PictureSizeMode = fmPictureSizeModeClip
End If
End Sub

Private Sub chkOnOff_Click()
Dim Style As Integer

'turn logo on or off
BackOnOff = chkOnOff.Value
frmHandsClk.mnuFileBackOn.Checked = BackOnOff


'set the label border style from a boolean value.
Style = (BackOnOff * -1)
If Style = 0 Then
    Style = 1
Else
    Style = 0
End If

'show the colours in the preview pane
lblPreview.BackStyle = Style
lblPrevTime.BackStyle = Style
lblPrevDay.BackStyle = Style
lblPrevDate.BackStyle = Style

'show selected option and save
Call frmHandsClk.ShowLogo
Call frmHandsClk.SaveStatus
End Sub

Private Sub cmbAlign_Change()
'stretch or tile the background picture. (see above)
Select Case LCase(Trim(cmbAlign.Text))
Case "[none]"
    Call Tile(False)
    Call Stretch(False)
    
Case "tile"
    Call Tile(True)
    Call Stretch(False)
    
Case "stretch"
    Call Tile(False)
    Call Stretch(True)
    
End Select
End Sub

Private Sub cmbAlign_Click()
'see above
Call cmbAlign_Change
End Sub

Private Sub cmbAlign_KeyPress(KeyAscii As Integer)
'don't permit the user to change the details
KeyAscii = 0
End Sub

Private Sub cmdBrowse_Click()
'call an open dialogue box to find a picture

'picture types supported by image control
Const Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur|Bitmaps (*.Bmp; *.Dib)|*.bmp;*.dib|Gif Images (*.Gif)|*.gif|Jpeg Images (*.Jpg)|*.jpg|Metafiles (*.Wmf; *.Emf)|*.wmf;*.emf|Icons (*.Ico; *.Cur)|*.ico;*.cur|All Files (*.)|*.*"

cmndlgBack.Filter = Filter

'account for the program being in the root directory
'(double backslash problem)
If GetPath(BackPath) = "" Then
    cmndlgBack.InitDir = App.Path
Else
    cmndlgBack.InitDir = GetPath(BackPath)
End If

cmndlgBack.ShowOpen

GotName = cmndlgBack.FileName

'if name is a file and it exists, then load
If (GotName <> "") And (Dir(GotName) <> "") Then
    txtPath.Text = GotName
    
    'load picture into preview box
    imgPreview.Picture = LoadPicture(GotName)
End If
End Sub

Private Sub cmdCancel_Click()
'unload form without changing background settings
Unload Me
End Sub

Private Sub cmdSet_Click()
'set logo into clock, but not necessaraly turn the
'logo on.
BackPath = txtPath.Text
StretchTile = cmbAlign.Text

'display selected option
frmHandsClk.imgLogo.Picture = imgPreview.Picture
Call frmHandsClk.ShowLogo

'save settings
Call frmHandsClk.SaveStatus
End Sub

Private Sub Form_Activate()
'get .ini values and use them for the preview screen
Call GetData

'set colours of time
lnHour.BorderColor = ColHour
lnMinute.BorderColor = ColMin
lnSecond.BorderColor = ColSec
lblPreview.BackColor = CAnBak
lblPrevTime.ForeColor = CTmFon
lblPrevTime.BackColor = CTmBak
lblPrevDay.ForeColor = CDyFon
lblPrevDay.BackColor = CDyBak
lblPrevDate.ForeColor = CDtFon
lblPrevDate.BackColor = CDtBak

txtPath.SetFocus
txtPath.SelStart = 0
txtPath.SelLength = Len(txtPath.Text)
End Sub

Private Sub timPreview_Timer()
'show clock hands for the preview clock and current
'time

Static NormalTime As String
Static SecAngle As Integer
Static MinAngle As Integer
Static HourAngle As Integer

Dim Counter As Integer
Dim AnaCentreX As Integer
Dim AnaCentreY As Integer
Dim DotX As Integer
Dim DotY As Integer

Const DotRadius = 830
Const SecRadius = 790
Const MinRadius = 750
Const HourRadius = 580
Const CounterWeight = 50    'the bit past the centre opposite where the hand is pointing to
Const CounterWeightSec = 250

'set values

'set the clock centre for the clock hands
AnaCentreX = framPreview.Left + lblPreview.Left - (lblPreview.Width / 8)
AnaCentreY = framPreview.Top + lblPreview.Top + 180

'Digital aspect
'--------------------------------------------------
'calculate time in 12 hour clock
NormalTime = (Hour(Time) Mod 12)
If NormalTime = "0" Then NormalTime = "12"
NormalTime = NormalTime & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")

If lblPrevTime.Caption <> Time Then
    If HourType Then '24 hour clock
        'show current time
        lblPrevTime.Caption = Time
    Else
        lblPrevTime.Caption = NormalTime
    End If
End If
    
If lblPrevDate.Caption <> Date Then
    lblPrevDay.Caption = GetDay
    lblPrevDate.Caption = Date
End If
'--------------------------------------------------

'Analogue Aspect
'--------------------------------------------------
SecAngle = (180 + (Second(Time) * -6)) Mod 360
MinAngle = (180 + (Minute(Time) * -6)) Mod 360
HourAngle = (180 + ((Hour(Time) * -30)) - (Minute(Time) / 2)) Mod 360

'get points for clock hands
lnSecond.X1 = AnaCentreX + (Sin(SecAngle * Pi / 180) * SecRadius)
lnSecond.Y1 = AnaCentreY + (Cos(SecAngle * Pi / 180) * SecRadius)
lnSecond.X2 = AnaCentreX + (Sin(SecAngle * Pi / 180) * CounterWeightSec)
lnSecond.Y2 = AnaCentreY + (Cos(SecAngle * Pi / 180) * CounterWeightSec)

lnMinute.X1 = AnaCentreX + (Sin(MinAngle * Pi / 180) * MinRadius)
lnMinute.Y1 = AnaCentreY + (Cos(MinAngle * Pi / 180) * MinRadius)
lnMinute.X2 = AnaCentreX - (Sin(MinAngle * Pi / 180) * CounterWeight)
lnMinute.Y2 = AnaCentreY - (Cos(MinAngle * Pi / 180) * CounterWeight)

lnHour.X1 = AnaCentreX + (Sin(HourAngle * Pi / 180) * HourRadius)
lnHour.Y1 = AnaCentreY + (Cos(HourAngle * Pi / 180) * HourRadius)
lnHour.X2 = AnaCentreX - (Sin(HourAngle * Pi / 180) * CounterWeight)
lnHour.Y2 = AnaCentreY - (Cos(HourAngle * Pi / 180) * CounterWeight)

'print the minute dots
'NOTE : doesn't work. Frame will not show dots.
'--------------------------------------------------
'show big ones at 5 minute intervals
'If frmBack.ForeColor <> ColDot Then
'    frmBack.ForeColor = ColDot
'End If

'frmBack.DrawWidth = 2
'For Counter = 1 To 360 Step 30
'    'framPreview.Left + AnaCentreX
'    DotX = (framPreview.Left + AnaCentreX) + (Sin(Counter * Pi / 180) * DotRadius)
'    DotY = (framPreview.Top + AnaCentreY) + (Cos(Counter * Pi / 180) * DotRadius)
'    Line (DotX, DotY)-(DotX + 5, DotY + 5)
'Next Counter

'show small dots at one minute intervals
'frmBack.DrawWidth = 1
'For Counter = 1 To 360 Step 6
'    DotX = (framPreview.Left + AnaCentreX) + (Sin(Counter * Pi / 180) * DotRadius)
'    DotY = (framPreview.Height + AnaCentreY) + (Cos(Counter * Pi / 180) * DotRadius)
'    PSet (DotX, DotY)
'Next Counter
'--------------------------------------------------

End Sub

Public Sub GetData()
'load values form the SetClock.ini file for use by
'other procedures

Dim Check As String
Dim Path As String
Dim ErrorNum As Integer

'find any problems with .ini file
On Error Resume Next
Open FilePath For Input As #30
    ErrorNum = Err
Close #30
On Error GoTo 0


'if no problem with .ini file...
If ErrorNum = 0 Then
    Open FilePath For Input As #14
    
    'load values line by line
    While Not EOF(14)
        Line Input #14, Check
        
        Select Case LCase(GetBefore(Check))
        Case "24hour"
            If LCase(GetAfter(Check)) = "no" Then
                HourType = False
            Else
                HourType = True
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
            
            If (BackPath = "") Or (BackPath = "0") Or (Dir(BackPath) = "") Then
                BackPath = ""
                'clear preview picture
                imgPreview.Picture = LoadPicture
            Else
                imgPreview.Picture = LoadPicture(BackPath)
                txtPath.Text = BackPath
            End If
            
            'load picture into preview box
            'frmHandsClk.Picture = frmHandsClk.imgLogo.Picture
            
        Case "backtile"
            StretchTile = GetAfter(Check)
            Select Case LCase(StretchTile)
            Case "[none]"
                cmbAlign.ListIndex = 0
                Call Tile(False)
                Call Stretch(False)
            Case "tile"
                cmbAlign.ListIndex = 1
                Call Tile(True)
                Call Stretch(False)
            Case "stretch"
                cmbAlign.ListIndex = 2
                Call Tile(False)
                Call Stretch(True)
            End Select
            
        Case "backonoff"
            BackOnOff = GetAfter(Check)
            
        End Select
    Wend
    
    Close #14
    
    'changing the value of a check box automatically
    'calls its click event
    chkOnOff.Value = BackOnOff

End If
End Sub

Public Function GetDay() As String
'get name of current day
Select Case Weekday(Date)
Case 1
    GetDay = "Sunday"
Case 2
    GetDay = "Monday"
Case 3
    GetDay = "Tuesday"
Case 4
    GetDay = "Wednesday"
Case 5
    GetDay = "Thursday"
Case 6
    GetDay = "Friday"
Case 7
    GetDay = "Saturday"
End Select

End Function

Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)
Const DeleteKey = 46

'prevent user from changing data
If KeyCode = DeleteKey Then KeyCode = 0
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
'prevent user from entering/changing data
KeyAscii = 0
End Sub
