VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmScheme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colour Schemes"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmScheme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   3
      Left            =   1410
      TabIndex        =   29
      Top             =   2220
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   10
      Left            =   1410
      TabIndex        =   22
      Top             =   4320
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   9
      Left            =   1410
      TabIndex        =   23
      Top             =   4020
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   8
      Left            =   1410
      TabIndex        =   24
      Top             =   3720
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   7
      Left            =   1410
      TabIndex        =   25
      Top             =   3420
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   6
      Left            =   1410
      TabIndex        =   26
      Top             =   3120
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   5
      Left            =   1410
      TabIndex        =   27
      Top             =   2820
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   4
      Left            =   1410
      TabIndex        =   28
      Top             =   2520
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   2
      Left            =   1410
      TabIndex        =   32
      Top             =   1920
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   31
      Top             =   1620
      Width           =   285
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   255
      Index           =   0
      Left            =   1410
      TabIndex        =   30
      Top             =   1320
      Width           =   285
   End
   Begin VB.Frame framPreview 
      Caption         =   "Preview"
      Height          =   3675
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   280
         Width           =   240
      End
      Begin VB.Timer timPreview 
         Interval        =   100
         Left            =   0
         Top             =   3120
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
         TabIndex        =   9
         Top             =   295
         Width           =   495
      End
      Begin VB.Label lblPreview 
         BorderStyle     =   1  'Fixed Single
         Height          =   1890
         Left            =   120
         TabIndex        =   10
         Top             =   555
         Width           =   2055
      End
      Begin VB.Label lblPrevTime 
         Alignment       =   2  'Center
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
         TabIndex        =   7
         Top             =   2810
         Width           =   2055
      End
      Begin VB.Label lblPrevDate 
         Alignment       =   2  'Center
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
         TabIndex        =   8
         Top             =   3180
         Width           =   2055
      End
      Begin VB.Image imgTitleIcon 
         Height          =   230
         Left            =   160
         Picture         =   "frmScheme.frx":20F2
         Stretch         =   -1  'True
         Top             =   280
         Width           =   230
      End
      Begin MSForms.Image imgPreview 
         Height          =   2990
         Left            =   135
         Top             =   575
         Width           =   2040
         BorderStyle     =   0
         Size            =   "3598;5274"
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
   Begin VB.CommandButton cmdScheme 
      Caption         =   "&Set Scheme"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Scheme"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox cmbColour 
      Height          =   315
      ItemData        =   "frmScheme.frx":23FC
      Left            =   1440
      List            =   "frmScheme.frx":23FE
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Original Colours"
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cdlgColour 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDots 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dot Colour"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label lblDtBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Background"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblDtFont 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Font Colour"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   4020
      Width           =   1695
   End
   Begin VB.Label lblDyBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day Background"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblDyFont 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day Font Colour"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   3420
      Width           =   1695
   End
   Begin VB.Label lblTmBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Background"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTmFont 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Font Colour"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2820
      Width           =   1695
   End
   Begin VB.Label lblHandBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hand Background"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Second Hand"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblMin 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minute Hand"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label lblHour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hour Hand"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblColour 
      BackStyle       =   0  'Transparent
      Caption         =   "Colour Schemes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgBack 
      Height          =   840
      Left            =   120
      Picture         =   "frmScheme.frx":2400
      Stretch         =   -1  'True
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form was introduced in program version 6.0
'This will change/save any colour schemes created
'by the user. This form uses a new .ini file called
'"ColSceme.ini" to store any colour schemes. Should
'this file not exist, it will create the file and
'enter all the default colour schemes.
'23/11/2000     -      DiskJunky
'==================================================

Option Base 1


Const IniFile = "ColSceme.ini"
Const vbName = "[Default]"
Const Pi = 3.14159265358979

'default colours
Private Enum DefaultCol
    vbHourHand = 16711935         'light purple
    vbMinuteHand = 12583104       'purple
    vbSecondHand = 8421631        'light red
    vbAnalogueBackground = 65535  'yellow
    vbDots = -2147483630          'black
    vbTimeFont = 8388736          'dark purple
    vbTimeBackground = 65535      'yellow
    vbDayFont = 65535             'yellow
    vbDayBackground = 12583104     'purple
    vbDateFont = 65535            'yellow
    vbDateBackground = 12583104    'purple
End Enum

Private Enum Colours
    vbHour = 1
    vbMinute = 2
    vbSecond = 3
    Dots = 4
    HandBack = 5
    TimeFont = 6
    TimeBack = 7
    DayFont = 8
    DayBack = 9
    DateFont = 10
    DateBack = 11
End Enum


'a new data type to hold the colours
Private Type ColScheme
    Name As String
    HourHand As Long
    MinuteHand As Long
    SecondHand As Long
    AnalogueBack As Long
    Dots As Long
    TimeFont As Long
    TimeBack As Long
    DayFont As Long
    DayBack As Long
    DateFont As Long
    DateBack As Long
End Type

Dim HourType As Boolean 'if 24h or not

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

'set up a large array of possible colour schemes
Dim Schemes(32000) As ColScheme
Dim ColIndex As Integer
Dim MaxIndex As Integer

Dim IniPath As String

Public Sub LoadSchemes()
'This procedure loads the colour schemes into
'an array storing the necessary data.
Dim FileNum As Integer
Dim ErrNum As Integer
Dim Check As String


MaxIndex = 1

'trap any file errors that night occur
On Error Resume Next
FileNum = FreeFile
Open IniPath For Input As #FileNum
    ErrNum = Err
Close #FileNum
On Error GoTo 0

'do not continue to load if there are problems
'with the ini file.
If ErrNum <> 0 Then
    If ErrNum = 53 Then
        'if "File Not Found" then create file and
        'load the default values.
        Call CreateDefaults
        Call LoadSchemes
        Exit Sub
    End If
    
    Exit Sub
End If

MaxIndex = 0
Open IniPath For Input As #FileNum
    While Not EOF(FileNum)
        'reterieve string data
        Line Input #FileNum, Check
        
        Select Case LCase(GetBefore(Check))
        'load each colour scheme into the array
        Case "name"
            'increment array counter and load a new
            'colour scheme.
            MaxIndex = MaxIndex + 1
            Schemes(MaxIndex).Name = GetAfter(Check)
            
            'add scheme to combo box
            cmbColour.AddItem Schemes(MaxIndex).Name, (MaxIndex - 1)
            
        Case "hourhand"
            Schemes(MaxIndex).HourHand = GetAfter(Check)
            
        Case "minhand"
            Schemes(MaxIndex).MinuteHand = GetAfter(Check)
        
        Case "sechand"
            Schemes(MaxIndex).SecondHand = GetAfter(Check)
        
        Case "analogueback"
            Schemes(MaxIndex).AnalogueBack = GetAfter(Check)
        
        Case "dots"
            Schemes(MaxIndex).Dots = GetAfter(Check)
        
        Case "timefont"
            Schemes(MaxIndex).TimeFont = GetAfter(Check)
        
        Case "timeback"
            Schemes(MaxIndex).TimeBack = GetAfter(Check)
        
        Case "dayfont"
            Schemes(MaxIndex).DayFont = GetAfter(Check)
        
        Case "dayback"
            Schemes(MaxIndex).DayBack = GetAfter(Check)
        
        Case "datefont"
            Schemes(MaxIndex).DateFont = GetAfter(Check)
        
        Case "dateback"
            Schemes(MaxIndex).DateBack = GetAfter(Check)
        
        End Select
    Wend
Close #FileNum

'account for an empty file (the default values are
'ALWAYS Maxindex 1.
If MaxIndex = 0 Then
    Call CreateDefaults
    Call LoadSchemes
    Exit Sub
End If

cmbColour.ListIndex = 0
End Sub


Public Sub SaveSchemes()
'This procedure save the colour schemes into an
'ini file for later reading.

Dim Counter As Integer
Dim FileNum As Integer
Dim ErrNum As Integer

FileNum = FreeFile

'trap any file errors that night occur
On Error Resume Next
Open IniPath For Output As #FileNum
    ErrNum = Err
Close #FileNum
On Error GoTo 0

'do not continue to load if there are problems
'with the ini file.
If ErrNum <> 0 Then
    Exit Sub
End If


Open IniPath For Output As #FileNum
    'Print headers
    Print #FileNum, "This file is for use by CompApp Clock " & App.Major & "." & App.Minor & "." & App.Revision
    Print #FileNum, "and is not to be tampered with or the program"
    Print #FileNum, "may cause undesired effects."
    Print #FileNum, "- DiskJunky"
    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[DEFAULT COLOUR SCHEME]"
    
    'store defaults
    Print #FileNum, "Name=" & vbName
    Print #FileNum, "SecHand=" & DefaultCol.vbSecondHand
    Print #FileNum, "MinHand=" & DefaultCol.vbMinuteHand
    Print #FileNum, "HourHand=" & DefaultCol.vbHourHand
    Print #FileNum, "AnalogueBack=" & DefaultCol.vbAnalogueBackground
    Print #FileNum, "Dots=" & DefaultCol.vbDots
    Print #FileNum, "TimeFont=" & DefaultCol.vbTimeFont
    Print #FileNum, "TimeBack=" & DefaultCol.vbTimeBackground
    Print #FileNum, "DayFont=" & DefaultCol.vbDayFont
    Print #FileNum, "DayBack=" & DefaultCol.vbDayBackground
    Print #FileNum, "DateFont=" & DefaultCol.vbDateFont
    Print #FileNum, "DateBack=" & DefaultCol.vbDateBackground

    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[USER DEFINED COLOUR SCHEMES]"
    Print #FileNum, ""
    
    'store each colour scheme (omitting [default])
    For Counter = 2 To MaxIndex
        With Schemes(Counter)
            'if scheme is not empty, then save
            If .Name <> "" Then
                Print #FileNum, "Name=" & .Name
                Print #FileNum, "SecHand=" & .SecondHand
                Print #FileNum, "MinHand=" & .MinuteHand
                Print #FileNum, "HourHand=" & .HourHand
                Print #FileNum, "AnalogueBack=" & .AnalogueBack
                Print #FileNum, "Dots=" & .Dots
                Print #FileNum, "TimeFont=" & .TimeFont
                Print #FileNum, "TimeBack=" & .TimeBack
                Print #FileNum, "DayFont=" & .DayFont
                Print #FileNum, "DayBack=" & .DayBack
                Print #FileNum, "DateFont=" & .DateFont
                Print #FileNum, "DateBack=" & .DateBack
            End If
        End With
        
        Print #FileNum, ""
    Next Counter
    
Close #FileNum

End Sub


Private Sub CreateDefaults()
'This file will wipe any stored colour schemes and
'restore the original colour values.

Dim FileNum As Integer
Dim ErrNum As Integer

FileNum = FreeFile

'trap any file errors that night occur
On Error Resume Next
Open IniPath For Output As #FileNum
    ErrNum = Err
Close #FileNum
On Error GoTo 0

'do not continue to load if there are problems
'with the ini file.
If ErrNum <> 0 Then
    Exit Sub
End If

Open IniPath For Output As #FileNum
    'Print headers
    Print #FileNum, "This file is for use by CompApp Clock " & App.Major & "." & App.Minor & "." & App.Revision
    Print #FileNum, "and is not to be tampered with or the program"
    Print #FileNum, "may cause undesired effects."
    Print #FileNum, "- DiskJunky"
    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[DEFAULT COLOUR SCHEME]"
    
    'store defaults
    Print #FileNum, "Name=" & vbName
    Print #FileNum, "SecHand=" & DefaultCol.vbSecondHand
    Print #FileNum, "MinHand=" & DefaultCol.vbMinuteHand
    Print #FileNum, "HourHand=" & DefaultCol.vbHourHand
    Print #FileNum, "AnalogueBack=" & DefaultCol.vbAnalogueBackground
    Print #FileNum, "Dots=" & DefaultCol.vbDots
    Print #FileNum, "TimeFont=" & DefaultCol.vbTimeFont
    Print #FileNum, "TimeBack=" & DefaultCol.vbTimeBackground
    Print #FileNum, "DayFont=" & DefaultCol.vbDayFont
    Print #FileNum, "DayBack=" & DefaultCol.vbDayBackground
    Print #FileNum, "DateFont=" & DefaultCol.vbDateFont
    Print #FileNum, "DateBack=" & DefaultCol.vbDateBackground

    Print #FileNum, ""

Close #FileNum
End Sub

Private Sub cmbColour_Change()
'if a scheme was selected, then
If cmbColour.ListIndex >= 0 Then
    Call LoadColoursFromArray(cmbColour.ListIndex + 1)
    Call PreviewSetColours '(cmbColour.ListIndex + 1)
End If
End Sub

Private Sub cmbColour_Click()
'if a scheme was selected, then
If cmbColour.ListIndex >= 0 Then
    Call LoadColoursFromArray(cmbColour.ListIndex + 1)
    Call PreviewSetColours '(cmbColour.ListIndex + 1)
End If
End Sub

Private Sub cmbColour_KeyPress(KeyAscii As Integer)
'if a scheme was selected, then
If KeyAscii = 13 Then
    Call LoadColoursFromArray(cmbColour.ListIndex + 1)
    Call PreviewSetColours '(cmbColour.ListIndex + 1)
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdColour_Click(Index As Integer)
Dim GotMatch As Integer

'show colour dialogue box and save the colour
cdlgColour.Color = GetColour(Index + 1)
cdlgColour.FLAGS = cdlCCRGBInit
cdlgColour.ShowColor

Call SetColour((Index + 1), cdlgColour.Color)

'see if changed colours mathc any colour scheme
GotMatch = FindMatch
If GotMatch <> 0 Then
    cmbColour.ListIndex = GotMatch - 1
Else
    cmbColour.Text = "New Scheme"
End If
End Sub

Private Sub cmdSave_Click()
Dim GotMatch As Integer
Dim Counter As Integer
Dim Multiple As Integer

Multiple = 0
'account for multiple names
For Counter = 1 To MaxIndex
    If LCase(cmbColour.Text) = LCase(Schemes(Counter).Name) Then
        Multiple = True
        Exit For
    End If
Next Counter

'If Multiple Then
'    'cannot save scheme
'    cmbColour.SelStart = 0
'    cmbColour.SelLength = Len(cmbColour.Text)
'    Exit Sub
'End If

'overwrite data if necessary
If Multiple = 0 Then
    'enter data into the array and save the values.
    MaxIndex = MaxIndex + 1
End If

Schemes(MaxIndex).Name = cmbColour.Text
Schemes(MaxIndex).HourHand = ColHour
Schemes(MaxIndex).MinuteHand = ColMin
Schemes(MaxIndex).SecondHand = ColSec
Schemes(MaxIndex).Dots = ColDot
Schemes(MaxIndex).AnalogueBack = CAnBak
Schemes(MaxIndex).TimeFont = CTmFon
Schemes(MaxIndex).TimeBack = CTmBak
Schemes(MaxIndex).DayFont = CDyFon
Schemes(MaxIndex).DayBack = CDyBak
Schemes(MaxIndex).DateFont = CDtFon
Schemes(MaxIndex).DateBack = CDtBak

Call SaveSchemes
cmbColour.Clear
Call LoadSchemes

'see if changed colours mathc any colour scheme
GotMatch = FindMatch
If GotMatch <> 0 Then
    cmbColour.ListIndex = GotMatch - 1
Else
    'default
    cmbColour.ListIndex = 0
End If

End Sub

Private Sub cmdScheme_Click()
'apply the current scheme to the actual clock.

Dim GotMatch As Integer

'if no item is currently selected, then exit sub-routine
If cmbColour.ListIndex = -1 Then
    Exit Sub
End If

'apply the scheme
Call SetScheme(cmbColour.ListIndex + 1)

'display the current scheme in the combo box
GotMatch = FindMatch
If GotMatch > 0 Then
    cmbColour.ListIndex = GotMatch - 1
End If
End Sub

Private Sub Form_Activate()
Dim GotMatch As Integer

'see if the current colour settings match any of the
'colour schemes
GotMatch = FindMatch

If GotMatch > 0 Then
    'match found, show scheme name
    cmbColour.ListIndex = (GotMatch - 1)
Else
    'prompt to save scheme
    cmbColour.Text = "New Scheme"
    cmbColour.SelStart = 0
    cmbColour.SelLength = Len(cmbColour.Text)
End If

End Sub

Private Sub Form_Load()
'get the current path of the colour scheme ini file.
IniPath = AddFile(App.Path, IniFile)

'load schemes into the array
Call LoadSchemes

'set the preview pane
Call PreviewData
End Sub

Public Function FindMatch() As Integer
'This function will return the array index of the
'current array element matching the colour settings
'of the clock. It will return zero if not found.
'(hence the reason for Option Base 1)

'Colour variables
'Dim ColHour As Long
'Dim ColMin As Long
'Dim ColSec As Long
'Dim ColDot As Long
'Dim CAnBak As Long
'Dim CTmFon As Long   'time font
'Dim CTmBak As Long   'time back ground
'Dim CDyFon As Long   'day font
'Dim CDyBak As Long   'day ...
'Dim CDtFon As Long   'date ..
'Dim CDtBak As Long   'date ..

Dim Counter As Integer
Dim RetIndex As Integer

RetIndex = 0

'set the current value to compare with
'ColHour = frmHandsClk.lnHour.BorderColor
'ColMin = frmHandsClk.lnMinute.BorderColor
'ColSec = frmHandsClk.lnSecond.BorderColor
'ColDot = frmHandsClk.ForeColor
'CAnBak = frmHandsClk.BackColor
'CTmFon = frmHandsClk.lblShowTime.ForeColor
'CTmBak = frmHandsClk.lblShowTime.BackColor
'CDyFon = frmHandsClk.lblShowDay.ForeColor
'CDyBak = frmHandsClk.lblShowDay.BackColor
'CDtFon = frmHandsClk.lblShowDate.ForeColor
'CDtBak = frmHandsClk.lblShowDate.BackColor


'compare values
For Counter = 1 To MaxIndex
    If (Schemes(Counter).HourHand = ColHour) And _
       (Schemes(Counter).MinuteHand = ColMin) And _
       (Schemes(Counter).SecondHand = ColSec) And _
       (Schemes(Counter).AnalogueBack = CAnBak) And _
       (Schemes(Counter).Dots = ColDot) And _
       (Schemes(Counter).TimeFont = CTmFon) And _
       (Schemes(Counter).TimeBack = CTmBak) And _
       (Schemes(Counter).DateFont = CDtFon) And _
       (Schemes(Counter).DateBack = CDtBak) And _
       (Schemes(Counter).DayFont = CDyFon) And _
       (Schemes(Counter).DayBack = CDyBak) Then
        
        'match found, return index
        RetIndex = Counter
        Exit For
    End If
Next Counter

'return result
FindMatch = RetIndex
End Function

Private Sub SetScheme(Index As Integer)
'This procedure will take the current scheme
'selected and apply it to the clock.
'Index represents the array element number.

'assign colours
frmHandsClk.lnHour.BorderColor = Schemes(Index).HourHand
frmHandsClk.lnMinute.BorderColor = Schemes(Index).MinuteHand
frmHandsClk.lnSecond.BorderColor = Schemes(Index).SecondHand
frmHandsClk.ForeColor = Schemes(Index).Dots
frmHandsClk.lblShowHands.BackColor = Schemes(Index).AnalogueBack
frmHandsClk.lblShowTime.ForeColor = Schemes(Index).TimeFont
frmHandsClk.lblShowTime.BackColor = Schemes(Index).TimeBack
frmHandsClk.lblShowDay.ForeColor = Schemes(Index).DayFont
frmHandsClk.lblShowDay.BackColor = Schemes(Index).DayBack
frmHandsClk.lblShowDate.ForeColor = Schemes(Index).DateFont
frmHandsClk.lblShowDate.BackColor = Schemes(Index).DateBack

'save values
frmHandsClk.SaveStatus
End Sub

'------------------------------------------------
'This section contains the code that operates the
'preview frame and animates the preview clock.
'------------------------------------------------

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
AnaCentreX = (framPreview.Width / 2) 'framPreview.Left + lblPreview.Left -
AnaCentreY = (framPreview.Height / 2) - 400 ' + lblPreview.Top + 180

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
'    PSet (DotX, DotY)
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

Private Sub PreviewData()
'load values form the SetClock.ini file for use by
'other procedures

Dim Check As String
Dim Path As String
Dim BackPath As String
Dim StretchTile As String
Dim ErrorNum As Integer
Dim FilePath As String

'set the path of the SetClock.ini file
FilePath = AddFile(App.Path, "SetClock.Ini")

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
            'BackPath = GetAfter(Check)
            
            'If BackPath = "0" Then
            '    BackPath = ""
            '    'clear preview picture
            '    imgPreview.Picture = LoadPicture
            'Else
            '    imgPreview.Picture = LoadPicture(BackPath)
            '    'txtPath.Text = BackPath
            'End If
            
            'load picture into preview box
            'frmHandsClk.Picture = frmHandsClk.imgLogo.Picture
            
        Case "backtile"
            'StretchTile = GetAfter(Check)
            'Select Case LCase(StretchTile)
            'Case "[none]"
            '    'cmbAlign.ListIndex = 0
            '    Call Tile(False)
            '    Call Stretch(False)
            'Case "tile"
            '    'cmbAlign.ListIndex = 1
            '    Call Tile(True)
            '    Call Stretch(False)
            'Case "stretch"
            '    'cmbAlign.ListIndex = 2
            '    Call Tile(False)
            '    Call Stretch(True)
            'End Select
            
        Case "backonoff"
            'BackOnOff = GetAfter(Check)
            
        End Select
    Wend
    
    Close #14
    
    'changing the value of a check box automatically
    'calls its click event
    'chkOnOff.Value = BackOnOff
    
    'set the preview colours
    Call PreviewSetColours
End If
End Sub

Private Sub PreviewSetColours()
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
End Sub

Private Sub LoadColoursFromArray(Index As Integer)
'get the colour from the array
cmbColour.Text = Schemes(Index).Name
ColHour = Schemes(Index).HourHand
ColMin = Schemes(Index).MinuteHand
ColSec = Schemes(Index).SecondHand
ColDot = Schemes(Index).Dots
CAnBak = Schemes(Index).AnalogueBack
CTmFon = Schemes(Index).TimeFont
CTmBak = Schemes(Index).TimeBack
CDyFon = Schemes(Index).DayFont
CDyBak = Schemes(Index).DayBack
CDtFon = Schemes(Index).DateFont
CDtBak = Schemes(Index).DateBack
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

'------------------------------------------------
'Preview section finished.
'------------------------------------------------

Private Sub SetColour(Object As Colours, Colour As Long)
'This sets a given objects colours.

Select Case LCase(Object)
Case vbHour
    ColHour = Colour
Case vbMinute
    ColMin = Colour
Case vbSecond
    ColSec = Colour
Case Dots
    ColDot = Colour
Case HandBack
    CAnBak = Colour
Case TimeFont
    CTmFon = Colour
Case TimeBack
    CTmBak = Colour
Case DayFont
    CDyFon = Colour
Case DayBack
    CDyBak = Colour
Case DateFont
    CDtFon = Colour
Case DateBack
    CDtBak = Colour
End Select

Call PreviewSetColours
End Sub

Private Function GetColour(Colour As Colours) As Long
'returns the colour of the selected object.

Select Case LCase(Colour)
Case vbHour
    GetColour = ColHour
Case vbMinute
    GetColour = ColMin
Case vbSecond
    GetColour = ColSec
Case Dots
    GetColour = ColDot
Case HandBack
    GetColour = CAnBak
Case TimeFont
    GetColour = CTmFon
Case TimeBack
    GetColour = CTmBak
Case DayFont
    GetColour = CDyFon
Case DayBack
    GetColour = CDyBak
Case DateFont
    GetColour = CDtFon
Case DateBack
    GetColour = CDtBak

Case Else
    GetColour = 0
End Select

End Function
