VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timed ShutDown Clock"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   2265
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      Picture         =   "frmAbout.frx":14AD4
      ScaleHeight     =   1425
      ScaleWidth      =   4425
      TabIndex        =   5
      Top             =   0
      Width           =   4455
   End
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1845
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line lnBreak 
      X1              =   120
      X2              =   4320
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed To ;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      TabIndex        =   3
      Top             =   1710
      Width           =   705
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2000"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      TabIndex        =   2
      Top             =   1575
      Width           =   885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Eric O'Sullivan"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "The Clock was made by..."
      Top             =   1575
      Width           =   1575
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   1495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
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

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim AboutHeight As Integer
Dim Tick As Long
Dim Speed As Integer
Dim ScrollBmp As BitmapStruc
Dim ScrollText As String

Private Sub cmdOk_Click()
timScroll.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Dim NextLine As String
'Dim GotInfo As String
Dim Tada As String
Dim Result As Long

'play the 'TaDa.Wav' file.
Tada = AddFile(App.Path, "Tada.Wav")

Call PlaySound(Tada)

'set the spacer line size
lnBreak.X1 = Screen.TwipsPerPixelX * 8
lnBreak.X2 = (Me.ScaleWidth) - (Screen.TwipsPerPixelX * 8)

'start scrolling from the start
Tick = 0

'set scrolling speed (one pixel per tick)
Speed = Screen.TwipsPerPixelY

'create the background and copy picture onto it
ScrollBmp.Area.Bottom = picText.Height / Screen.TwipsPerPixelY
ScrollBmp.Area.Right = picText.Width / Screen.TwipsPerPixelX
Call CreateNewBitmap(ScrollBmp.hDcMemory, ScrollBmp.hDcBitmap, ScrollBmp.hDcPointer, ScrollBmp.Area, frmAbout, picText.BackColor, InPixels)

'set the display text
NextLine = Chr(13) & Chr(10)
ScrollText = "Timed ShutDown Clock  v " & App.Major & "." & App.Minor
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & "CompApp Technologies™"
ScrollText = ScrollText & vbCrLf & "Copyright 2000"
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & "Support"
ScrollText = ScrollText & vbCrLf & "If there are any problems with this"
ScrollText = ScrollText & vbCrLf & "product, please don't hesitate to "
ScrollText = ScrollText & vbCrLf & "contact us."
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & "E-mail"
ScrollText = ScrollText & vbCrLf & "DiskJunky@hotmail.com"
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & "Web Site"
ScrollText = ScrollText & vbCrLf & "http://www.compapp.co-ltd.com"
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & "Programmer"
ScrollText = ScrollText & vbCrLf & "Eric O'Sullivan"
ScrollText = ScrollText & vbCrLf & ""
ScrollText = ScrollText & vbCrLf & ""

'Call EnterText(GotInfo)

'centre command button
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2

'display version and author information and Owner
lblVersion.AutoSize = True
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "        By Eric O' Sullivan"
If lblVersion.Left > 0 Then
    lblVersion.Left = Me.ScaleWidth - lblVersion.Width - 140
End If

If Owner = "" Then
    lblOwner.Caption = "Licensed to <Unknown...>"
Else
    lblOwner.Caption = "Licensed to " & Owner
End If

'activate timer and show command button
Tick = 0
timScroll.Enabled = True
cmdOk.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'remove the background from memory
Call DeleteBitmap(ScrollBmp.hDcMemory, ScrollBmp.hDcBitmap, ScrollBmp.hDcPointer)
Unload Me
End Sub

Private Sub timScroll_Timer()
'scroll the credits up wards.

'in nanoseconds
Const TimePerPixel = 50

Static BackArea As Rect
Static StartingTick As Long
Static StartingTop As Integer
Static LineHeight As Integer
Static TotalTextHeight As Integer
Static DisplayHeight As Integer
Static TextFont As FontStruc

Dim Counter As Integer
Dim StartNum As Integer
Dim FinishNum As Integer
Dim Result As Long
Dim TempBmp As BitmapStruc

'if just starting then
If Tick = 0 Then
    'set the co-ordinates of the background
    'picText.Cls
    
    BackArea.Top = 0
    BackArea.Left = 0
    BackArea.Right = (picText.ScaleWidth / Screen.TwipsPerPixelX)
    BackArea.Bottom = (picText.ScaleHeight / Screen.TwipsPerPixelY)
    
    DoEvents
    
    Tick = GetTickCount
End If

'if X nanoseconds have elapsed, move text up one pixel
If (Tick + TimePerPixel - StartingTick) < GetTickCount Then
    DoEvents
    
    If LineHeight = 0 Then
        'set initial values
        LineHeight = picText.TextHeight("I")
        DisplayHeight = picText.Height
        StartingTop = DisplayHeight
        TotalTextHeight = CInt(LineHeight) * GetLineCount(ScrollText)
        
        'set font details
        TextFont.Alignment = vbCentreAlign
        TextFont.Bold = picText.FontBold
        TextFont.Colour = picText.ForeColor
        TextFont.Italic = picText.FontItalic
        TextFont.Name = picText.FontName
        TextFont.PointSize = picText.FontSize
        TextFont.StrikeThru = picText.FontStrikethru
        TextFont.Underline = picText.FontUnderline
        
        'copy background
        picText.Cls
        Result = BitBlt(ScrollBmp.hDcMemory, 0, 0, ScrollBmp.Area.Right, ScrollBmp.Area.Bottom, picText.hDc, 0, 0, SRCCOPY)
    End If
    
    'create a temperory bitmap to draw the text on
    TempBmp.Area = ScrollBmp.Area
    Call CreateNewBitmap(TempBmp.hDcMemory, TempBmp.hDcBitmap, TempBmp.hDcPointer, TempBmp.Area, frmAbout, picText.BackColor, InPixels)
    Result = BitBlt(TempBmp.hDcMemory, 0, 0, TempBmp.Area.Right, TempBmp.Area.Bottom, ScrollBmp.hDcMemory, 0, 0, SRCCOPY)
    
    'get the starting and finishing values for the For loop.
    If StartingTop < 0 Then
        StartNum = (Abs(StartingTop) / LineHeight) - 1
    Else
        StartNum = 0
    End If
    FinishNum = StartNum + (DisplayHeight / LineHeight) + 1
    
    'scroll the appropiate text
    For Counter = StartNum To FinishNum
        Call MakeText(TempBmp.hDcMemory, GetLine(Counter, ScrollText), ((StartingTop + (Counter * LineHeight)) / Screen.TwipsPerPixelY), 0, LineHeight / Screen.TwipsPerPixelY, TempBmp.Area.Right, TextFont, InPixels)
    Next Counter

    'copy the resulting bitmap to the screen
    Result = BitBlt(picText.hDc, 0, 0, TempBmp.Area.Right, TempBmp.Area.Bottom, TempBmp.hDcMemory, 0, 0, SRCCOPY)
    
    'delete the temperory bitmap
    Call DeleteBitmap(TempBmp.hDcMemory, TempBmp.hDcBitmap, TempBmp.hDcPointer)
    
    StartingTop = StartingTop - Screen.TwipsPerPixelY
    If StartingTop < -(TotalTextHeight + LineHeight) Then
        'scroll text again
        StartingTop = DisplayHeight '+ LineHeight
    End If
    
    'StartingTick = GetTickCount
    Tick = GetTickCount
End If
End Sub

Private Function Round(ByVal Number As Single, Nearist As Single) As Single
'rounds the number off to the nearist number
'specified by Nearist. eg, if nearist is 5  and
'Number is 37 then it is rounded off to 35. If
'Number was 37.5 the number is rounded off to 40.

Dim BeforeVal As Single
Dim AfterVal As Single
Dim Between As Single

BeforeVal = Nearist * (Number \ Nearist)
AfterVal = BeforeVal + Nearist

Between = (BeforeVal + AfterVal) / 2

If Number >= Between Then
    'round to next highest number
    Round = AfterVal
Else
    'round to lowest number
    Round = BeforeVal
End If
End Function
