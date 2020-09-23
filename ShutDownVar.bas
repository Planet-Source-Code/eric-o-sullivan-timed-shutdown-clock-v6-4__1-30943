Attribute VB_Name = "ClockCode"
'This module contains the general setting and
'code for the main program. It also contains the
'global variables used by the various screens.
'
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================

Option Explicit             'all variables MUST be declared
Option Base 1               'all array elements start at "0"
Option Private Module   'This module is not to be accessed by any
                                    'other programs externally.

Private Type ShutDown
    Day As Integer
    CloseHour As Integer
    CloseMin As Integer
    CloseSec As Integer
    DelayTime As Integer
    DelayOn As Boolean
    ShutWin As Boolean
End Type


'the center points of the hands on the clock
Public Const Centre = 960
Public Const CentreDot = 910

Public Const Ask = "Ask"
Public Const Change = "Change"

Public Const FileName = "SetClock.ini"

Public Const DontLoadPic = 1

'vbSunday         1   Sunday (default)
'vbMonday        2   Monday
'vbTuesday       3   Tuesday
'vbWednesday  4   Wednesday
'vbThursday     5   Thursday
'vbFriday         6   Friday
'vbSaturday     7   Saturday

Public Week(7) As ShutDown

'used with FileName to determine the path of the
'.ini file
Public Method As String     'shut down method
Public FilePath As String   '.ini path

'background details
Public BackPath As String
Public StretchTile As String    'Stretch / Tile / None
Public BackOnOff As Boolean

'the machines registered owner
Public Owner As String

'password variables
Public Password As String
Public PassActive As Boolean
Public AskOrChange As String
Public CorrectPass As Boolean

'flags to prevent multiple file access.
Public Loading As Boolean
Public Saving As Boolean
Public Searching As Boolean

'Run program at startup/start minimized
Public StartUp As Boolean
Public StartMin As Boolean

'the height of the taskbar (no matter what resolution
'is displayed).
Public TaskBarHeight As Integer

'window positioning
Public SnapOn As Boolean
Public LastPos As PointAPI

'Is the form frmHandsClk permanently on top.
Public IsOnTop As Boolean

'the dimensions of the Time label
Public LabelRect As Rect

'see if the clock is still loading during startup
Public StillLoading As Boolean

'the amount of time the computer has been idle for
Public ComputerIdleTime As String
Public PSTime As String  'Predicted Shutdown Time (for idle shutdown)

'whether or not the program should shut the computer down
'after a specified amount of idle time
Public IdleShut As Boolean
Public IdleTimeInSec As Long

Public PreventShut As Boolean  'stop other apps shutting down windows

'the background for the digital time
Public BmpTime As BitmapStruc

'whether or not the analogue clock is being shown
Public AnaOn As Boolean

Public Sub Main()
'This is where the program first loads. Settings are loaded, values are
'set etc.

Dim Path As String
Dim Response As Integer
Dim ErrorNum As Integer
Dim Result As Long

On Error GoTo MyErrHandler

'----------
'make variable settings and do all possible preliminary checks
'without accessing the form until nececcary.

'if a previous instance of the program is active
'then stop the loading of this instance.
'Note : this only works in the load event.
If App.PrevInstance Then
    End
End If

'default setting for the registry until the data
'is loaded from the .ini file.
StartUp = False

'calculate the height of the taskbar (no matter
'what resolution is set).
TaskBarHeight = (Screen.TwipsPerPixelX * 28)

'set password to nil
CorrectPass = False

'set file access flags to false
Loading = False
Saving = False

'determine the .ini filepath
'this is in this procedure aswell for error handling
'eg, file missing/deleted during execution
FilePath = AddFile(App.Path, FileName)

If Owner = "" Then
    Owner = GetOwnerInReg
End If

'trap missing file - .exe bug fixed - 18/1/00
On Error Resume Next

Open Path For Input As #20
    ErrorNum = Err
Close #20

'GetAttr() will cause run-time error 5 if file
'doesn't exist. This will NOT happen during design-
'time.
If ErrorNum = 0 Then
    If (GetAttr(Path) Mod 2) <> 0 Then
        Response = MsgBox("Warning! Settings and colours can't be" & Chr(13) & Chr(10) & "saved because SetClock.ini file is read-only", vbOKOnly + vbExclamation, "File Access Error")
    End If
End If
On Error GoTo MyErrHandler

'-----------
'now that all variables etc. have been set, load and apply the settings
'to the form.

'Load the main form
Load frmShut
Load frmHandsClk

'show that the clock is loading
Call frmHandsClk.LoadIcon

Call frmHandsClk.CheckStatus '(DontLoadPic)     '** exe error start ** (Fixed)

frmHandsClk.lblShowTime.Caption = ""

Call frmHandsClk.SetColour   'see procedure

'set position of clock to bottom-right
Call HideShow(True)
Call MoveClock
Call TrayToTitle(frmHandsClk)

'load timed shut down form
Load frmShut

frmHandsClk.Show
DoEvents

Call SetTimeDimensions

'exit Main
Exit Sub

MyErrHandler:
    ' An error occurred.
    
    'If Err.Number = 0 Then
    '    Resume Next
    'End If
    
    ' Display the error.
    MsgBox ("Error: " & Err.Number _
            & " " & Err.Description _
            & " " & Err.Source)
    ' Continue running the program.
    Resume Next
End Sub

Public Sub GetTimeBackground()
'This will get a snapshot of the label area and pu the result
'into the time background picture.

Dim Result As Long

If frmHandsClk.Visible Then
    frmHandsClk.Cls
    Call frmHandsClk.DrawDots

    'copy the appropiate section of the picture onto the bitmap
    'create a new bitmap for the digital values
    Call DeleteBitmap(BmpTime.hDcMemory, BmpTime.hDcBitmap, BmpTime.hDcPointer)
    Call CreateNewBitmap(BmpTime.hDcMemory, BmpTime.hDcBitmap, BmpTime.hDcPointer, BmpTime.Area, frmHandsClk, 0, InPixels)
    
    DoEvents
    Result = BitBlt(BmpTime.hDcMemory, 0, 0, (BmpTime.Area.Right - BmpTime.Area.Left), (BmpTime.Area.Bottom - BmpTime.Area.Top), frmHandsClk.hDc, 0, BmpTime.Area.Top, SRCCOPY) 'BmpTime.Area.Left, BmpTime.Area.Right, SRCCOPY)
End If
End Sub

Public Function AskForPassword() As Boolean
'this will determine whether or not to ask for a
'password.

If (PassActive) And (Not CorrectPass) Then
    'ask for password
    AskOrChange = Ask
    Load frmPass
    frmPass.Show
    AskForPassword = True
Else
    'no password needed
    AskForPassword = False
End If
End Function

Public Function GetBefore(Sentence As String) As String
'This procedure returns all the character of a
'string before the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Before As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If (Counter <> Len(Sentence)) And (Counter <> 0) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If

GetBefore = Before
End Function


Public Function GetAfter(Sentence As String) As String
'This procedure returns all the character of a
'string after the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Rest As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If

GetAfter = Rest
End Function

Public Sub HideShow(Optional ByVal StartUp As Boolean = False)
'This procedure will resize the form so that it can either show the
'clock with or without the analogue part.

Dim AnaHeight As Integer
Dim OldHeight As Integer
Dim Gap As Integer

If StillLoading Then
    Exit Sub
End If

'calculate the height of one pixel in twips
Gap = -1 * (Screen.TwipsPerPixelY)

'move digital clock up or down
With frmHandsClk
    If (Not AnaOn) Then
        'if "Dont Show analogue" then;
        
        'if the analogue clock is not visible anyway, then exit
        If Not .lblShowHands.Visible Then
            Exit Sub
        End If
            
        .lblShowTime.Top = 0
        .lblShowDay.Top = .lblShowTime.Height + Gap
        .lblShowDate.Top = .lblShowDay.Top + .lblShowDay.Height + Gap
        
        'hide and disable the analogue clock
        .lblShowHands.Visible = False
        .lnSecond.Visible = False
        .lnMinute.Visible = False
        .lnHour.Visible = False
        .timHand.Enabled = False
        DoEvents
        
        'resize the clock
        OldHeight = frmHandsClk.Height
        .Height = (.lblShowDate.Top + .lblShowDate.Height + Gap) + (.Height - .ScaleHeight)
        
        'move the clock
        .Top = .Top + (OldHeight - ((.lblShowTime.Height * 3) + (Gap * 3))) - (.Height - .ScaleHeight)
        
        If Not StartUp Then
            Call SetTimeDimensions
        End If
    Else
        'else "Show analogue"
        
        'if the analogue clock is visible anyway, then exit
        If .lblShowHands.Visible Then
            Exit Sub
        End If
    
        .lblShowTime.Top = .lblShowHands.Height + Gap
        .lblShowDay.Top = .lblShowTime.Top + .lblShowTime.Height + Gap
        .lblShowDate.Top = .lblShowDay.Top + .lblShowDay.Height + Gap
        
        'show and enable the analogue clock
        .lblShowHands.Visible = True
        .lnSecond.Visible = True
        .lnMinute.Visible = True
        .lnHour.Visible = True
        .timHand.Enabled = True
        
        'move the clock
        AnaHeight = (frmHandsClk.lblShowTime.Height + frmHandsClk.lblShowDay.Height + frmHandsClk.lblShowDate.Height)
        frmHandsClk.Top = frmHandsClk.Top - (AnaHeight + frmHandsClk.lblShowHands.Height) + AnaHeight '- (Gap * 0)
        frmHandsClk.Height = frmHandsClk.lblShowHands.Height + (frmHandsClk.Height - frmHandsClk.ScaleHeight) + (frmHandsClk.lblShowTime.Height * 3) + (Gap * 3)
    
        'Call SetTimeDimensions
        Call LoadPictureOntoForm(frmHandsClk, StartUp)
    End If
    
    'reset clock position
    '.Cls
End With
End Sub

Public Sub LoadPictureOntoForm(Form As Form, Optional ByVal StartUp As Boolean = False)
'Paint the image (automatically stretching or
'tiling the image) onto the form. Painting the
'image reduces the flicker normally caused when
'one control passes over another. (Previous
'versions used an image control to stretch/tile
'which resulted in a flicker when the hands
'moved over it).

Const X = 0
Const Y = 1

Dim TileX As Integer
Dim TileY As Integer
Dim SourceWidth As Single
Dim SourceHeight As Single
Dim ErrNum As Long
Dim TimeHeight As Long
Dim TimeTop As Long
Dim TimeLeft As Long
Dim TimeWidth As Long
Dim ShowAnalogue As Boolean
Dim Width As Integer
Dim Height As Integer
Dim Gap(2) As Integer
Dim Result As Long

'prevent any errors from stopping execution
On Error Resume Next

Call LockWindow(Form)

Form.AutoRedraw = True

'set details and form sizes in case the analogue is not checked
ShowAnalogue = Form.mnuFileAna.Checked
Gap(X) = Screen.TwipsPerPixelX
Gap(Y) = Screen.TwipsPerPixelY
Width = Form.ScaleWidth
Height = (Form.lblShowTime.Height + Form.lblShowDay.Height + Form.lblShowDate.Height + Form.lblShowHands.Height) + (Gap(Y) * 3)

'the picture needs to be re-set each time.
Form.Picture = LoadPicture
ErrNum = Err

Select Case LCase(StretchTile)
Case "[none]"
    'already re-set
Case "tile"
    'tile the image, column by column
    SourceWidth = Form.imgLogo.Width
    SourceHeight = Form.imgLogo.Height
    For TileX = 0 To Int(Width / SourceWidth)
        For TileY = 0 To Int(Height / SourceHeight)
            'paint image
            Form.PaintPicture Form.imgLogo.Picture, TileX * SourceWidth, TileY * SourceHeight, SourceWidth, SourceHeight, 0, 0
        Next TileY
    Next TileX

Case "stretch"
    'stretch image
    TimeTop = 0
    TimeLeft = 0
    TimeHeight = Width
    TimeWidth = Height
    Form.PaintPicture Form.imgLogo.Picture, 0, 0, Width, Height, 0, 0, Form.imgLogo.Width, Form.imgLogo.Height

    ErrNum = Err
End Select

'this has to be set back to False because the
'program uses PSet has would cause a very large
'flicker.
Form.AutoRedraw = False

Call UnLockWindow

'Call GetTimeBackground
If Not StartUp Then
    Call SetTimeDimensions
End If

'only draw the dots if the appropiate option is selected
If ShowAnalogue And (Not StartUp) Then
    Call frmHandsClk.DrawDots
End If
End Sub

Public Sub DoShutMethod(Method As String)
'shut down the computer according to what method
'is specified.


'Shut Down
'Power Down
'Force Close
'Restart
'Log Off

'EWX_LOGOFF = 0
'EWX_SHUTDOWN = 1
'EWX_REBOOT = 2
'EWX_FORCE = 4
'EWX_POWEROFF = 8

Dim RetVal As Integer

Select Case LCase(Method)
Case "shut down"
    Call WINShutdown
Case "restart"
    Call WINReboot
Case "force close"
    Call WINForceClose
Case "log off"
    Call WINLogUserOff
Case "power down"
    WINPowerDown
End Select
End Sub

Public Function FindPos(Text As String) As Integer
'find the position if "=" in a string and return it
Const Mychar = ","
Dim Counter As Integer

For Counter = 1 To Len(Text)
    If Mid(Text, Counter, 1) = Mychar Then
        FindPos = Counter
        Exit For
    End If
Next Counter
End Function

Public Function GetNum(AsciiNum As Integer) As Integer
'only numeric characters allowed
'8=Del ; 9=Tab ; 48-57="0" to "9"
Select Case AsciiNum
Case 8, 9, 48 To 57
Case Else
    'no character
    AsciiNum = 0
End Select

GetNum = AsciiNum
End Function

Public Function GetAscii(Character As String) As Integer
'get the ascii number of a numeric charcter
Dim AsciiNum As Integer

Character = Left(Character, 1)
AsciiNum = Val(Character)
AsciiNum = AsciiNum + 48
GetAscii = AsciiNum
End Function

Public Sub PutOnTop()
'This will put the form frmHandsClk on top of
'other program windows or not.

If IsOnTop Then
    Call StayOnTop(frmHandsClk)
Else
    Call NotOnTop(frmHandsClk)
End If

frmHandsClk.mnuFileAdvOnTop.Checked = IsOnTop

'save new setting
Call frmHandsClk.SaveStatus
End Sub

Public Sub MoveClock()
'This procdure moves the clock to the far right of
'the screen, just above the taskbar regardless of
'resolution.

Dim WorkArea As Rect

WorkArea = GetWorkArea

'horizontal
If (LastPos.X + frmHandsClk.Width) > Screen.Width Then
    'snap window to the right side of the screen
    LastPos.X = (WorkArea.Right * Screen.TwipsPerPixelX) - frmHandsClk.Width
Else
    If LastPos.X < 0 Then
        'snap window to the left
        LastPos.X = (WorkArea.Left * Screen.TwipsPerPixelX)
    End If
End If

'vertical
If (LastPos.Y + frmHandsClk.Height) > Screen.Height Then
    'snap window to the bottom
    LastPos.Y = (WorkArea.Bottom * Screen.TwipsPerPixelY) - frmHandsClk.Height
Else
    If LastPos.Y < 0 Then
        'snap window to the top
        LastPos.Y = (WorkArea.Top * Screen.TwipsPerPixelY)
    End If
End If

If StillLoading Then
    frmHandsClk.Visible = False
End If

'set the clock position
Call LockWindow(frmHandsClk)
frmHandsClk.Left = LastPos.X
frmHandsClk.Top = LastPos.Y
'frmHandsClk.Visible = True
Call UnLockWindow

End Sub

Public Sub PutDotsOnForm(ColDot As Long)    'not used in 6.2
'This procedure puts the dots onto the frmHandsClk form and
'displays them. The procedure uses the api calls of the DrawTextmod
'module to produce an off-screen bitmap before blitting the bitmap
'back onto the form, along with the dots.

Dim DotPtr As Long
Dim DotMem As Long
Dim DotBmp As Long
Dim Junk As Long
Dim Result As Long

Dim XCo As Integer
Dim YCo As Integer
Dim Width As Long
Dim Height As Long
Dim Counter As Integer

'convert these to pixels
Width = (frmHandsClk.Width / Screen.TwipsPerPixelX)
Height = (frmHandsClk.lblShowHands.Height / Screen.TwipsPerPixelY)

'create a bitmap to draw dots on
DotMem = CreateCompatibleDC(frmHandsClk.hDc)
DotBmp = CreateCompatibleBitmap(frmHandsClk.hDc, Width, Height)
DotPtr = SelectObject(DotMem, DotBmp)

'hide the clock hands
frmHandsClk.lnHour.Visible = False
frmHandsClk.lnMinute.Visible = False
frmHandsClk.lnSecond.Visible = False

'take a snap shot of the analogue area
frmHandsClk.Cls
DoEvents
Result = BitBlt(DotMem, 0, 0, Width, Height, frmHandsClk.hDc, 0, 0, SRCCOPY)

'draw the dots onto the off-screen bitmap
'------------------
'draw small dots
For Counter = 0 To 360 Step 6
    XCo = ((CentreDot + ((50 + Cos(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelX)
    YCo = ((CentreDot + ((50 + Sin(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelY)
    Call DrawRect(DotMem, ColDot, XCo, YCo, XCo + 1, YCo + 1)
Next Counter
    
'draw big dots
For Counter = 0 To 360 Step 30
    XCo = ((CentreDot + ((50 + Cos(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelX)
    YCo = ((CentreDot + ((50 + Sin(Counter * 3.14 / 180) * 900))) / Screen.TwipsPerPixelY)
    Call DrawRect(DotMem, ColDot, XCo, YCo, XCo + 2, YCo + 2)
Next Counter
'------------------

'put the bitmap onto the form
Result = BitBlt(frmHandsClk.hDc, 0, 0, Width, Height, DotMem, 0, 0, SRCCOPY)

'show the clock hands again
frmHandsClk.lnHour.Visible = True
frmHandsClk.lnMinute.Visible = True
frmHandsClk.lnSecond.Visible = True

'delete the off-screen bitmap
Junk = SelectObject(DotMem, DotPtr)
Junk = DeleteObject(DotBmp)
Junk = DeleteDC(DotMem)
End Sub

Public Function GetLine(ByVal LineNum As Long, Line As String) As String
'This function will return a line that is LineNum distance from the start.
'eg, if LineNum = 5 then the function will return the fifth line. If there
'is not 5 lines in the text, then the function will return blank.

Dim Counter As Integer
Dim Temp As Long
Dim Index As Long
Dim LastPos As Integer
Dim NextPos As Integer
Dim Captured As String
Dim NumOfLines As Integer

NumOfLines = GetLineCount(Line)
If NumOfLines < LineNum Then
    'return blank if there aren't that many lines in the string
    GetLine = ""
    Exit Function
End If

For Counter = 0 To NumOfLines
        LastPos = -1
        
        'get the starting position
        For Index = 1 To LineNum
            LastPos = InStr(LastPos + 2, Line, vbCrLf)
        Next Index
        
        'the starting position of the string was found, find the finishing
        'position.
        NextPos = InStr(LastPos + 2, Line, vbCrLf)
        If NextPos > 0 Then
            If (LineNum = 0) Then
                'the line is at the left of the string
                Captured = Left(Line, NextPos - 1)
            Else
                If LastPos = 0 Then
                    'line number not found
                    Captured = ""
                Else
                    'line is in the middle of the string
                    Captured = Mid(Line, LastPos + 2, NextPos - LastPos - 2)
                End If
            End If
        Else
            'line is at the end of the string
            Captured = Right(Line, Len(Line) - LastPos - 1)
        End If
        
        'stop searching
        Exit For
Next Counter

'return result
GetLine = Captured
'If LineNum = 22 Then Stop
End Function

Public Function GetLineCount(ByVal Text As String) As Integer
'This function will return the number of lines in the text

Dim Counter As Integer
Dim LastPos As Long

If Text = "" Then
    'there are no lines in a blank string
    GetLineCount = 0
    Exit Function
End If

LastPos = -1
Counter = 0
Do While LastPos <> 0
    Counter = Counter + 1
    LastPos = InStr(LastPos + Len(vbCrLf), Text, vbCrLf)
Loop  'LastPos will =0 when InStr cannot find any more occurances of vbCrlf

GetLineCount = Counter
End Function

Public Sub SetTimeDimensions()
'this procedure will set the size and position of where to displaye the
'time in relation to whether or not the analogue clock is displayed.

'set the label dimensions
LabelRect.Left = frmHandsClk.lblShowTime.Left / Screen.TwipsPerPixelX
LabelRect.Top = frmHandsClk.lblShowTime.Top / Screen.TwipsPerPixelY
LabelRect.Bottom = (frmHandsClk.lblShowTime.Top + frmHandsClk.lblShowTime.Height) / Screen.TwipsPerPixelY
LabelRect.Right = (frmHandsClk.lblShowTime.Left + frmHandsClk.lblShowTime.Width) / Screen.TwipsPerPixelX

'store them and get a screenshot of the clock area.
BmpTime.Area = LabelRect
Call GetTimeBackground
End Sub
