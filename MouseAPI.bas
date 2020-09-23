Attribute VB_Name = "MouseAPI"
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Sub MoveMouseTo(XCo As Integer, YCo As Integer, Optional Delay As Long, Optional Steps As Integer)
'This procedure moves the mouse to a specified point on the screen
'in a given number of steps

Const x = 0
Const y = 1
Const DefaultDelay = 1

Dim GotPos As POINTAPI
Dim Counter As Integer
Dim Jump As Integer
Dim Direction(2) As Integer
Dim Distance(2) As Integer
Dim Move(2) As Long
Dim Result As Long

If Delay = 0 Then
    Delay = DefaultDelay
End If

GotPos = MouseHere
XCo = XCo / Screen.TwipsPerPixelX
YCo = YCo / Screen.TwipsPerPixelY

Distance(x) = XCo - GotPos.x
Distance(y) = YCo - GotPos.y

'find the direction to move the mouse cursor in
For Counter = x To y
    If Distance(Counter) < 0 Then
        Direction(Counter) = -1
    Else
        Direction(Counter) = 1
    End If
Next Counter

Distance(x) = Abs(Distance(x))
Distance(y) = Abs(Distance(y))

If Steps <> 0 Then
    'the the number of steps was specified then, use them
    Jump = Steps
Else
    'The number of steps = the largest distance to move (by pixel)
    If Abs(Distance(x)) > Abs(Distance(y)) Then
        'the largest distance between the two points is horizontal
        Jump = Abs(Distance(x))
    Else
        'else the largest distance is vertical or is equal
        Jump = Abs(Distance(y))
    End If
End If

For Counter = 1 To Jump
    'move the mouse
    
    'move horizontal
    Move(x) = GotPos.x + (((Distance(x) / Jump) * Counter) * Direction(x))
    
    'move vertical
    Move(y) = GotPos.y + (((Distance(y) / Jump) * Counter) * Direction(y))
    
    'set the cursor position
    Result = SetCursorPos(Move(x), Move(y))
    
    Call PauseTime(Delay)
Next Counter
End Sub

Public Function MouseHere(Optional InTwips As Boolean = False) As POINTAPI
'This function returns the position of the mouse on the screen in
'twips if asked for, otherwise the result is returned in pixels.

Dim GotPos As POINTAPI
Dim Result As Long

Result = GetCursorPos(GotPos)

If InTwips Then
    'resturn result in twips if asked for, else in pixels
    GotPos.x = GotPos.x * Screen.TwipsPerPixelX
    GotPos.y = GotPos.y * Screen.TwipsPerPixelY
End If

MouseHere = GotPos
End Function

Public Sub SwapVal(ByRef Val1 As Integer, ByRef Val2 As Integer)
'This procedure will swap the two values around

Dim Temp As Integer

Temp = Val1
Val1 = Val2
Val2 = Temp
End Sub

Public Sub PauseTime(Delay As Long)
'This procedure will wait a specified number of milliseconds

Dim StartTick As Long

StartTick = GetTickCount

While GetTickCount < (StartTick + Delay)
    DoEvents
Wend
End Sub
