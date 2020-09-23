Attribute VB_Name = "TaskBar"
'This module provides information about the screen's work area
'and the task-bar.
'
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type Rect '
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum AlignmentConst
    vbLeft = 0
    vbRight = 1
    vbTop = 2
    vbBottom = 3
End Enum

Const SPI_GETWORKAREA As Long = 48
Const X = 0
Const Y = 1

Public Function GetWorkArea() As Rect
'Get the area the user is working with (minus the task bar)
'in PIXELS

Dim Result As Long
Dim WorkArea As Rect

Result = SystemParametersInfo(SPI_GETWORKAREA, 0&, WorkArea, 0&)
GetWorkArea = WorkArea
End Function

Public Function GetAlignment() As AlignmentConst
'Find the alignment of the taskbar

Dim WorkArea As Rect
Dim Align As AlignmentConst

WorkArea = GetWorkArea

If WorkArea.Left <> 0 Then
    'the taskbar MUST be right aligned
    Align = vbLeft
Else
    If WorkArea.Top <> 0 Then
        'The taskbar MUST be bottom aligned
        Align = vbTop
    Else
        If ((WorkArea.Bottom - WorkArea.Top) * Screen.TwipsPerPixelY) = Screen.Height Then
            'If the workarea height is equal to the screen height then
            'the taskbar MUST be left aligned
            Align = vbRight
        Else
            Align = vbBottom
        End If
    End If
End If

GetAlignment = Align
End Function

Public Function TaskBarDimensions() As Rect
'Find out what the taskbars', left, top, right and bottom values are
'in TWIPS

Dim WorkArea As Rect
Dim TaskBarDet As Rect
Dim TwipsPP(2) As Byte 'Twips Per Pixel

WorkArea = GetWorkArea
TwipsPP(X) = Screen.TwipsPerPixelX
TwipsPP(Y) = Screen.TwipsPerPixelY

'set the taskbars' default values to the screen size
TaskBarDet.Top = 0
TaskBarDet.Bottom = Screen.Height
TaskBarDet.Left = 0
TaskBarDet.Right = Screen.Width

'change the appropiate value according to alignment
Select Case GetAlignment
Case vbLeft
    TaskBarDet.Right = (WorkArea.Left * TwipsPP(X))
Case vbRight
    TaskBarDet.Left = (WorkArea.Right * TwipsPP(X))
Case vbTop
    TaskBarDet.Bottom = (WorkArea.Top * TwipsPP(Y))
Case vbBottom
    TaskBarDet.Top = (WorkArea.Bottom * TwipsPP(Y))
End Select

'return result
TaskBarDimensions = TaskBarDet
End Function

Public Sub SnapWindow(MyFrm As Form, Distance As Integer)
'This procedure will snap the window to the edges of the work area
'(like winamp) if the form is within a certain distance of the edges
'(measured in pixels).

Dim WorkArea As Rect
Dim DistTwip(2) As Long

If Distance < 1 Then
    'a value of zero is meaningless to this procedure and a value of
    'less than zero is invalid.
    Exit Sub
End If

'find out if the edge of the window is within snapping distance
WorkArea = GetWorkArea
DistTwip(X) = Screen.TwipsPerPixelX
DistTwip(Y) = Screen.TwipsPerPixelY

If WithinDistance((MyFrm.Top / DistTwip(Y)), WorkArea.Top, Distance) Then
    'snap window to the top
    MyFrm.Top = WorkArea.Top * DistTwip(Y)
End If

If WithinDistance((MyFrm.Left / DistTwip(X)), WorkArea.Left, Distance) Then
    'snap window to the left
    MyFrm.Left = WorkArea.Left * DistTwip(X)
End If

If WithinDistance(((MyFrm.Top + MyFrm.Height) / DistTwip(Y)), WorkArea.Bottom, Distance) Then
    'snap window to the bottom
    MyFrm.Top = (WorkArea.Bottom * DistTwip(Y)) - MyFrm.Height
End If

If WithinDistance(((MyFrm.Left + MyFrm.Width) / DistTwip(X)), WorkArea.Right, Distance) Then
    'snap window to the right
    MyFrm.Left = (WorkArea.Right * DistTwip(X)) - MyFrm.Width
End If
End Sub

Public Function WithinDistance(Value As Long, Edge As Long, ByVal Distance As Long) As Boolean
'Find out if the value is within the distance of the edge
If (Value > (Edge - Distance)) And (Value < (Edge + Distance)) Then
    WithinDistance = True
Else
    WithinDistance = False
End If
End Function

Public Sub CheckIfOutsideScreen(Form As Form)
'This moves a form inside the work area of the screen if the
'form is outside the work area.

Dim WorkArea As Rect
Dim Twip(2) As Integer

'if the form is minimized or maximized then don't do this
'- it will generate an error otherwise.
If (Form.WindowState = vbMinimized) Or (Form.WindowState = vbMaximized) Then
    Exit Sub
End If

WorkArea = GetWorkArea

'convert workarea to twips
Twip(X) = Screen.TwipsPerPixelX
Twip(Y) = Screen.TwipsPerPixelY

WorkArea.Top = WorkArea.Top * Twip(Y)
WorkArea.Bottom = WorkArea.Bottom * Twip(Y)
WorkArea.Left = WorkArea.Left * Twip(X)
WorkArea.Right = WorkArea.Right * Twip(X)

'horizontal
If (Form.Left + Form.Width) > WorkArea.Right Then
    Form.Left = WorkArea.Right - Form.Width
End If

If Form.Left < WorkArea.Left Then
    Form.Left = WorkArea.Left
End If

'vertical
If (Form.Top + Form.Height) > WorkArea.Bottom Then
    Form.Top = WorkArea.Bottom - Form.Height
End If

If Form.Top < WorkArea.Top Then
    Form.Top = WorkArea.Top
End If
End Sub

Public Sub CentreForm(FormName As Form)
'This procedure will centre the form in the work area
'The form parameter "StartUpPosition = CenterScreen" does not
'centre the form in the work area, ie it will not take into account
'the position/height of the taskbar when positioning the form

Dim WorkArea As Rect

WorkArea = AreaToTwips(GetWorkArea)

FormName.Left = ((WorkArea.Right - FormName.Width) / 2) - WorkArea.Left
FormName.Top = ((WorkArea.Bottom - FormName.Height) / 2) - WorkArea.Top
End Sub

Public Function AreaToTwips(WorkArea As Rect) As Rect
'This function will convert a rect structure to twips

WorkArea.Left = WorkArea.Left * Screen.TwipsPerPixelX
WorkArea.Right = WorkArea.Right * Screen.TwipsPerPixelX
WorkArea.Top = WorkArea.Top * Screen.TwipsPerPixelY
WorkArea.Bottom = WorkArea.Bottom * Screen.TwipsPerPixelY

AreaToTwips = WorkArea
End Function

