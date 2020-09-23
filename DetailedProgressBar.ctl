VERSION 5.00
Begin VB.UserControl ProgressBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   EditAtDesignTime=   -1  'True
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   HasDC           =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4530
   ToolboxBitmap   =   "DetailedProgressBar.ctx":0000
   Begin VB.Label lblEvent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Event Pickup Label"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Shape shpValue 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      DrawMode        =   8  'Xor Pen
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50% Complete"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1650
      TabIndex        =   0
      Top             =   480
      Width           =   1035
   End
   Begin VB.Shape shpContainer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'All the events, properties and procedures are
'listed alphabetically. I got most of what I
'wanted from the sample "CtlPlus", but the actual
'code is mine. The warnings are from the sample
'too (I thought it best to leave them in). Enjoy,
'Eric O'Sullivan

'Note :  to change the text displayed in the
'description part of the properties box, use the
'Procedure Attributes in the Tools menu.


'===============================================
'               Public Declarations ;)
'===============================================

Dim MaxVal As Double
Dim MinVal As Double
Dim ProgValue As Double
Dim ProgEnabled As Boolean
Dim CurrentText As String

'create new constant collection (see BorderStyle)
Public Enum BorderConstants
    Normal = 0
    Fixed_Single = 1
End Enum

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

'===============================================
'                  Event Settings
'===============================================

Private Sub lblEvent_Click()
RaiseEvent Click
End Sub

Private Sub lblEvent_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lblEvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, _
    ScaleX(X + lblEvent.Left, vbTwips, vbContainerPosition), _
    ScaleY(Y + lblEvent.Height, vbTwips, vbContainerPosition))
End Sub

Private Sub lblEvent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, _
    ScaleX(X + lblEvent.Left, vbTwips, vbContainerPosition), _
    ScaleY(Y + lblEvent.Height, vbTwips, vbContainerPosition))
End Sub

Private Sub lblEvent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'the ultimate copy+paste job |^| ! :)
RaiseEvent MouseUp(Button, Shift, _
    ScaleX(X + lblEvent.Left, vbTwips, vbContainerPosition), _
    ScaleY(Y + lblEvent.Height, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_InitProperties()
MaxVal = 100
MinVal = 0
ProgValue = 50
ProgEnabled = True
CurrentText = UserControl.Name 'ProgValue & "% Complete"
Caption = UserControl.Name
SideStyle = UserControl.BorderStyle
SetValues
End Sub

Private Sub UserControl_Resize()
SetValues
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'load custom property values from storage. You need
'to load the properties because otherwise the
'user has to set all the properties each time the
'control is loaded (ie, when the parent form is
'loaded).

lblDisplay.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
ProgValue = PropBag.ReadProperty("Value", 50)
MinVal = PropBag.ReadProperty("Min", 0)
MaxVal = PropBag.ReadProperty("Max", 100)
shpValue.FillColor = PropBag.ReadProperty("FillColor", &H80FFFF)
shpContainer.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
shpContainer.BorderColor = PropBag.ReadProperty("BorderColor", &H0&)
lblDisplay.ForeColor = PropBag.ReadProperty("ForeColor", &H0&)
ProgEnabled = UserControl.Enabled

Call SetValues
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
'save any changes made to the properties before
'the form/control is unloaded. Very useful

'propbag.write([property],[value],[default])
Call PropBag.WriteProperty("Caption", lblDisplay.Caption, UserControl.Name)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
Call PropBag.WriteProperty("Value", ProgValue, 50)
Call PropBag.WriteProperty("Min", MinVal, 0)
Call PropBag.WriteProperty("Max", MaxVal, 100)
Call PropBag.WriteProperty("BackColor", shpContainer.BackColor, &HFFFFFF)
Call PropBag.WriteProperty("BorderColor", shpContainer.BorderColor, &H0&)
Call PropBag.WriteProperty("FillColor", shpValue.FillColor, &H80FFFF)
Call PropBag.WriteProperty("ForeColor", lblDisplay.ForeColor, &H0&)
Call PropBag.WriteProperty("Enabled", ProgEnabled, True)

End Sub

'===============================================
'                Property Settings
'===============================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returnes or sets the background colour of the progress bar"
    BackColor = shpContainer.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    shpContainer.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returnes or sets the border colour"
    BorderColor = shpContainer.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpContainer.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Let BorderStyle(ByVal Style As BorderConstants) 'BorderStyleConstants)
Attribute BorderStyle.VB_Description = "Returnes or sets the border style. Normal (default) or fixed single."
Const None = 0
Const FixedSingle = 1

'I have to supply my own validation constants
Select Case Style
Case None
Case FixedSingle
Case Else
    Exit Property
End Select

UserControl.BorderStyle = Style

PropertyChanged "BorderStyle"
End Property

Public Property Get BorderStyle() As BorderConstants
BorderStyle = UserControl.BorderStyle
End Property

Public Property Let Caption(ByVal Display As String)
Attribute Caption.VB_Description = "Returnes or sets the text that is displayed in the control"
lblDisplay.Caption = Display
CurrentText = Display

PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
Caption = lblDisplay.Caption
End Property

Public Property Let Enabled(ByVal SeeMe As Boolean)
Attribute Enabled.VB_Description = "Enable or disable the control."
Static LastBackColor As Long
Static LastFillColor As Long
Static LastBorderColor As Long

ProgEnabled = SeeMe
lblDisplay.Enabled = SeeMe
UserControl.Enabled = SeeMe

If (LastBackColor = 0) And (LastFillColor = 0) And (LastBorderColor = 0) Then
    LastBackColor = BackColor
    LastFillColor = FillColor
    LastBorderColor = BorderColor
End If

If SeeMe Then
    shpValue.FillColor = LastFillColor
    shpContainer.BackColor = LastBackColor
    shpContainer.BorderColor = LastBorderColor
Else
    'save colors before changing them
    LastBackColor = BackColor
    LastFillColor = FillColor
    LastBorderColor = BorderColor
    
    FillColor = &H80000018 'blue-gray
    BackColor = &HC0C0C0 'light gray
    BorderColor = &H808080 'dark gray
End If
End Property

Public Property Get Enabled() As Boolean
Enabled = ProgEnabled
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returnes or sets the colour the control fills with. Note, the color used is the XOR Pen of the background aswell as the fill colour. See XOR Pen for more details."
    FillColor = shpValue.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    shpValue.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns or sets the font"
    Set Font = lblDisplay.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblDisplay.Font = New_Font
    PropertyChanged "Font"
    ' Manually added: Changing the font
    '   may require adjusting the position
    '   of the Label control.
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns or sets the foreground colour"
    ForeColor = lblDisplay.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblDisplay.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Let Max(ByVal MaxNum As Double)
Attribute Max.VB_Description = "Returns or set the maximum value of the progress bar. It can hold any number up to a Double."
If MaxNum > MinVal Then
    MaxVal = MaxNum
    
    'make sure value is between ranges.
    Value = ProgValue
End If
End Property

Public Property Get Max() As Double
Max = MaxVal

'change display
Value = ProgValue
End Property

Public Property Let Min(ByVal MinNum As Double)
Attribute Min.VB_Description = "Returnes or sets the minimum value of the progress bar. It can hold any number up to a Double."
If MinNum < MaxVal Then
    MinVal = MinNum

    'make sure value is between ranges.
    Value = ProgValue
End If
End Property

Public Property Get Min() As Double
Min = MinVal
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
UserControl.Refresh
End Sub

Public Property Let Value(ByVal Total As Double)
Attribute Value.VB_Description = "The current value that the progress bar will fill up to. It can hold any number up to a Double. :)"
'move progress bar to new value

Dim Twips As Integer

Dim TopVal As Single
Dim Difference As Single
Dim WidthMax As Long
Dim NewWidth As Double

Const Start = 0

'set the number of twips per pixel
Twips = Screen.TwipsPerPixelX

If Total > MaxVal Then
    Total = MaxVal
End If
    
If Total < MinVal Then
    Total = MinVal
End If

If (Total >= MinVal) And (Total <= MaxVal) Then
    'bring the starting value to zero and record
    'the difference and add it to maxval, thus
    'keeping the scale intact. *smug bastard*
    
    ProgValue = Total
    
    Select Case MinVal
    Case Is < 0
        Difference = Posit(MinVal)
        TopVal = MaxVal + Difference
        Total = Total + Difference
    Case Is > 0
        TopVal = MaxVal - MinVal
        Total = Total - MinVal
    Case Else
        TopVal = MaxVal
    End Select
    
    'get the end point of the control, depending on
    'the border sytle
    Select Case UserControl.BorderStyle
    Case 0
        'none
        WidthMax = shpContainer.Width - 20
    Case 1
        'Fixed Single
        WidthMax = shpContainer.Width
    End Select
    
    'prevent division by zero
    If TopVal = 0 Then Exit Property
    
    NewWidth = Int((WidthMax / TopVal) * Total)
    'there are roughly 5 twips per pixel, so I only
    'change the value if it is one of these values.
    If TopVal <> 0 Then
        'no point setting the value if it's already
        'set. Also reduces flicker when changing
        'the value during long operations.
        If (Round(NewWidth, Twips) <> shpValue.Width) Then
            shpValue.Width = Round(NewWidth, Twips)
        End If
    Else
        shpValue.Width = 0
    End If
End If
End Property

Public Property Get Value() As Double
Value = ProgValue
End Property

'===============================================
'                   Procedures
'===============================================

Private Function Posit(Number As Double) As Double
Attribute Posit.VB_Description = "Hidden and not to be used by the user."
Attribute Posit.VB_MemberFlags = "40"
'returns the positive value of a number
Posit = Sqr(Number ^ 2)
End Function

Private Sub SetValues()
'basically all the initilization stuff like
'at the start or resizing.
shpContainer.Top = 0
shpContainer.Left = 0

'if control height is greater than text size then

If BorderStyle = 0 Then
    '0= no border style
    'don't ask, it has to do with the border
    'being chopped off at the right side.
    shpContainer.Width = UserControl.Width
    shpContainer.Height = UserControl.Height
    
    'set the progress bar size
    shpValue.Left = shpContainer.Left + 10
    shpValue.Top = shpContainer.Top + 10
    shpValue.Height = shpContainer.Height - 10
Else
    '1= fixed single
    'move border inside the container
    shpContainer.Top = -10
    shpContainer.Left = -10
    shpContainer.Width = UserControl.Width - 20
    shpContainer.Height = UserControl.Height - 20
    
    'set the progress bar size
    shpValue.Left = shpContainer.Left
    shpValue.Top = shpContainer.Top
    shpValue.Height = shpContainer.Height

End If

'set event label size
lblEvent.Left = 0
lblEvent.Top = 0
lblEvent.Width = UserControl.Width
lblEvent.Height = UserControl.Height
lblEvent.Caption = ""


'verticaly center the caption
If shpContainer.Height > 150 Then
    lblDisplay.Top = (shpContainer.Top + (shpContainer.Height / 2)) - (lblDisplay.Height / 2)
Else
    lblDisplay.Top = shpContainer.Top
End If

If UserControl.BorderStyle = 1 Then
    'move caption up a bit
    lblDisplay.Top = lblDisplay.Top - 30
End If

'center the caption
lblDisplay.Left = (shpContainer.Left + (shpContainer.Width / 2)) - (lblDisplay.Width / 2)

'set values (see "UserControl_Initproperties()"
'and "UserControl_Resize()" )
Max = MaxVal
Min = MinVal
Value = ProgValue
Enabled = UserControl.Enabled
Text = UserControl.Name
End Sub

Private Function Round(ByVal Number As Single, Nearist As Integer) As Integer
'rounds the number off to the nearist number
'specified by Nearist. eg, if nearist is 5  and
'Number is 37 then it is rounded off to 35. If
'Number was 37.5 the number is rounded off to 40.

Dim BeforeVal As Integer
Dim AfterVal As Integer
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

