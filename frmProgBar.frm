VERSION 5.00
Begin VB.Form frmProgBar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress Bar"
   ClientHeight    =   375
   ClientLeft      =   3675
   ClientTop       =   2580
   ClientWidth     =   4575
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   120
   End
   Begin VB.Shape shpShade 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      DrawMode        =   7  'Invert
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shpBorder 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hello, this is my new progress bar for use with everthing"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Dim Add As Integer
Const Steps = 200
Dim Counter As Integer
Dim pause As Variant

If Add = 0 Then
    Add = shpBorder.Width \ Steps
    shpShade.Width = 0
End If

For Counter = 1 To (Steps + Add)
shpShade.Width = shpShade.Width + Add

If (shpShade.Width > shpBorder.Width) Then
    shpShade.Width = shpBorder.Width
    'Timer1.Enabled = False
    Exit For
End If

'show percentage completed
Label1.Caption = ((100 / shpBorder.Width) * (shpShade.Width) Mod 101) & "% Complete"

For pause = 1 To 300000
Next pause
Next Counter
End Sub

Private Sub Timer1_Timer()
Static Add As Integer
Const Steps = 200

If Add = 0 Then
    Add = shpBorder.Width \ Steps
    shpShade.Width = 0
End If

shpShade.Width = shpShade.Width + Add

If (shpShade.Width > shpBorder.Width) Then
    shpShade.Width = shpBorder.Width
    Timer1.Enabled = False
End If
End Sub
