VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Screen"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame framAsk 
      Caption         =   "Password"
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdCan 
         Caption         =   "&Cancel"
         Height          =   335
         Left            =   3240
         TabIndex        =   16
         Top             =   470
         Width           =   855
      End
      Begin VB.CheckBox chkPassOn 
         Alignment       =   1  'Right Justify
         Caption         =   "Activate Password"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "C&hange"
         Height          =   335
         Left            =   3240
         TabIndex        =   10
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "&Enter"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Press [Return] to set password."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   570
         Width           =   2055
      End
      Begin VB.Label lblEnter 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame framChange 
      Caption         =   "Change Password"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set Password"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtRetype 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtOld 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblRetype 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Type Password"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblNew 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New Password"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblOld 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Old Password"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPass"
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

'the different sizes for the two different functions
'of this screen
Const ChangeHeight = 2430
Const ChangeWidth = 4785
Const AskHeight = 1230 '1590 '1470
Const AskWidth = 4305 '3345

Const RetKey = 13

Private Sub chkPassOn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'activate/deactivate password
PassActive = Not PassActive

chkPassOn.Value = (PassActive * -1)
End Sub

Private Sub cmdCan_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmHandsClk.Show
End Sub

Private Sub cmdChange_Click()
AskOrChange = Change
Call SetScreen
End Sub

Private Sub cmdEnter_Click()
If txtPass.Text = Password Then
    'correct password entered
    CorrectPass = True
    txtPass.Text = ""
    Unload Me
    frmHandsClk.Show
Else
    'incorrect password
    CorrectPass = False
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End If
End Sub

Private Sub cmdSet_Click()
If (txtRetype.Text <> txtNew.Text) Then
    'confirm new password again
    txtRetype.SetFocus
    txtRetype.SelStart = 0
    txtRetype.SelLength = Len(txtRetype.Text)
    Exit Sub
Else
    If Password = txtOld.Text Then
        'set password
        Password = txtNew.Text
        CorrectPass = True
        txtNew.Text = ""
        txtOld.Text = ""
        txtRetype.Text = ""
        Call frmHandsClk.SaveStatus
        Unload Me
        frmHandsClk.Show
    Else
        txtOld.SetFocus
    End If
End If
    
End Sub

Private Sub Form_Activate()
Call SetScreen
End Sub

Private Sub SetScreen()
'show a different screen depending on which function
'is needed.
Select Case AskOrChange
Case "Ask"
    'ask the user for the password
    If Password <> "" Then
        chkPassOn.Enabled = False
    End If
    chkPassOn.Value = (PassActive * -1)
    frmPass.Height = AskHeight
    frmPass.Width = AskWidth
    framAsk.Visible = True
    framChange.Visible = False
    txtPass.SetFocus

Case "Change"
    'change the existing/set a new password
    frmPass.Height = ChangeHeight
    frmPass.Width = ChangeWidth
    framAsk.Visible = False
    framChange.Visible = True
    
    'if there is no password, set one.
    cmdSet.Enabled = False
    If Password = "" Then
        txtOld.Enabled = False
        txtOld.BackColor = frmPass.BackColor
        txtNew.SetFocus
    Else
        txtOld.Enabled = True
        txtOld.BackColor = txtNew.BackColor
        txtOld.SetFocus
    End If
    
End Select

Me.Visible = True
AskOrChange = False
End Sub

Private Sub Form_Load()
If Password = "" Then
    'enter a new password
    AskOrChange = Change
End If
End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRetype.SetFocus
End If
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
If KeyAscii = RetKey Then
    txtNew.SetFocus
End If
End Sub

Private Sub txtOld_LostFocus()
If (txtOld.Text <> Password) And (AskOrChange = Change) Then
    txtOld.SetFocus
    txtOld.SelStart = 0
    txtOld.SelLength = Len(txtOld.Text)
End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = RetKey Then
    cmdEnter_Click
    Call ActivateCheck(txtPass.Text)
End If
End Sub

Private Sub txtRetype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'set password
    cmdSet_Click
End If

End Sub

Private Sub ActivateCheck(Pass As String)
'This procedure activtes or deactivates the check
'box.

If Pass = Password Then
    chkPassOn.Enabled = True
Else
    chkPassOn.Enabled = False
End If

chkPassOn.Value = (PassActive * -1)
End Sub

Private Sub txtRetype_KeyUp(KeyCode As Integer, Shift As Integer)
'enable or disable command button
If txtNew.Text = txtRetype.Text Then
    cmdSet.Enabled = True
Else
    cmdSet.Enabled = False
End If
End Sub
