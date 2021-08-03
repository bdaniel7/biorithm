VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3840
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   3840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      Caption         =   "&OK"
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   810
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmAbout.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   810
      TabIndex        =   1
      Top             =   105
      Width           =   2835
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const msgAbout = "Copyright 1994-2001 Daniel Blendea" & vbCr & "Contact : bdaniel7@excite.com"

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case 13
      Unload Me
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
lblOk.BackColor = vbButtonFace
Me.Move 9500, 3000
'cmdExit.Width = Me.ScaleWidth
'cmdExit.Move 0, Me.ScaleHeight - cmdExit.Height 'Me.Width / 4
lblOk.Width = 1000
lblOk.Move Me.Width - 3 * lblOk.Width + (lblOk.Width / 2), Me.ScaleHeight - 450, 1000, 250
lblAbout.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & vbCrLf & msgAbout
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.BackColor = vbButtonFace
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.BackColor = vbButtonFace
End Sub
Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.BackColor = vbButtonFace
End Sub

Private Sub lblOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.BackColor = vbGrayText
End Sub

Private Sub lblOk_Click()
Unload Me
End Sub
