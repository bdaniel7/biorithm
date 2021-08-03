VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3675
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2895
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
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   2895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   540
      Top             =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   405
      TabIndex        =   1
      Top             =   570
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   735
      Top             =   1650
   End
   Begin MSComctlLib.ProgressBar Progb 
      Height          =   2880
      Left            =   1740
      TabIndex        =   0
      Top             =   180
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   5080
      _Version        =   393216
      Appearance      =   0
      Max             =   50
      Orientation     =   1
   End
   Begin VB.Label lblCaption 
      Height          =   345
      Left            =   330
      TabIndex        =   2
      Top             =   3135
      Width           =   2280
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim start As Single
Dim procent, proc As Byte, SoundBuffer() As Byte
Const Vmax = 70
Public ShowType
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2 ' Don't use default sound
Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
BeginPlaySound 102
Unload Me
End Sub
Private Sub Switch()
Select Case ShowType
   Case "load"
      lblCaption.Caption = "Loading, please wait..."
      BeginPlaySound 101
   Case "unload"
      lblCaption.Caption = "Unloading, please wait..."
      BeginPlaySound 102
End Select
End Sub
Private Sub Form_Load()
Switch
Timer2.Enabled = False
Progb.Max = Vmax
Progb.Min = 0
Progb.Left = Me.Width / 2 - Progb.Width / 2
lblCaption.Left = Me.Width / 2 - lblCaption.Width / 2
End Sub
Private Sub BeginPlaySound(resId)
SoundBuffer = LoadResData(resId, "SOUND")
sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub
Private Sub StopPlaySound()
sndPlaySound ByVal vbNullString, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
'StopPlaySound
End Sub

Private Sub Timer1_Timer()
Select Case ShowType
   Case "load"
      If start = 0! Then
         start = Timer
      End If
      procent = Vmax * (Timer - start) / 2
      If procent < Vmax Then
         Progb.Value = procent
         proc = procent
      Else
         Progb.Value = Vmax
         proc = Vmax
         lblCaption.Caption = "Done !"
         Timer1.Enabled = False
         Timer2.Enabled = True
         Timer2_Timer
      End If
   Case "unload"
         If start = 0! Then
         start = Timer
      End If
      procent = Vmax * (Timer - start) / 2
      If procent < Vmax Then
         Progb.Value = Progb.Max - procent
         proc = procent
      Else
         Progb.Value = 0
         proc = Vmax
         lblCaption.Caption = "Quittin'"
         Timer1.Enabled = False
         Timer2.Enabled = True
         Timer2_Timer
      End If
   End Select
End Sub

Private Sub Timer2_Timer()
Unload Me
End Sub
