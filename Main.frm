VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   Caption         =   "Bioritm"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "News Gothic MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCompat 
      Caption         =   "Compatibilitate"
      Height          =   495
      Left            =   6795
      TabIndex        =   15
      Top             =   4965
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   7185
      TabIndex        =   11
      Top             =   2085
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.ComboBox cboAnT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6315
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1620
      Width           =   1530
   End
   Begin VB.ComboBox cboZi 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3240
      Width           =   1515
   End
   Begin VB.ComboBox cboLuna 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4860
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4755
      Width           =   1500
   End
   Begin VB.ComboBox cboAn 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4845
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4305
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7515
      Top             =   2505
   End
   Begin VB.CommandButton cmdCalcul 
      Caption         =   "&Calculeazã"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6660
      TabIndex        =   4
      Top             =   4290
      Width           =   1605
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FontTransparent =   0   'False
      Height          =   3510
      Left            =   405
      ScaleHeight     =   3450
      ScaleWidth      =   5745
      TabIndex        =   5
      Top             =   450
      Width           =   5805
      Begin VB.HScrollBar HSB 
         Height          =   225
         Left            =   420
         TabIndex        =   6
         Top             =   2115
         Width           =   4965
      End
   End
   Begin VB.Line IntLine 
      X1              =   6855
      X2              =   8475
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Line FizLine 
      X1              =   8520
      X2              =   9810
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line AfectLine 
      X1              =   6825
      X2              =   8445
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Label lblBioFizic 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6795
      TabIndex        =   14
      Top             =   45
      Width           =   1665
   End
   Begin VB.Label lblBioAfect 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6765
      TabIndex        =   13
      Top             =   990
      Width           =   1665
   End
   Begin VB.Label lblBioInt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   12
      Top             =   525
      Width           =   1665
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2835
      TabIndex        =   10
      Top             =   4785
      Width           =   1275
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3270
      TabIndex        =   9
      Top             =   4290
      Width           =   1275
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1740
      TabIndex        =   8
      Top             =   4305
      Width           =   1275
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   7
      Top             =   4290
      Width           =   1275
   End
   Begin VB.Menu mnuBioritm 
      Caption         =   "&Bioritm..."
      Begin VB.Menu mnuCompat 
         Caption         =   "&Compatibilitate"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDespre 
         Caption         =   "&Despre"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Bio As New Bioritm
Dim i, j, An As Integer, maxZile As Integer, Luna As Byte
Dim Zile() As String, Luni(12) As Byte, Coef()
Dim Xi, Yi, Xf, Yf, k, p, S, Bisect As Boolean, IsBis, Step, Culoare
Dim X0, Y0, Xmax, Ymax, Out, AnN As Integer, LunaN As Byte, ZiN As Byte, AnTarget As Integer
Dim Y2, Y3, X2, X3, ZileAn As Integer, Compat(34, 4) As Single
Const hLine = 460
Const hStep = 33
Const wStep = 20
Const Diff = 190
Public procIntel As Integer, procAfect As Integer, procFizic As Integer, procTotal As Integer, msgFiz As String, msgAfect As String, msgIntel As String, msgTotal As String
Private Sub cboAn_Click()
   AnN = Val(cboAn.List(cboAn.ListIndex))
End Sub

Private Sub cboAn_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cboAnT.SetFocus
   KeyAscii = 0
End If
End Sub

Private Sub cboAnT_Click()
   AnTarget = Val(cboAnT.List(cboAnT.ListIndex))
End Sub

Private Sub cboAnT_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If cboAnT.Text <> "" And cboZi.Text <> "" And cboLuna.Text <> "" And cboAn.Text <> "" Then
      cmdCalcul_Click
      KeyAscii = 0
   End If
End If
End Sub

Private Sub cboLuna_Click()
   LunaN = 0
   LunaN = Val(Combo1.List(cboLuna.ListIndex))
End Sub

Private Sub cboLuna_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cboAn.SetFocus
   KeyAscii = 0
End If
End Sub

Private Sub cboZi_Click()
   ZiN = Val(cboZi.List(cboZi.ListIndex))
End Sub
Private Sub cboZi_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cboLuna.SetFocus
   KeyAscii = 0
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case vbKeyZ
      If (Shift And vbAltMask) Then
         cboZi_Click
      End If
   Case vbKeyL
      If (Shift And vbAltMask) Then
         cboLuna_Click
      End If
   Case vbKeyN
      If (Shift And vbAltMask) Then
         cboAn_Click
      End If
   Case vbKeyN
      If (Shift And vbAltMask) Then
         cboAnT_Click
      End If
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case vbKeyEscape
      'Unload Me
End Select
End Sub
Private Sub Form_Resize()
'Init Pic
'ScrollBarInit Pic
'cmdCalcul_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unhook
End Sub

Private Sub HSB_Change()
   HSB.Refresh
   Init Pic
   ScrollBarInit Pic
   Step = HSB.Value
   DrawPointsOn Pic, Step
   DrawLinesOn Pic, Step
   PrintDaysOn Pic, Step
End Sub

Private Sub HSB_Scroll()
   HSB_Change
End Sub

Private Sub mnuCompat_Click()
cmdCompat.Visible = True
End Sub

Private Sub mnuDespre_Click()
Load frmAbout
frmAbout.Show vbModal
'MsgBox "Bioritmul uman." & vbCrLf & "Copyright 1994-2001 © Daniel Blendea." & vbCrLf & "Feedback : bdaniel7@excite.com", vbInformation
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub
Private Sub PrintDaysOn(Object, Step)
If Bisect = True Then
   IsBis = 5
Else
   IsBis = 4
End If
For i = 1 To 32
      Out = i + Step
      If Out > maxZile Then
         Exit Sub
      Else
         Select Case Len(Coef(3, Out))
            Case Is = 3
               S = 9
            Case Is = 4
               S = 10
            Case Is = 5
               S = 11
         End Select
         Object.ForeColor = vbBlack
         Object.CurrentX = i * hStep - S
         Object.CurrentY = wStep + 440
         Object.Print Coef(3, i + Step)      '//   ziua si luna din an
         
         Object.ForeColor = vbBlue
         Object.CurrentX = i * hStep - S
         Object.CurrentY = wStep + 470
         Object.Print Coef(2, i + Step)      '//   intelectual
         
         Object.ForeColor = vbYellow
         Object.CurrentX = i * hStep - S
         Object.CurrentY = wStep + 490
         Object.Print Coef(7, i + Step)      '//   afectiv
         
         Object.ForeColor = vbRed
         Object.CurrentX = i * hStep - S
         Object.CurrentY = wStep + 510
         Object.Print Coef(5, i + Step)      '//   fizic
      End If
Next i
End Sub
Private Sub DrawPointsOn(Object, Step)
Object.ScaleMode = vbPixels
Object.DrawWidth = 5
If Bisect = True Then
   IsBis = 5
Else
   IsBis = 4
End If
For i = 1 To 32
      Out = i + Step
      If Out > maxZile Then
         Exit Sub
      Else
         Xi = i * hStep
         Yi = Coef(4, Out) * wStep - Diff
         Object.PSet (Xi, Yi), vbBlue
         Y2 = Coef(6, Out) * wStep - Diff
         Object.PSet (Xi, Y2), vbRed
         Y3 = Coef(8, Out) * wStep - Diff
         Object.PSet (Xi, Y3), vbYellow
      End If
Next i
Object.FontBold = False
End Sub
Private Sub DrawLinesOn(Object, Step)
Object.ScaleMode = vbPixels
Object.DrawWidth = 1
If Bisect = True Then
   IsBis = 5
Else
   IsBis = 4
End If
   For i = 1 To 32
      Out = i + Step
      If Out > maxZile - 1 Then
         Exit Sub
      Else
         Xi = i * hStep
         Yi = Coef(4, Out) * wStep - Diff
         Xf = (i + 1) * hStep
         Yf = Coef(4, Out + 1) * wStep - Diff
         Culoare = vbGreen
         Object.Line (Xi, Yi)-(Xi, hLine), Culoare
         Object.Line (Xi, Yi)-(Xf, Yf), vbBlue
         
         Xi = i * hStep
         Yi = Coef(6, Out) * wStep - Diff
         Xf = (i + 1) * hStep
         Yf = Coef(6, Out + 1) * wStep - Diff
         Object.Line (Xi, Yi)-(Xi, hLine), Culoare
         Object.Line (Xi, Yi)-(Xf, Yf), vbRed
         
         Xi = i * hStep
         Yi = Coef(8, Out) * wStep - Diff
         Xf = (i + 1) * hStep
         Yf = Coef(8, Out + 1) * wStep - Diff
         Object.Line (Xi, Yi)-(Xi, hLine), Culoare

         Object.Line (Xi, Yi)-(Xf, Yf), vbYellow
      End If
   Next i
End Sub

Private Sub Form_Load()

Bio.TestScreen
Resize Me, , , 15000, 10500
Step = 0
Init Pic
ScrollBarInit Pic
gHW = Me.hWnd
Hook
LoadCombo
DrawLegend Me
LoadCompat
End Sub

Private Sub Resize(Object As Object, Optional LeftY As Single = 0.5, Optional Topx As Single = 0.5, Optional myWidth As Integer, Optional myHeight As Integer)
Object.Width = myWidth
Object.Height = myHeight
Object.Left = LeftY * Screen.Width - Width \ 2
Object.Top = Topx * Screen.Height - Height \ 2
End Sub
Private Sub ScrollBarInit(Object)
Object.ScaleMode = vbPixels
HSB.Width = Object.ScaleWidth - 1
HSB.Left = Object.Left
HSB.Top = Object.ScaleHeight - HSB.Height
HSB.Max = 337 'Object.ScaleWidth
HSB.LargeChange = HSB.Max / 11
HSB.SmallChange = 1
End Sub
Private Sub Init(Object)
Object.ScaleMode = vbPixels
Object.Cls
Object.Refresh
Object.DrawWidth = 2
Object.Left = Me.CurrentX
Object.Top = Me.CurrentY
Object.Width = Me.ScaleWidth '180000
Object.Height = Me.ScaleHeight - 1200
Xmax = Object.ScaleWidth - 1
Ymax = Object.ScaleHeight - 1
CtlInit
End Sub
Private Sub CtlInit()
cmdCalcul.Top = Pic.Height + 450: cmdCalcul.Left = 9000

cboZi.Top = Pic.Height + 600: cboZi.Left = 500
cboLuna.Top = Pic.Height + 600: cboLuna.Left = cboZi.Width + 700
cboAn.Top = Pic.Height + 600: cboAn.Left = cboLuna.Width + 2400
cboAnT.Top = Pic.Height + 600: cboAnT.Left = cboAn.Width + 5000

Label1.Top = Pic.Height + 300: Label1.Left = cboZi.Left
Label2.Top = Pic.Height + 300: Label2.Left = cboLuna.Left
Label3.Top = Pic.Height + 300: Label3.Left = cboAn.Left
Label4.Top = Pic.Height + 300: Label4.Left = cboAnT.Left

lblBioInt.Top = Pic.Height + 250: lblBioInt.Left = cmdCalcul.Left + cmdCalcul.Width + 500
lblBioInt.ForeColor = vbBlue: lblBioInt.Caption = "Bioritm intelectual : "

lblBioAfect.Top = Pic.Height + lblBioInt.Height + 100: lblBioAfect.Left = cmdCalcul.Left + cmdCalcul.Width + 500
lblBioAfect.ForeColor = vbYellow: lblBioAfect.Caption = "Bioritm emoþional : "

lblBioFizic.Top = Pic.Height + lblBioAfect.Height + 350: lblBioFizic.Left = cmdCalcul.Left + cmdCalcul.Width + 500
lblBioFizic.ForeColor = vbRed: lblBioFizic.Caption = "Bioritm fizic : "

Label1.Caption = "&Ziua"
Label2.Caption = "&Luna"
Label3.Caption = "Anul &naºterii"
Label4.Caption = "Anul þin&tã"
End Sub
Private Sub DrawLegend(Object)
   Me.IntLine.BorderWidth = 3: Me.IntLine.BorderColor = vbBlue
   Me.IntLine.X1 = 12825: Me.IntLine.Y1 = 8955
   Me.IntLine.X2 = 14175: Me.IntLine.Y2 = 8955
   
   Me.AfectLine.BorderWidth = 3: Me.AfectLine.BorderColor = vbYellow
   Me.AfectLine.X1 = 12825: Me.AfectLine.Y1 = 9210
   Me.AfectLine.X2 = 14175: Me.AfectLine.Y2 = 9210
   
   Me.FizLine.BorderWidth = 3: Me.FizLine.BorderColor = vbRed
   Me.FizLine.X1 = 12825: Me.FizLine.Y1 = 9465
   Me.FizLine.X2 = 14175: Me.FizLine.Y2 = 9465
End Sub
Public Sub CalcBio(ByVal AnNastere As Integer, ByVal LunaNastere As Byte, ByVal ZiNastere As Byte, Optional ByVal AnCurent As Integer)
maxZile = 0
If AnCurent = 0 Then AnCurent = Year(Date)
If AnCurent Mod 4 = 0 Then
   ZileAn = 366
Else
   ZileAn = 365
End If
For i = 1 To 12
   Luni(i) = Bio.MonthLen(i, AnCurent)      '// calculeaza nr de zile din fiecare luna
   If Luni(i) = 29 Then
      Bisect = True
   Else
      Bisect = False
   End If
Next i
For i = 1 To 12
   maxZile = maxZile + Luni(i)   '// calculeaza nr de zile din an : 365/366
Next i
ReDim Zile(5, maxZile)  '// 3 linii si maxZile coloane
k = 0
i = 1
j = 1
For i = 1 To 12   '// pentru fiecare luna
   For j = 1 To Luni(i)    '// pentru nr de zile coresp fiecarei luni
      Bio.CalculBioritm AnNastere, LunaNastere, ZiNastere, AnCurent, i, j
      Zile(1, k + 1) = k + 1        '//   linia 1 = numarul zilei (1-365)
      Zile(2, k + 1) = Bio.bIntel   '//   linia 2 = val coeficient bioritm intelectual
      Zile(3, k + 1) = j & "." & i  '//   linia 3 = ziua.luna din an
      Zile(4, k + 1) = Bio.bFizic   '//   linia 4 = bioritm fizic
      Zile(5, k + 1) = Bio.bAfectiv '//   linia 4 = bioritm afectiv
      'Text1.Text = Text1.Text & vbCrLf & Zile(1, k + 1) & " / " & Zile(3, k + 1) & " - " & Zile(4, k + 1)
      k = k + 1
   Next j
Next i
ReDim Coef(8, maxZile)
For i = 1 To maxZile
   Coef(1, i) = Zile(1, i)    '//   linia 1 = numarul zilei
   Coef(3, i) = Zile(3, i)    '//   linia 3 = ziua.luna din an
   Coef(2, i) = Zile(2, i)    '//   linia 2 = val coeficient bioritm intelectual
   Coef(5, i) = Zile(4, i)    '//   linia 5 = val coef bioritm fizic
   Coef(7, i) = Zile(5, i)    '//   linia 7 = val coef bioritm afectiv
   'Text1.Text = Text1.Text & vbCrLf & Coef(1, i) & " / " & Coef(3, i) & " - " & Coef(5, i)
Next i
For k = 1 To ZileAn
   Select Case Coef(2, k)
   Case Is < 16
      p = 32 - Coef(2, k)
   Case Else
      p = Coef(2, k)
   End Select
   Coef(4, k) = p    '//   linia 4 = valoare ciclu intelectual
   'Text1.Text = Text1.Text & vbCrLf & Coef(1, k) & " / " & Coef(3, k) & " - " & Coef(4, k)
Next k
For k = 1 To ZileAn
   Select Case Coef(5, k)
   Case Is < 11
      p = 22 - Coef(5, k)
   Case Else
      p = Coef(5, k)
   End Select
   Coef(6, k) = p    '//   linia 6 = valoare ciclu fizic
   'Text1.Text = Text1.Text & vbCrLf & Coef(1, k) & " / " & Coef(3, k) & " - " & Coef(6, k)
Next k
For k = 1 To ZileAn
   Select Case Coef(7, k)
   Case Is < 13
      p = 27 - Coef(7, k)
   Case Else
      p = Coef(7, k)
   End Select
   Coef(8, k) = p    '//   linia 6 = valoare ciclu afectiv
   'Text1.Text = Text1.Text & vbCrLf & Coef(1, k) & " / " & Coef(3, k) & " - " & Coef(8, k)
Next k
'Debug.Print k & vbCrLf & maxZile
End Sub
Private Sub cmdCalcul_Click()

Timer1.Enabled = True
Timer1_Timer
Me.Caption = "Bioritm "
Me.Caption = Me.Caption & " pentru data naºterii  " & _
Format(Format(DateSerial(AnN, LunaN, ZiN), "DDDD - dd.MMMM.yyyy"), ">") & _
" în anul " & AnTarget
CalcBio AnN, LunaN, ZiN, AnTarget
'Init Pic
'ScrollBarInit Pic
Pic.Cls
DrawPointsOn Pic, Step
DrawLinesOn Pic, Step
PrintDaysOn Pic, Step

End Sub

Public Sub CheckCompat(ByVal Zi1 As Byte, ByVal Luna1 As Byte, ByVal An1 As Integer, ByVal Zi2 As Byte, ByVal luna2 As Byte, ByVal An2 As Integer)
Dim Data1 As Date, Data2 As Date, Defazaj As Integer, DifIntel As Byte, DifAfect As Byte, DifFizic As Byte
Data1 = DateSerial(An1, Luna1, Zi1)
Data2 = DateSerial(An2, luna2, Zi2)
Defazaj = DateDiff("d", Data1, Data2, vbMonday, vbFirstJan1) + 1

DifIntel = Defazaj Mod 33
DifAfect = Defazaj Mod 28
DifFizic = Defazaj Mod 23
For i = 1 To 34
'Debug.Print Compat(i, 1) & " : " & Compat(i, 4)
   Select Case DifIntel
   Case Compat(i, 1)
         procIntel = Compat(i, 4)             '//   compat intelectuala
   Case 17
         procIntel = Compat(18, 4)
   Case 11.5
         procIntel = Compat(12, 4)
   End Select
Next i
For i = 1 To 29
'Debug.Print Compat(i, 1) & " : " & Compat(i, 3)
Select Case DifAfect
   Case Compat(i, 1)
      procAfect = Compat(i, 3)               '//   compat emotionala
   Case 17, 16.5
      procAfect = Compat(18, 3)
   Case 11.5
      procAfect = Compat(12, 3)
   End Select
Next i
For i = 1 To 25
'Debug.Print Compat(i, 1) & " : " & Compat(i, 2)
   Select Case DifFizic
   Case Compat(i, 1)
      procFizic = Compat(i, 2)             '//   compat fizica
   Case 17
       procFizic = Compat(18, 2)
   End Select
Next i
procTotal = (procIntel + procAfect + procFizic) / 3
Select Case procTotal
   Case 0 To 25
      msgTotal = "Relaþii slabe."
   Case 26 To 50
      msgTotal = "Relaþii mijlocii."
   Case 51 To 75
      msgTotal = "Relaþii puternice."
   Case 76 To 100
      msgTotal = "Relaþii remarcabile."
End Select
Debug.Print "Total : " & procTotal & " %"
Debug.Print "Intelectual : " & procIntel & " %"
Debug.Print "Emotional : " & procAfect & " %"
Debug.Print "Fizic : " & procFizic & " %"
Debug.Print msgTotal
End Sub
Private Sub cmdCompat_Click()
'CheckCompat 7, 4, 1979, 10, 6, 1980
'CheckCompat 7, 4, 1979, 13, 4, 1979
End Sub

Private Sub LoadCompat()
'// Coloana 1 = Defazaj zile
Compat(1, 1) = 0: Compat(2, 1) = 1: Compat(3, 1) = 2: Compat(4, 1) = 3: Compat(5, 1) = 4: Compat(6, 1) = 5
Compat(7, 1) = 6: Compat(8, 1) = 7: Compat(9, 1) = 8: Compat(10, 1) = 9: Compat(11, 1) = 10
Compat(12, 1) = 11: Compat(13, 1) = 11.5: Compat(14, 1) = 12: Compat(15, 1) = 13: Compat(16, 1) = 14
Compat(17, 1) = 15: Compat(18, 1) = 16: Compat(19, 1) = 16.5: Compat(20, 1) = 18: Compat(21, 1) = 19
Compat(22, 1) = 20: Compat(23, 1) = 21: Compat(24, 1) = 22: Compat(25, 1) = 23: Compat(26, 1) = 24
Compat(27, 1) = 25: Compat(28, 1) = 26: Compat(29, 1) = 28: Compat(30, 1) = 29: Compat(31, 1) = 30
Compat(32, 1) = 31: Compat(33, 1) = 32: Compat(34, 1) = 33
'//      Coloana 2 = Grad compatibilitate fizica
Compat(1, 2) = 100: Compat(2, 2) = 91: Compat(3, 2) = 83: Compat(4, 2) = 74: Compat(5, 2) = 65
Compat(6, 2) = 56: Compat(7, 2) = 48: Compat(8, 2) = 39: Compat(9, 2) = 30: Compat(10, 2) = 22
Compat(11, 2) = 13: Compat(12, 2) = 4: Compat(13, 2) = 0: Compat(14, 2) = 4: Compat(15, 2) = 13
Compat(16, 2) = 22: Compat(17, 2) = 30: Compat(18, 2) = 39: Compat(19, 2) = 0: Compat(20, 2) = 56
Compat(21, 2) = 65: Compat(22, 2) = 74: Compat(23, 2) = 83: Compat(24, 2) = 91: Compat(25, 2) = 100
'//      Coloana 3 = Grad de compatibilitate emotionala
Compat(1, 3) = 100: Compat(2, 3) = 93: Compat(3, 3) = 86: Compat(4, 3) = 79: Compat(5, 3) = 71
Compat(6, 3) = 64: Compat(7, 3) = 57: Compat(8, 3) = 50: Compat(9, 3) = 43: Compat(10, 3) = 36
Compat(11, 3) = 29: Compat(12, 3) = 21: Compat(13, 3) = 0: Compat(14, 3) = 14: Compat(15, 3) = 7
Compat(16, 3) = 0: Compat(17, 3) = 7: Compat(18, 3) = 14: Compat(19, 3) = 0: Compat(20, 3) = 29
Compat(21, 3) = 36: Compat(22, 3) = 43: Compat(23, 3) = 50: Compat(24, 3) = 57: Compat(25, 3) = 64
Compat(26, 3) = 71: Compat(27, 3) = 79: Compat(28, 3) = 86: Compat(29, 3) = 100
'//      Coloana 4 = Grad de compatibilitate intelectuala
Compat(1, 4) = 100: Compat(2, 4) = 94: Compat(3, 4) = 88: Compat(4, 4) = 82: Compat(5, 4) = 76
Compat(6, 4) = 70: Compat(7, 4) = 64: Compat(8, 4) = 58: Compat(9, 4) = 52: Compat(10, 4) = 46
Compat(11, 4) = 39: Compat(12, 4) = 33: Compat(13, 4) = 0: Compat(14, 4) = 27: Compat(15, 4) = 21
Compat(16, 4) = 15: Compat(17, 4) = 9: Compat(18, 4) = 3: Compat(19, 4) = 0: Compat(20, 4) = 9
Compat(21, 4) = 15: Compat(22, 4) = 21: Compat(23, 4) = 27: Compat(24, 4) = 33: Compat(25, 4) = 39
Compat(26, 4) = 46: Compat(27, 4) = 52: Compat(28, 4) = 58: Compat(29, 4) = 70: Compat(30, 4) = 76
Compat(31, 4) = 82: Compat(32, 4) = 88: Compat(33, 4) = 94: Compat(34, 4) = 100
End Sub
Private Sub LoadCombo()
   For i = 1 To 200
      cboAn.List(i - 1) = 1912 + i
   Next
   For i = 1 To 12
      Combo1.List(i - 1) = i
      cboLuna.List(i - 1) = Format(MonthName(i), ">")
   Next
   For i = 1 To 31
      cboZi.List(i - 1) = i
   Next
   For i = 1 To 200
      cboAnT.List(i - 1) = i + 1900
   Next
   'cboZi.SelText = Day(Date)
   'cboLuna.SelText = Month(Date)
   'cboAn.SelText = Year(Date)
   'cboAnT.SelText = Year(Date)
End Sub
Private Sub Timer1_Timer()
If AnTarget = 0 Or ZiN = 0 Or LunaN = 0 Or AnN = 0 Then
   cmdCalcul.Enabled = False
Else
   cmdCalcul.Enabled = True
End If
End Sub
