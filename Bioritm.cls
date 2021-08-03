VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bioritm"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private MyShell As New Shell
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim NrZile As Integer
'Dim Ani(2, 200)
Dim DataN As Date, DataCrt As Date
Dim i
Const Intel = 33
Const Afectiv = 28
Const Fizic = 23
Const msgIntel = "Coeficient bioritm intelectual : "
Const msgAfectiv = "Coeficient bioritm afectiv : "
Const msgFizic = "Coeficient bioritm fizic : "
Dim Bioritm(31) As Byte
Public bIntel As Integer, bAfectiv As Integer, bFizic As Integer
Public Sub CalculBioritm(ByVal AnN As Integer, ByVal LunaN As Byte, ByVal ZiN As Byte _
 , Optional ByVal AnCrt As Integer, Optional ByVal LunaCrt As Byte, Optional ByVal ZiCrt As Byte)
If AnCrt = 0 Then AnCrt = Year(Date)
If LunaCrt = 0 Then LunaCrt = Month(Date)
If ZiCrt = 0 Then ZiCrt = Day(Date)
DataN = DateSerial(AnN, LunaN, ZiN)
DataCrt = DateSerial(AnCrt, LunaCrt, ZiCrt)
NrZile = DateDiff("d", DataN, DataCrt, vbMonday, vbFirstJan1) + 1

bIntel = NrZile Mod Intel
bAfectiv = NrZile Mod Afectiv
bFizic = NrZile Mod Fizic

'MsgBox msgIntel & bIntel & vbCr & _
'msgAfectiv & bAfectiv & vbCr & _
'msgFizic & bFizic & "." & vbCr & NrZile & " zile.", vbInformation, "Bioritm pentru ziua de " & Format(DataCrt, "dd mmmm yyyy")
End Sub
Public Function GetWinDir() As String
Dim n As Integer
Dim strPath As String
strPath = Space$(144)
n = GetWindowsDirectory(strPath, 144)
strPath = Left$(strPath, n)
GetWinDir = strPath
End Function
Public Function GetWinSysDir() As String
Dim n As Integer
Dim strPath As String
strPath = Space$(144)
n = GetSystemDirectory(strPath, 144)
strPath = Left$(strPath, n)
GetWinSysDir = strPath
End Function
Public Function MonthLen(ByVal Luna As Byte, Optional An As Integer) As Byte
Dim intLen As Byte
   Select Case Luna
      Case 2
         If An = 0 Then
            MsgBox "Pentru luna februarie trebuie introdus ºi anul", _
            vbInformation + vbOKOnly, "An bisect sau nu ?"
            Exit Function
         Else
            If An Mod 4 = 0 Then
               intLen = 29
            Else
               intLen = 28
            End If
         End If
      Case 1, 3, 5, 7, 8, 10, 12
         intLen = 31
      Case 4, 6, 9, 11
         intLen = 30
      Case Else
         MsgBox "Nu sunt decât 12 luni într-un an !", vbExclamation + vbOKOnly, "EROARE !"
         Exit Function
   End Select
MonthLen = intLen
End Function
Public Sub TestScreen()
Dim ScreenW, ScreenH, Answer, CplPath
ScreenW = Screen.Width / Screen.TwipsPerPixelX
ScreenH = Screen.Height / Screen.TwipsPerPixelY
CplPath = GetWinSysDir & "\Desk.cpl"
If ScreenW < 1024 Or ScreenH < 768 Then
   Answer = MsgBox("Pentru a rula aceastã aplicaþie în condiþii optime se recomandã" & vbCrLf & _
   "setarea rezoluþiei monitorului la 1024/786 pixeli." & vbCrLf & vbCrLf & _
   "Schimbaþi acum rezoluþia ecranului ? ", vbQuestion + vbYesNo + vbDefaultButton1)
   If Answer = vbYes Then
      MyShell.ControlPanelItem CplPath
   End If
End If
End Sub
