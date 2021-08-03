Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetActiveWindow _
Lib "User32" () As Long
Private Declare Function GetLastError _
Lib "kernel32" ( _
 _
) As Long
Private Declare Function EnumChildWindows _
Lib "User32" ( _
    ByVal hWnd As Long, _
    ByVal lpWndProc As Long, _
    ByVal lp As Long _
) As Long
Private Declare Function GetLocaleInfo _
Lib "kernel32" Alias "GetLocaleInfoA" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long _
) As Long

Private Declare Function EnumSystemLocales Lib "kernel32" _
(ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetUserDefaultLCID _
Lib "kernel32" ( _
 _
) As Long
Private Declare Function GetSystemDefaultLangID _
Lib "kernel32" ( _
 _
) As Integer

Private Declare Function GetUserDefaultLangID _
Lib "kernel32" ( _
 _
) As Integer

Private Declare Function GetSystemDefaultLCID _
Lib "kernel32" ( _
 _
) As Long

Public Lang
Sub GetLocal()
Call GetLocaleInfo(1048, &H9, Lang, LenB(Lang))
Main.Caption = Lang
MsgBox GetUserDefaultLangID
End Sub
Sub EnumW()
    Dim hWnd As Long
    Dim x As Long
    'Get a handle to the active window
    hWnd = GetActiveWindow()
    If (hWnd) Then
        'Call EnumChildWindows API, which calls
        'ChildWindowProc for each child window and then ends
        x = EnumChildWindows(hWnd, AddressOf ChildWindowProc, 0)
    End If
End Sub

'Called by EnumChildWindows API function
Function ChildWindowProc( _
    ByVal hWnd As Long, _
    ByVal lp As Long _
) As Long
    'hWnd and lp parameters are passed in by EnumChildWindows
    Main.Pic.Print "Window: "; hWnd
    'Return success (in C, 1 is True and 0 is False)
    ChildWindowProc = 1
End Function

Sub EnumLocal()
Call EnumSystemLocales(AddressOf EnumProc, 1)
End Sub
Function EnumProc(LocaleStr As String) As Long
Debug.Print LocaleStr
EnumProc = 1
End Function



