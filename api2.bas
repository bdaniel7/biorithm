Attribute VB_Name = "Api"
Option Explicit

Private Const GWL_USERDATA = -21
Public Const GWL_WNDPROC = -4
Private lpPrevWndProc As Long
Public gHW As Long
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_CLOSE = &H10
Private Const SC_SIZE = &HF000&
Private Const SC_MOVE = &HF010&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_NEXTWINDOW = &HF040&
Private Const SC_PREVWINDOW = &HF050&
Private Const SC_CLOSE = &HF060&
Private Const SC_VSCROLL = &HF070&
Private Const SC_HSCROLL = &HF080&
Private Const SC_MOUSEMENU = &HF090&
Private Const SC_KEYMENU = &HF100&
Private Const SC_ARRANGE = &HF110&
Private Const SC_RESTORE = &HF120&
Private Const SC_TASKLIST = &HF130&
Private Const SC_SCREENSAVE = &HF140&
Private Const SC_HOTKEY = &HF150&
Private Const SC_ICON = SC_MINIMIZE
Private Const SC_ZOOM = SC_MAXIMIZE
Private Const WM_MOVE = &H3

Dim temp As Long
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_EXSTYLE = -20

Public Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public procOld As Long
Private Const WM_SYSCOMMAND = &H112
Dim m
Private Const WM_USER = &H400
Function WindowProc(ByVal hw As Long, ByVal uMsg As _
Long, ByVal wParam As Long, ByVal lParam As Long) As _
Long
   
   Select Case uMsg
      Case WM_SYSCOMMAND   'alege o comanda din menu system
           Select Case wParam
            Case SC_CLOSE
               m = MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Quittin' ?")
               If m = vbNo Then Exit Function
         End Select
     End Select
   WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
Public Sub Hook()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
   temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub
