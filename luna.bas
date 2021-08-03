Attribute VB_Name = "Module1"
Option Explicit

Function MonthLength(Luna As Byte, Optional An As Integer) As Byte
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
MonthLength = intLen
End Function
