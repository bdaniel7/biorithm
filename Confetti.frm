VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Form_Click()
   Dim CX, CY, Msg, XPos, YPos   ' Declare variables.
   ScaleMode = 3   ' Set ScaleMode to
         ' pixels.
   DrawWidth = 5   ' Set DrawWidth.
   ForeColor = QBColor(4)   ' Set foreground to red.
   FontSize = 24   ' Set point size.
   CX = ScaleWidth / 2   ' Get horizontal center.
   CY = ScaleHeight / 2   ' Get vertical center.
   Cls   ' Clear form.
   Msg = "Happy New Year!"
   CurrentX = CX - TextWidth(Msg) / 2   ' Horizontal position.
   CurrentY = CY - TextHeight(Msg)   ' Vertical position.
   Print Msg   ' Print message.
   Do
      XPos = Rnd * ScaleWidth   ' Get horizontal position.
      YPos = Rnd * ScaleHeight   ' Get vertical position.
      PSet (XPos, YPos), QBColor(Rnd * 15)   ' Draw confetti.
      DoEvents   ' Yield to other
   Loop   ' processing.
End Sub

