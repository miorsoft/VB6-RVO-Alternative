VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "fMain"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   13125
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    MAINLOOP
End Sub
 
 
Private Sub Form_Click()
    MODE = MODE + 1
    If MODE > 4 Then MODE = 0
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels

    WorldW = Me.ScaleWidth
    WorldH = Me.ScaleHeight

    Init_RVO 200 * 2 * 1.3        '1.25

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If FOLLOW = 0 Then FOLLOWworst: Exit Sub
        If FOLLOW > 1 Then FOLLOW = 1: Exit Sub
        If FOLLOW = 1 Then FOLLOW = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    doloop = False

End Sub
