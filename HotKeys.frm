VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Please press ""ctrl+g+r"" in any combination (e.g ""r+g+ctrl"")"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I saw this submition
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=59841&lngWId=1
'And this came into my mind, i know this is not some great code, but it worxXx (the hotkeys work, in any combination)
Dim Key(1 To 3) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Key(1) = Ctrl
'Key(2) = G
'Key(3) = R

Select Case KeyCode
Case vbKeyControl
    Key(1) = True
Case vbKeyG
    Key(2) = True
Case vbKeyR
    Key(3) = True
End Select

If Key(1) = True And Key(2) = True And Key(3) = True Then
    MsgBox "Yeh you did it!"
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key(1) = False
Key(2) = False
Key(3) = False
End Sub

Private Sub Form_Load()
Key(1) = False
Key(2) = False
Key(3) = False
End Sub
