VERSION 5.00
Begin VB.Form Form_Background 
   BackColor       =   &H80000012&
   Caption         =   "Sopir Manager"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form_Background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Login.Show
End Sub

Private Sub Form_GetFocus()
   MsgBox "test"
End Sub
