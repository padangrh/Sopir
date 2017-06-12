VERSION 5.00
Begin VB.Form Form_PIN 
   BackColor       =   &H80000012&
   Caption         =   "Security"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3945
   Icon            =   "Form_PIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_cancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton btn_ok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txt_pin 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Input PIN"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form_PIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pass As String

Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub btn_ok_Click()
    If txt_pin = pass Then
        Form_Claim.confirm_claim
        Unload Me
    Else
        MsgBox "PIN Salah"
    End If
End Sub

Private Sub txt_pin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        btn_ok_Click
    End If
End Sub
