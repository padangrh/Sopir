VERSION 5.00
Begin VB.Form Form_Navi 
   BackColor       =   &H80000007&
   Caption         =   "Main Menu"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7065
   Icon            =   "Form_Navi.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7185
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_logout 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton btn_control 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton btn_claim 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Claim Fee"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton btn_kunjungan 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kunjungan"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   720
      X2              =   1680
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   720
      X2              =   1680
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   720
      X2              =   1680
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   720
      X2              =   1680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   720
      X2              =   720
      Y1              =   1200
      Y2              =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "MAIN MENU"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   -360
      TabIndex        =   4
      Top             =   360
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7080
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Form_Navi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_claim_Click()
    Form_Claim.Show
    Unload Me
End Sub

Private Sub btn_control_Click()
    Form_Control.Show
    Unload Me
End Sub

Private Sub btn_kunjungan_Click()
    Form_Entry.Show
    Unload Me
End Sub

Private Sub btn_logout_Click()
    Form_Login.Show
    Unload Me
End Sub

Private Sub Form_Load()
    If userstatus = 0 Then
        btn_control.Enabled = False
    ElseIf userstatus = 1 Then
        btn_kunjungan.Enabled = False
        btn_claim.Enabled = False
    End If
End Sub
