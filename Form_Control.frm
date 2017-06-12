VERSION 5.00
Begin VB.Form Form_Control 
   BackColor       =   &H80000007&
   Caption         =   "Control Panel"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6165
   Icon            =   "Form_Control.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5520
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "< Back"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Laporan"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tarif"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Control Panel"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form_User.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Form_Tarif.Show
    Unload Me
End Sub

Private Sub Command3_Click()
    Form_Member.Show
    Unload Me
End Sub

Private Sub Command4_Click()
    Form_Laporan.Show
    Unload Me
End Sub

Private Sub Command5_Click()
    Form_Navi.Show
    Unload Me
End Sub
