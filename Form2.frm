VERSION 5.00
Begin VB.Form Form_Login 
   BackColor       =   &H00000000&
   Caption         =   "Login"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7815
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_exit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton btn_login 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txt_password 
      BackColor       =   &H00C0FFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox txt_username 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000B&
      BorderWidth     =   3
      X1              =   600
      X2              =   600
      Y1              =   960
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "CHIP Member Manager"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_exit_Click()
    Unload Me
End Sub

Private Sub btn_login_Click()
    Dim rsuser As ADODB.Recordset
    Set rsuser = con.Execute("select * from user where LOWER(username)='" & LCase(txt_username) & "'")
    If rsuser.EOF Or rsuser.BOF Then
        MsgBox "user tidak ditemukan"
        Exit Sub
    Else
        If rsuser!Password = txt_password Then
            username = UCase(txt_username)
            userstatus = rsuser!status
            Form_Navi.Show
            Unload Me
        Else
            MsgBox "password salah"
            Exit Sub
        End If
    End If
        
End Sub

Private Sub Form_Load()
    connect
End Sub

Private Sub txt_password_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       btn_login_Click
    End If
    
End Sub
