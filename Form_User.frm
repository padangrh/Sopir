VERSION 5.00
Begin VB.Form Form_User 
   BackColor       =   &H80000007&
   Caption         =   "User Manager"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7665
   Icon            =   "Form_User.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_delete 
      BackColor       =   &H008080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton btn_back 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton btn_add 
      BackColor       =   &H0080FF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   375
   End
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton btn_save 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ComboBox cb_status 
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
      Height          =   450
      ItemData        =   "Form_User.frx":628A
      Left            =   3840
      List            =   "Form_User.frx":6297
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2760
      Width           =   3015
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
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
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
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ListBox list_user 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3660
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   1920
      X2              =   1920
      Y1              =   960
      Y2              =   5160
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "User Manager"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "Form_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsuser As ADODB.Recordset

Private Sub btn_add_Click()
    con.Execute ("insert into user values('New User', '', 0)")
    refresh_list
    list_user.ListIndex = list_user.ListCount - 1
End Sub

Private Sub btn_back_Click()
    Form_Control.Show
    Unload Me
End Sub

Private Sub btn_cancel_Click()
    reset
End Sub

Private Sub btn_delete_Click()
    con.Execute ("delete from user where username = '" & rsuser!username & "'")
    refresh_list
    reset
End Sub

Private Sub btn_save_Click()
    If txt_username = "" Then
        MsgBox "Username tidak boleh kosong"
        Exit Sub
    End If
    
    If Len(txt_password) < 4 Then
        MsgBox "Password terlalu pendek"
        Exit Sub
    End If
    
    con.Execute ("update user set username = '" & txt_username & "', password = '" & txt_password & "', status = '" & cb_status.ListIndex & "' where username='" & rsuser!username & "'")
    refresh_list
    reset
End Sub

Private Sub Form_Load()
    refresh_list
    reset
End Sub

Private Sub refresh_list()
    list_user.Clear
    btn_add.Enabled = True
    Set rsuser = con.Execute("select * from user")
    rsuser.MoveFirst
    Do While Not rsuser.EOF
        list_user.AddItem (rsuser!username)
        If (rsuser!username) = "New User" Then
            btn_add.Enabled = False
        End If
        rsuser.MoveNext
    Loop
End Sub

Private Sub setUser(name As String)
    rsuser.MoveFirst
    Do While Not rsuser.EOF
        If rsuser!username = name Then
            Exit Do
        End If
        rsuser.MoveNext
    Loop
End Sub

Private Sub reset()
    txt_username = ""
    txt_password = ""
    cb_status.ListIndex = 1
    list_user.ListIndex = -1
    btn_save.Enabled = False
End Sub

Private Sub list_user_Click()
    If list_user.ListIndex = -1 Then
        Exit Sub
    End If
    
    btn_save.Enabled = True
    setUser (list_user.Text)
    txt_username = rsuser!username
    txt_password = rsuser!Password
    cb_status.ListIndex = rsuser!status
End Sub
