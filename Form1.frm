VERSION 5.00
Begin VB.Form Form_Entry 
   BackColor       =   &H00000000&
   Caption         =   "Member Sopir CHIP"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_back 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txt_pin2 
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
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   6000
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   38
      Text            =   "123456"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox txt_pin1 
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
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   6000
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   37
      Text            =   "123456"
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox txt_terakhir_tgl 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   35
      Text            =   "123456"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ComboBox cb_status 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      ItemData        =   "Form1.frx":628A
      Left            =   6000
      List            =   "Form1.frx":629D
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox txt_kunjungan 
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
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   33
      Text            =   "123456"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txt_terakhir 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   32
      Text            =   "123456"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txt_perusahaan 
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
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   31
      Text            =   "123456"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txt_hp 
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
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   30
      Text            =   "123456"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txt_nama 
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
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6000
      TabIndex        =   29
      Text            =   "123456"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txt_dus9 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   21
      Text            =   "123456"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_dus7 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   19
      Text            =   "123456"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txt_dus5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Text            =   "123456"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txt_dus4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Text            =   "123456"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txt_dus3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Text            =   "123456"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_dus2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10440
      TabIndex        =   11
      Text            =   "123456"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txt_dus1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10440
      TabIndex        =   10
      Text            =   "123456"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txt_nomor 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   12120
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   120
      Top             =   7080
   End
   Begin VB.CommandButton btn_save 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save - Print"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton btn_update 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Update Data"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txt_id 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton btn_new 
      BackColor       =   &H00C0FFFF&
      Caption         =   "New Member"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "PIN Number"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   9000
      Y1              =   960
      Y2              =   7680
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "Status"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "Total Kunjungan"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Kunjungan Terakhir"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Perusahaan"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Nomor HP"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Nama"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   3000
      X2              =   50000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Dus 9"
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
      Height          =   255
      Left            =   11760
      TabIndex        =   22
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Dus 7"
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
      Height          =   255
      Left            =   11760
      TabIndex        =   20
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Dus 5"
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
      Height          =   255
      Left            =   11760
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Dus 4"
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
      Height          =   255
      Left            =   9720
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Dus 3"
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
      Height          =   255
      Left            =   9720
      TabIndex        =   14
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Dus 2"
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
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Dus 1"
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
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Nomor Kunjungan"
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
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lbl_admin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Richard"
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
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lbl_time 
      BackColor       =   &H00000000&
      Caption         =   "Wed, 26/08/2016 15:23:16"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   7095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   3000
      X2              =   50000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   3000
      X2              =   50000
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Member ID :"
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
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   3000
      X2              =   3000
      Y1              =   0
      Y2              =   7680
   End
End
Attribute VB_Name = "Form_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_validation As Boolean
Dim rsmember As ADODB.Recordset
Dim time_count As Integer

Private Sub btn_back_Click()
    Form_Navi.Show
    Unload Me
End Sub

Private Sub btn_new_Click()
    If id_validation = False Then
        MsgBox "Member id tidak valid"
        Exit Sub
    End If
    
    If rsmember.Fields("status") <> 0 Then
        MsgBox "Member telah aktif"
        Exit Sub
    End If
    
    If Len(txt_pin1) < 4 Then
        MsgBox "PIN terlalu pendek"
        txt_pin1 = ""
        txt_pin2 = ""
        txt_pin1.SetFocus
        Exit Sub
    End If
    
    If txt_pin1 <> txt_pin2 Then
        MsgBox "Konfirmasi PIN gagal, input PIN sekali lagi"
        txt_pin1 = ""
        txt_pin2 = ""
        txt_pin1.SetFocus
        Exit Sub
    End If
    
    If Val(txt_kunjungan) > 0 Then
        Dim count As Integer
        count = 0
        Do While count < Val(txt_kunjungan)
            con.Execute ("insert into kunjungan values('" & Val(txt_nomor) & "', '" & txt_id & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & Format(Now, "hh:mm:ss") & "', '" & username & "', 0,0,0,0,0,0,0,1,0)")
            txt_nomor = Format(Val(txt_nomor) + 1, String(5, "0"))
            count = count + 1
        Loop
    End If
    
    con.Execute ("update member set nama='" & txt_nama & "', phone='" & txt_hp & "', perusahaan='" & txt_perusahaan & "', kunjungan='" & Val(txt_kunjungan) & "', status=1, pin_number = '" & txt_pin1 & "' where member_id='" & txt_id & "'")
    MsgBox "member telah berhasil didaftarkan"
    reset
End Sub

Private Sub btn_save_Click()
    If id_validation = False Then
        MsgBox "Member id tidak valid"
        Exit Sub
    End If
    
    If rsmember.Fields("status") = 0 Then
        MsgBox "Member belum aktif"
        Exit Sub
    End If
    
    If MsgBox("Cetak bukti Kunjungan", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    con.Execute ("insert into kunjungan values('" & Val(txt_nomor) & "', '" & txt_id & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & Format(Now, "hh:mm:ss") & "', '" & lbl_admin.Caption & "', '" & Val(txt_dus1) & "', '" & Val(txt_dus2) & "', '" & Val(txt_dus3) & "', '" & Val(txt_dus4) & "', '" & Val(txt_dus5) & "', '" & Val(txt_dus7) & "', '" & Val(txt_dus9) & "', 0, 0)")
    Dim last_visit As Integer
    If Val(txt_dus1) + Val(txt_dus2) + Val(txt_dus3) + Val(txt_dus4) + Val(txt_dus5) + Val(txt_dus7) + Val(txt_dus9) > 0 Then
        last_visit = Val(txt_nomor)
    Else
        last_visit = rsmember!last_visit
    End If
    
    con.Execute ("update member set kunjungan = kunjungan + 1, last_visit = '" & last_visit & "' where member_id='" & txt_id & "'")
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print Tab(2); "Member Sopir CH Nipah";
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.Print Tab(3); "                                           "
    Printer.Print Tab(3); "-------------------------------------------------";
    Printer.Print Tab(3); "Nomor Member: "; rsmember.Fields("member_id");
    Printer.Print Tab(3); "Nama Member: "; rsmember.Fields("nama");
    Printer.Print Tab(3); "Operator: "; lbl_admin.Caption;
    Printer.Print Tab(3); "Total Kunjungan: "; CStr(Val(txt_kunjungan) + 1);
    Printer.Print Tab(3); Format(Now, "dd mmmm yyyy hh:mm:ss");
    Printer.Print Tab(3); "-------------------------------------------------";
    Printer.Print Tab(3); "Dus 1:  "; txt_dus1; Tab(25); "Dus 5:  "; txt_dus5;
    Printer.Print Tab(3); "Dus 2:  "; txt_dus2; Tab(25); "Dus 7:  "; txt_dus7;
    Printer.Print Tab(3); "Dus 3:  "; txt_dus3; Tab(25); "Dus 9:  "; txt_dus9;
    Printer.Print Tab(3); "Dus 4:  "; txt_dus4;
    Printer.Print Tab(3); "-------------------------------------------------";
    Printer.Print Tab(3); "Terima kasih telah membawa tamu ke"
    Printer.Print Tab(3); "    Kripik Balado Christine Hakim"
    Printer.EndDoc
    
    txt_nomor = Format(Val(txt_nomor) + 1, String(5, "0"))
    reset
End Sub

Private Sub btn_update_Click()
    If id_validation = False Then
        MsgBox "Member id tidak valid"
        Exit Sub
    End If
    
    If rsmember.Fields("status") = 0 Then
        MsgBox "Member belum aktif"
        Exit Sub
    End If
    
    If cb_status.ListIndex <> rsmember.Fields("status") Then
        Dim alasan As String
        alasan = InputBox("Alasan perubahan status member", "Status Change")
        If alasan = "" Then
            Exit Sub
        End If
        
        con.Execute ("insert into status_update values('" & txt_id & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & Format(Now, "hh:mmss") & "', '" & cb_status.Text & "', '" & alasan & "', '" & username & "')")
    End If
    
    If txt_pin1 = rsmember!pin_number Then
        If Len(txt_pin2) < 4 Then
            MsgBox "PIN terlalu pendek"
            txt_pin2 = ""
            txt_pin2.SetFocus
            Exit Sub
        End If
        con.Execute ("update member set pin_number='" & txt_pin2 & "' where member_id='" & txt_id & "'")
    End If
    
    If cb_status.ListIndex = 0 Then
        con.Execute ("update member set nama='', phone='', perusahaan='', last_visit=0, kunjungan = 0, status=0, pin_number='' where member_id='" & txt_id & "'")
    Else
        con.Execute ("update member set nama='" & txt_nama & "', phone='" & txt_hp & "', perusahaan='" & txt_perusahaan & "', status='" & cb_status.ListIndex & "' where member_id='" & txt_id & "'")
    End If
    MsgBox "Data member telah berhasil diubah"
    reset
End Sub

Private Sub Form_Load()
    time_count = 0
    lbl_admin.Caption = username
    Set Rec = con.Execute("select max(kunjungan.kunjungan_id) AS id From kunjungan")
    If IsNull(Rec!id) = True Then
       txt_nomor.Text = Format(1, String(5, "0"))
    Else
       txt_nomor.Text = Format(Rec!id + 1, String(5, "0"))
    End If
    
    reset
End Sub



Private Sub txt_id_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set rsmember = con.Execute("select * from member where member_id = '" & txt_id.Text & "'")
        If (rsmember.EOF Or rsmember.BOF) Then
            MsgBox "Member id tidak valid", vbOKOnly, "Warning"
            Exit Sub
        Else
            txt_nama.SetFocus
            txt_nama.Text = rsmember.Fields("nama")
            txt_hp.Text = rsmember.Fields("phone")
            txt_perusahaan.Text = rsmember.Fields("perusahaan")
            txt_terakhir.Text = "No. " + Format(rsmember.Fields("last_visit"), String(5, "0"))
            txt_kunjungan.Text = rsmember.Fields("kunjungan")
            cb_status.ListIndex = rsmember.Fields("status")
            id_validation = True
            
            Dim is_active As Boolean
            is_active = rsmember.Fields("status") > 0
            cb_status.Enabled = is_active
            txt_kunjungan.Enabled = Not is_active
            btn_new.Enabled = Not is_active
            btn_update.Enabled = is_active
            btn_save.Enabled = is_active
            
            Dim rsKunjungan As ADODB.Recordset
            Set rsKunjungan = con.Execute("select * from kunjungan where kunjungan_id = '" & rsmember!last_visit & "'")
            If Not (rsKunjungan.EOF Or rsKunjungan.BOF) Then
                txt_terakhir_tgl = Format(rsKunjungan!tanggal, "dd mmmm yyyy")
            End If
            
        End If
    End If
End Sub

Private Sub reset()
    If time_count > 0 Then
        txt_id.SetFocus
    End If
    id_validation = False
    txt_dus1.Text = 0
    txt_dus2.Text = 0
    txt_dus3.Text = 0
    txt_dus4.Text = 0
    txt_dus5.Text = 0
    txt_dus7.Text = 0
    txt_dus9.Text = 0
    
    txt_nama = ""
    txt_hp = ""
    txt_terakhir = ""
    txt_terakhir_tgl = ""
    txt_perusahaan = ""
    txt_kunjungan = ""
    txt_id = ""
    txt_pin1 = ""
    txt_pin2 = ""
    cb_status.ListIndex = -1
    
    btn_new.Enabled = False
    btn_update.Enabled = False
    btn_save.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If time_count = 2 Then
        txt_id.SetFocus
    End If
    
    If time_count <= 2 Then
        time_count = time_count + 1
    End If
    
    lbl_time = Format(Now, "dddd, dd mmmm yyyy (hh:mm:ss)")
End Sub
