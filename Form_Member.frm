VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Member 
   BackColor       =   &H80000007&
   Caption         =   "Member Manager"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13575
   Icon            =   "Form_Member.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_reset 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reset"
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton btn_cari 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cari"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt_filter 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   4095
   End
   Begin VB.ComboBox cb_field 
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
      ItemData        =   "Form_Member.frx":628A
      Left            =   1200
      List            =   "Form_Member.frx":6297
      TabIndex        =   4
      Text            =   "Field"
      Top             =   1320
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form_Member.frx":62B1
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   -2147483625
      ForeColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12000
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "yuyu"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   -120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   13560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Member Manager"
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
      TabIndex        =   1
      Top             =   360
      Width           =   13455
   End
End
Attribute VB_Name = "Form_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_back_Click()
    Form_Control.Show
    Unload Me
End Sub

Private Sub btn_cari_Click()
    If cb_field.ListIndex = -1 Or txt_filter = "" Then
        MsgBox "Filter tidak valid"
        Exit Sub
    End If
    
    If cb_field.ListIndex = 0 Then
       Adodc1.RecordSource = "select * from member where status <> 0 and member_id like '" & "%" + txt_filter + "%" & "'"
    ElseIf cb_field.ListIndex = 1 Then
       Adodc1.RecordSource = "select * from member where status <> 0 and nama like '" & "%" + txt_filter + "%" & "'"
    ElseIf cb_field.ListIndex = 2 Then
       Adodc1.RecordSource = "select * from member where status <> 0 and perusahaan like '" & "%" + txt_filter + "%" & "'"
    End If
    Adodc1.Refresh
End Sub

Private Sub btn_reset_Click()
    Adodc1.RecordSource = "select * from member where status <> 0"
    Adodc1.Refresh
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "DSN=sopir"
    Adodc1.RecordSource = "select * from member where status <> 0"
    Adodc1.Refresh
End Sub

