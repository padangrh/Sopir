VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Laporan 
   BackColor       =   &H80000007&
   Caption         =   "Laporan"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
   Icon            =   "Form_Laporan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Semua Data"
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
      Height          =   405
      Left            =   6120
      TabIndex        =   10
      Top             =   1470
      Width           =   1575
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   6360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btn_status 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1095
      Left            =   4200
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton btn_kunjungan 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kunjungan"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton btn_member 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton btn_pengeluaran 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pengeluaran"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483641
      CalendarForeColor=   12648447
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483624
      Format          =   116785153
      CurrentDate     =   42613
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483641
      CalendarForeColor=   12648447
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483624
      Format          =   116785153
      CurrentDate     =   42613
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "~"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Tanggal"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Laporan"
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
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "Form_Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_back_Click()
    Form_Control.Show
    Unload Me
End Sub

Private Sub btn_kunjungan_Click()
    Call openReport("laporan_kunjungan.rpt", "kunjungan", DTPicker1.Enabled)
End Sub

Private Sub btn_member_Click()
    Call openReport("laporan_member.rpt", "member", False)
End Sub

Private Sub btn_pengeluaran_Click()
    Call openReport("laporan_pengeluaran.rpt", "claim", DTPicker1.Enabled)
End Sub

Private Sub openReport(filename As String, tablename As String, use_date As Boolean)
    pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=sopir"
    Cr.connect = pass
    Cr.ReportFileName = App.Path & "\" + filename
    Cr.WindowState = crptMaximized
    If use_date Then
        Cr.SelectionFormula = "{" + tablename + ".tanggal} >= #" & Format(DTPicker1.Value, "yyyy-mm-dd") & "# and {" + tablename + ".tanggal} <= #" & Format(DTPicker2.Value, "yyyy-mm-dd") & "# "
        Cr.Formulas(0) = "tgl1='" & Format(DTPicker1.Value, "dd mmmm yyyy") & "'"
        Cr.Formulas(1) = "tgl2='" & Format(DTPicker2.Value, "dd mmmm yyyy") & "'"
    End If
    Cr.RetrieveDataFiles
    Cr.Action = 1
    Cr.reset
End Sub

Private Sub btn_status_Click()
    Call openReport("laporan_status.rpt", "status_update", DTPicker1.Enabled)
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = DateAdd("d", -30, Date)
    DTPicker2.Value = Date
End Sub
