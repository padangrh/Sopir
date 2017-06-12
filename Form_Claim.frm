VERSION 5.00
Begin VB.Form Form_Claim 
   BackColor       =   &H80000007&
   Caption         =   "Claim Fee"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13830
   Icon            =   "Form_Claim.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   7920
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   9600
      Top             =   840
   End
   Begin VB.CommandButton btn_claim 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Claim"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txt_tambah 
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
      Height          =   390
      Left            =   9480
      TabIndex        =   23
      Text            =   "10.000"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txt_id 
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
      Height          =   390
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label txt_nomor 
      BackColor       =   &H80000007&
      Caption         =   "00001"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   31
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Nomor"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   30
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lbl_kunjungan 
      BackColor       =   &H80000007&
      Caption         =   "5 kali"
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
      Left            =   10920
      TabIndex        =   27
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lbl_grandtotal 
      BackColor       =   &H80000007&
      Caption         =   "500.000"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10560
      TabIndex        =   26
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000007&
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   25
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000007&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      Caption         =   "Tambahan :"
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
      Left            =   7800
      TabIndex        =   22
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   7920
   End
   Begin VB.Label lbl_bonus 
      BackColor       =   &H80000007&
      Caption         =   "50.000"
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
      Left            =   9000
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Bonus :"
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
      Left            =   7800
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   "Total Kunjungan"
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
      Left            =   7800
      TabIndex        =   19
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lbl_total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "500.000"
      BeginProperty Font 
         Name            =   "Geometr706 BlkCn BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      X1              =   1080
      X2              =   5880
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lbl_dus9 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   17
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000007&
      Caption         =   "Dus 9"
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
      Left            =   1080
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lbl_dus7 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   15
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      Caption         =   "Dus 7"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lbl_dus5 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "Dus 5"
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
      Left            =   1080
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbl_dus4 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   11
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Caption         =   "Dus 4"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lbl_dus3 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   9
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Dus 3"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl_dus2 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Dus 2"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbl_dus1 
      BackColor       =   &H80000007&
      Caption         =   "10 x 15.000 = 150.000"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Dus 1"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbl_username 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "Richard"
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
      Left            =   9840
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lbl_time 
      BackColor       =   &H80000007&
      Caption         =   "Wednesday, 24 August 2016 (17:13:20)"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   14760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Member ID"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form_Claim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_validation As Boolean
Dim rsmember As ADODB.Recordset
Dim rstarif As ADODB.Recordset
Dim grandtotal As Long
Dim time_count As Integer
Dim claim_bonus As Boolean
Dim semua_kunjungan As String

Private Sub btn_back_Click()
    Form_Navi.Show
    Unload Me
End Sub

Private Sub btn_claim_Click()
    Dim final As Long
    final = grandtotal + Val(txt_tambah)
    If final = 0 Then
        MsgBox "Tidak ada fee yang bisa diclaim"
        Exit Sub
    End If
    
    confirm_claim
End Sub

Public Sub confirm_claim()
    grandtotal = grandtotal + Val(txt_tambah)
    
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
    Printer.Print Tab(3); "Operator: "; lbl_username.Caption;
    Printer.Print Tab(3); "Nomor Faktur: "; txt_nomor;
    Printer.Print Tab(3); Format(Now, "dd mmmm yyyy hh:mm:ss");
    Printer.Print Tab(3); "-------------------------------------------------";
    Printer.FontSize = 15
    Printer.Print Tab(3); "Total Member Fee: ";
    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.Print Tab(8); Format(grandtotal, "###,###,##0");
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.Print Tab(3); "                                                "
    Printer.Print Tab(3); "                                                "
    Printer.Print Tab(3); "Diterima Oleh"; Tab(27); "Dibayar Oleh";
    Printer.Print Tab(3); "                                                "
    Printer.Print Tab(3); "                                                "
    Printer.Print Tab(3); "                                                "
    Printer.Print Tab(3); "_____________"; Tab(27); "____________";
    Printer.EndDoc
    
    con.Execute ("insert into claim values('" & Val(txt_nomor) & "', '" & rsmember!member_id & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & Format(Now, "hh:mm:ss") & "','" & semua_kunjungan & "', '" & priceToNum(lbl_bonus) & "', '" & Val(txt_tambah) & "', '" & priceToNum(lbl_total) & "', '" & username & "'  )")
    con.Execute ("update kunjungan set claimed=1 where member_id='" & rsmember!member_id & "' and kunjungan_id<>'" & rsmember!last_visit & "' and claimed=0")
    If claim_bonus Then
        con.Execute ("update kunjungan set bonus=1 where member_id='" & rsmember!member_id & "' and bonus=0 limit 10")
    End If
    txt_nomor = Format(Val(txt_nomor) + 1, String(5, "0"))
    reset
End Sub

Private Sub Form_Load()
    grandtotal = 0
    claim_bonus = False
    id_validation = False
    semua_kunjungan = ""
    lbl_username = username
    Set rstarif = con.Execute("select * from tarif")
    time_count = 0
    Set Rec = con.Execute("select max(claim.claim_id) AS nomor From claim")
    If IsNull(Rec!nomor) = True Then
       txt_nomor = Format(1, String(5, "0"))
    Else
       txt_nomor = Format(Rec!nomor + 1, String(5, "0"))
    End If
    
    reset
End Sub

Private Function getTarif(nama As String) As Long
    rstarif.MoveFirst
    Do While Not rstarif.EOF
        If rstarif!nama = nama Then
            getTarif = rstarif!Value
            Exit Function
        End If
        rstarif.MoveNext
    Loop
    
    getTarif = 0
End Function

Private Sub txt_id_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_tambah.Enabled = True
        btn_claim.Enabled = True
        
        Set rsmember = con.Execute("select * from member where member_id = '" & txt_id.Text & "'")
        If (rsmember.EOF Or rsmember.BOF) Then
            MsgBox "Member id tidak valid", vbOKOnly, "Warning"
            Exit Sub
        ElseIf rsmember!status = 0 Then
            MsgBox "Member tidak aktif", vbOKOnly, "Warning"
            Exit Sub
        Else
            txt_tambah.SetFocus
            lbl_kunjungan = CStr(rsmember!kunjungan) + " kali"
            Dim rsKunjungan As ADODB.Recordset
            Set rsKunjungan = con.Execute("select * from kunjungan where member_id = '" & txt_id.Text & "' and claimed = 0")
            If rsKunjungan.EOF Or rsKunjungan.BOF Then
                Exit Sub
            Else
                rsKunjungan.MoveFirst
            End If
            
            Dim total As Long
            Dim bonus_count As Integer
            total = 0
            bonus_count = 0
            Dim jumlah_dus(10) As Integer
            Do While Not rsKunjungan.EOF
                If rsKunjungan!kunjungan_id <> rsmember!last_visit Then
                    jumlah_dus(1) = jumlah_dus(1) + rsKunjungan!dus1
                    jumlah_dus(2) = jumlah_dus(2) + rsKunjungan!dus2
                    jumlah_dus(3) = jumlah_dus(3) + rsKunjungan!dus3
                    jumlah_dus(4) = jumlah_dus(4) + rsKunjungan!dus4
                    jumlah_dus(5) = jumlah_dus(5) + rsKunjungan!dus5
                    jumlah_dus(7) = jumlah_dus(7) + rsKunjungan!dus7
                    jumlah_dus(9) = jumlah_dus(9) + rsKunjungan!dus9
                    
                    semua_kunjungan = semua_kunjungan + CStr(rsKunjungan!kunjungan_id) + ","
                End If
            
                rsKunjungan.MoveNext
            Loop
            
            Dim sub_total As Long
            sub_total = jumlah_dus(1) * getTarif("dus1")
            total = total + sub_total
            lbl_dus1 = CStr(jumlah_dus(1)) + " x " + CStr(Format(getTarif("dus1"), "###,###,##0")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(2) * getTarif("dus2")
            total = total + sub_total
            lbl_dus2 = CStr(jumlah_dus(2)) + " x " + CStr(Format(getTarif("dus2"), "###,###,##0")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(3) * getTarif("dus3")
            total = total + sub_total
            lbl_dus3 = CStr(jumlah_dus(3)) + " x " + CStr(Format(getTarif("dus3"), "###,###,###")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(4) * getTarif("dus4")
            total = total + sub_total
            lbl_dus4 = CStr(jumlah_dus(4)) + " x " + CStr(Format(getTarif("dus4"), "###,###,###")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(5) * getTarif("dus5")
            total = total + sub_total
            lbl_dus5 = CStr(jumlah_dus(5)) + " x " + CStr(Format(getTarif("dus5"), "###,###,###")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(7) * getTarif("dus7")
            total = total + sub_total
            lbl_dus7 = CStr(jumlah_dus(7)) + " x " + CStr(Format(getTarif("dus7"), "###,###,###")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            
            sub_total = jumlah_dus(9) * getTarif("dus9")
            total = total + sub_total
            lbl_dus9 = CStr(jumlah_dus(9)) + " x " + CStr(Format(getTarif("dus9"), "###,###,###")) + " = " + CStr(Format(sub_total, "###,###,##0"))
            lbl_total = Format(total, "###,###,##0")
            
            claim_bonus = False
            Set rsKunjungan = con.Execute("select * from kunjungan where member_id = '" & txt_id.Text & "' and bonus = 0")
            If Not rsKunjungan.EOF Then
                rsKunjungan.MoveFirst
                Do While Not rsKunjungan.EOF
                    If rsKunjungan!bonus = 0 Then
                        bonus_count = bonus_count + 1
                        If (bonus_count = 10) Then
                            claim_bonus = True
                        End If
                    End If
                    rsKunjungan.MoveNext
                Loop
            End If
            
            Dim bonus_nominal As Long
            If claim_bonus Then
                bonus_nominal = getTarif("bonus")
            Else
                bonus_nominal = 0
            End If
            
            lbl_bonus = Format(bonus_nominal, "###,###,##0")
            grandtotal = bonus_nominal + total
            lbl_grandtotal = Format(grandtotal, "###,###,##0")
            
        End If
    End If
End Sub

Private Sub reset()
    semua_kunjungan = ""
    txt_id = ""
    lbl_dus1 = ""
    lbl_dus2 = ""
    lbl_dus3 = ""
    lbl_dus4 = ""
    lbl_dus5 = ""
    lbl_dus7 = ""
    lbl_dus9 = ""
    
    lbl_kunjungan = "0 kali"
    lbl_bonus = "0"
    txt_tambah = "0"
    lbl_total = "0"
    lbl_grandtotal = "0"
    txt_tambah.Enabled = False
    btn_claim.Enabled = False
    claim_bonus = False
    
    If time_count > 0 Then
        txt_id.SetFocus
    End If
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

Private Sub txt_tambah_Change()
    Dim new_total As Long
    new_total = grandtotal + Val(txt_tambah)
    lbl_grandtotal = Format(new_total, "###,###,##0")
End Sub

Private Sub txt_tambah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        btn_claim_Click
    End If
End Sub

Private Function priceToNum(price As String) As Long
    price = Replace(price, ",", "")
    price = Replace(price, ".", "")
    priceToNum = Val(price)
End Function
