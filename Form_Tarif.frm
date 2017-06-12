VERSION 5.00
Begin VB.Form Form_Tarif 
   BackColor       =   &H80000007&
   Caption         =   "Setting Tarif"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7440
   Icon            =   "Form_Tarif.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txt_value 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ListBox list_tarif 
      BackColor       =   &H80000007&
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
      Height          =   4680
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl_tarif_name 
      BackColor       =   &H80000007&
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Fee Setting"
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
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form_Tarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstarif As ADODB.Recordset

Private Sub btn_back_Click()
    Form_Control.Show
    Unload Me
End Sub

Private Sub btn_cancel_Click()
    reset
End Sub

Private Sub btn_save_Click()
    con.Execute ("update tarif set value = '" & Val(txt_value) & "' where nama = '" & rstarif!nama & "'")
    reset
End Sub

Private Sub Form_Load()
    Set rstarif = con.Execute("select * from tarif")
    rstarif.MoveFirst
    Do While Not rstarif.EOF
        list_tarif.AddItem (rstarif!nama)
        rstarif.MoveNext
    Loop
    reset
End Sub

Private Sub setTarif(nama As String)
    rstarif.MoveFirst
    Do While Not rstarif.EOF
        If (nama = rstarif!nama) Then
            Exit Do
        End If
        rstarif.MoveNext
    Loop
End Sub

Private Sub reset()
    lbl_tarif_name = ""
    txt_value = 0
    txt_value.Enabled = False
    btn_save.Enabled = False
    btn_cancel.Enabled = False
    list_tarif.ListIndex = -1
End Sub

Private Sub list_tarif_Click()
    If list_tarif.ListIndex = -1 Then
        Exit Sub
    End If
    
    txt_value.Enabled = True
    btn_save.Enabled = True
    btn_cancel.Enabled = True
    setTarif (list_tarif.Text)
    lbl_tarif_name = rstarif!nama
    txt_value = rstarif!Value
End Sub
