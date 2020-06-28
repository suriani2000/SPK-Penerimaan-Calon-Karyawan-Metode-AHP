VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form3"
   ScaleHeight     =   9165
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat Penilaian"
      Height          =   615
      Left            =   6480
      TabIndex        =   21
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   2520
      TabIndex        =   19
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PILIH"
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7200
      Top             =   5040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      UserName        =   ""
      Password        =   ""
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
   Begin VB.CommandButton CMDKELUAR 
      Caption         =   "KELUAR"
      Height          =   615
      Left            =   10680
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton CMDHAPUS 
      Caption         =   "HAPUS"
      Height          =   615
      Left            =   8760
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton CMDSIMPAN 
      Caption         =   "SIMPAN"
      Height          =   615
      Left            =   9720
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Id Penilaian"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Nama Pelamar"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "K4"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "K3"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "K2"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "K1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "TGL PENILAIAN"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "KODE PELAMAR"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT DATA PENILAIAN"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form7.Show
End Sub
Private Sub CMDBATAL_Click()
Call Hapus_Teks
CMDSIMPAN.Caption = "Simpan"
End Sub

Private Sub CMDCARI_Click()
Form10.Show
End Sub

Private Sub CMDHAPUS_Click()
If Trim(TXTIDPASIEN.Text) = "" Then
    MsgBox "ID Pelamar Tidak Boleh Kosong", vbInformation, "Informasi"
    CMDPASIEN.SetFocus
Else
Dim psn As String
psn = MsgBox("Yakin Data Akan Dihapus ?", vbQuestion + vbYesNo, "Peringatan")
If psn = vbYes Then
    Call Hapus_Data
    CMDSIMPAN.Caption = "Simpan"
Else
    Call Hapus_Teks
    CMDSIMPAN.Caption = "Simpan"
End If
End If
End Sub

Private Sub Hapus_Data()
koneksiDB.Execute "delete from tbl_penilaian where id_penilaian='" & Text1.Text & "'"
Call Hapus_Teks
MsgBox "Data Telah Terhapus", vbInformation, "Pesan"
Call Segarkan
End Sub

Private Sub CMDKELUAR_Click()
Unload Me
End Sub

Private Sub CMDSIMPAN_Click()
Dim psn As String
psn = MsgBox("Yakin Data Akan Disimpan ?", vbQuestion + vbYesNo, "Pesan")
If psn = vbYes Then
    Select Case CMDSIMPAN.Caption
    Case "Simpan"
        Call Simpan_Data
    Case "Update"
        Call Simpan_Ulang
    End Select
    Call Hapus_Teks
    CMDSIMPAN.Caption = "Simpan"
Else
    MsgBox "Data Belum Disimpan.", vbInformation, "Informasi"
    Command2.SetFocus
End If
End Sub

Private Sub Simpan_Data()
If Trim(Text2.Text) = "" Then
    MsgBox "ID Pelamar tidak boleh kosong", vbInformation, "Informasi"
    CMDPASIEN.SetFocus
Else
Set RsCari = New ADODB.Recordset
RsCari.Open "select * from tbl_penilaian where idpenilaian='" & Text1.Text & "'", koneksiDB
    If Not RsCari.EOF Then
        MsgBox "Data " & Text1 & " Sudah Ada", vbCritical, "Pesan"
        CMDPASIEN.SetFocus
    Else
        koneksiDB.Execute "insert into tbl_penilaian (idpenilaian,kd_pelamar,tglpenilaian,k1,k2,k3,k4) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "')"
        Call Hapus_Teks
        MsgBox "Data Tersimpan", vbInformation, "Pesan"
        Call Segarkan
    End If
End If
End Sub

Private Sub Simpan_Ulang()
If Trim(Text2.Text) = "" Then
    MsgBox "ID pelamar tidak boleh kosong", vbInformation, "Informasi"
    CMDPASIEN.SetFocus
Else
koneksiDB.Execute " update tbl_penilaian set kd_pelamar ='" & Text2.Text & "',tglpenilaian ='" & Text4.Text & "',k1 ='" & Text5.Text & "',k2 ='" & Text6.Text & "',k3 ='" & Text7.Text & "',k4 ='" & Tetxt8.Text & "'where idpenilaian='" & txt1.Text & "'"
Call Hapus_Teks
MsgBox "Data Sudah Diubah", vbOKOnly, "Pesan"
Call Segarkan
End If
End Sub

Private Sub Command1_Click()
Form8.Show

End Sub

Private Sub Form_Load()
Call bukadb
End Sub

Private Sub Hapus_Teks()
Dim Control
For Each Control In Me.Controls
If TypeOf Control Is TextBox Then
Control.Text = ""
End If
If TypeOf Control Is ComboBox Then
Control.Text = "- Pilih -"
End If
Next Control
End Sub

Private Sub Segarkan()
Call bukadb
'Call Tampilkan_Dosen
'Set DataGrid1.DataSource = Adodc1
'With DataGrid1
'End With
'Call Atur_Grid
End Sub


