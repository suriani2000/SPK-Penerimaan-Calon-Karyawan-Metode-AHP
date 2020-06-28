VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   ScaleHeight     =   8925
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "DATAR  HASIL CALON PEGAWAI"
      Height          =   615
      Left            =   5880
      TabIndex        =   24
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PILIH"
      Height          =   615
      Left            =   3600
      TabIndex        =   23
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PILIH"
      Height          =   495
      Left            =   3480
      TabIndex        =   22
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   2160
      TabIndex        =   21
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   2160
      TabIndex        =   20
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   2160
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2160
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2160
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
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
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CMDHAPUS 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton CMDSIMPAN 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "K4"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "K3"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "k1"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "k2"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Pelamar"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "ID Pelamar"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "KETERANGAN HASIL"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "NILAI AKHIR"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ID PENILAIAN"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT HASIL PENILAIAN"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form9.Show
End Sub
Private Sub CMDBATAL_Click()
Call Hapus_Teks
CMDSIMPAN.Caption = "Simpan"
End Sub

Private Sub CMDCARI_Click()
Form10.Show
End Sub

Private Sub CMDHAPUS_Click()
If Trim(Text1.Text) = "" Then
    MsgBox "ID hasil Tidak Boleh Kosong", vbInformation, "Informasi"
    Command1.SetFocus
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
koneksiDB.Execute "delete from tbl_hasilpenilaian where idpenilaian='" & Text1.Text & "'"
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
    Command1.SetFocus
End If
End Sub

Private Sub Simpan_Data()
If Trim(Text2.Text) = "" Then
    MsgBox "ID Pelamar tidak boleh kosong", vbInformation, "Informasi"
    CMDPASIEN.SetFocus
Else
Set RsCari = New ADODB.Recordset
RsCari.Open "select * from tbl_hasilpenilaian where idpenilaian='" & Text1.Text & "'", koneksiDB
    If Not RsCari.EOF Then
        MsgBox "Data " & Text1 & " Sudah Ada", vbCritical, "Pesan"
        Command1.SetFocus
    Else
        koneksiDB.Execute "insert into tbl_hasilpenilaian (idpenilaian,nilaiakhir,keteranganhasil) value ('" & Text1.Text & "','" & Text8.Text & "','" & Text9.Text & "')"
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
koneksiDB.Execute " update tbl_hasilpenilaian set nilaiakhir ='" & Text8.Text & "',ketranganhasil ='" & Text9.Text & "' where idpenilaian='" & txt1.Text & "'"
Call Hapus_Teks
MsgBox "Data Sudah Diubah", vbOKOnly, "Pesan"
Call Segarkan
End If
End Sub

Private Sub Command1_Click()
Form10.Show

End Sub

Private Sub Command3_Click()
Form11.Show

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



Private Sub Text8_Change()
Text8.Text = Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)


End Sub
