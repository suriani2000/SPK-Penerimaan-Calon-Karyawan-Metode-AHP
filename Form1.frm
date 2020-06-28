VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDKELUAR 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   11400
      TabIndex        =   25
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   1200
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10800
      Top             =   6840
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2535
      Left            =   5040
      TabIndex        =   16
      Top             =   4080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.CommandButton CMDBATAL 
      Caption         =   "BATAL"
      Height          =   735
      Left            =   11400
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton CMDHAPUS 
      Caption         =   "HAPUS"
      Height          =   735
      Left            =   11400
      TabIndex        =   14
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton CMDSIMPAN 
      Caption         =   "SIMPAN"
      Height          =   735
      Left            =   11400
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   525
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "TGLLAMARAN"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "IPK"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "PENDIDIKAN"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "STATUS"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "JENIS KELAMIN"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "TELPON"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "ALAMAT"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   " TTL"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "NAMA LENGKAP"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "KODE PELAMAR"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT DATA PELAMAR"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsCari As New ADODB.Recordset

Private Sub CMDBATAL_Click()
Call Hapus_Teks
Text1.SetFocus
CMDSIMPAN.Caption = "Simpan"
End Sub

Private Sub CMDHAPUS_Click()
If Trim(Text1.Text) = "" Then
    MsgBox "Kode Pelamar Tidak Boleh Kosong", vbInformation, "Informasi"
    Text1.SetFocus
Else
Dim psn As String
psn = MsgBox("Yakin Data Akan Dihapus ?", vbQuestion + vbYesNo, "Peringatan")
If psn = vbYes Then
    Call Hapus_Data
    Text1.SetFocus
    CMDSIMPAN.Caption = "Simpan"
Else
    Call Hapus_Teks
    Text1.SetFocus
    CMDSIMPAN.Caption = "Simpan"
End If
End If
End Sub

Private Sub Hapus_Data()
koneksiDB.Execute "delete from tbl_pelamar where kd_pelamar='" & Text1.Text & "'"
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
    Text1.SetFocus
Else
    MsgBox "Data Belum Disimpan.", vbInformation, "Informasi"
    Text1.SetFocus
End If
End Sub

Private Sub Simpan_Data()
If Trim(Text1.Text) = "" Then
    MsgBox "Kode Pelamar tidak boleh kosong", vbInformation, "Informasi"
    Text1.SetFocus
Else
Set RsCari = New ADODB.Recordset
RsCari.Open "select * from tbl_pelamar where kd_pelamar='" & Text1.Text & "'", koneksiDB
    If Not RsCari.EOF Then
        MsgBox "Data " & Text1 & " Sudah Ada", vbCritical, "Pesan"
        Text1.Text = ""
        Text1.SetFocus
    Else
        koneksiDB.Execute "insert into tbl_pelamar (kd_pelamar,nm_karyawan,ttl,alamat,telpon,jenis_kelamin,status,pendidikan,ipk,tgl_lamaran) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "')"
        Call Hapus_Teks
        MsgBox "Data Tersimpan", vbInformation, "Pesan"
        Call Segarkan
    End If
End If
End Sub

Private Sub Simpan_Ulang()
If Trim(Text1.Text) = "" Then
    MsgBox "ID Pegawai tidak boleh kosong", vbInformation, "Informasi"
    Text1.SetFocus
Else
koneksiDB.Execute " update tbl_pelamar set nm_karyawan ='" & Text2.Text & "',ttl ='" & Text3.Text & "',alamat ='" & Text4.Text & "',telpon ='" & Text5.Text & "',jenis_kelamin ='" & Text6.Text & "',status ='" & Text7.Text & "',pendidikan ='" & Text8.Text & "',ipk ='" & Text9.Text & "',tgl_melamar='" & Text10.Text & "' where kd_pelamar='" & Text1.Text & "'"
Call Hapus_Teks
MsgBox "Data Sudah Diubah", vbOKOnly, "Pesan"
Call Segarkan
End If
End Sub

Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = DataGrid1.Columns(0)
Text2.Text = DataGrid1.Columns(1)
Text3.Text = DataGrid1.Columns(2)
Text4.Text = DataGrid1.Columns(3)
Text5.Text = DataGrid1.Columns(4)
Text6.Text = DataGrid1.Columns(5)
Text7.Text = DataGrid1.Columns(6)
Text8.Text = DataGrid1.Columns(7)
Text9.Text = DataGrid1.Columns(8)
Text10.Text = DataGrid1.Columns(9)
CMDSIMPAN.Caption = "Update"
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Tampilkan_pelamar
Atur_Grid
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

Sub Tampilkan_pelamar()
Call bukadb
Adodc1.ConnectionString = "DRIVER={MYSQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=db_calonkaryawan;UID=root;Option="
Adodc1.RecordSource = "tbl_pelamar"
Adodc1.RecordSource = "select * from tbl_pelamar"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Atur_Grid()
DataGrid1.Columns(0).Caption = "KODE PELAMAR"
DataGrid1.Columns(1).Caption = "NAMA PELAMAR"
DataGrid1.Columns(2).Caption = "TEMPAT, TANGGAL LAHIR"
DataGrid1.Columns(3).Caption = "ALAMAT"
DataGrid1.Columns(4).Caption = "TELPON"
DataGrid1.Columns(5).Caption = "JENIS KELAMIN"
DataGrid1.Columns(6).Caption = "STATUS"
DataGrid1.Columns(7).Caption = "PENDIDIKAN"
DataGrid1.Columns(8).Caption = "IPK"
DataGrid1.Columns(9).Caption = "TANGGAL MELAMAR"
DataGrid1.Columns(0).Width = 2500
DataGrid1.Columns(1).Width = 5000
DataGrid1.Columns(2).Width = 3000
DataGrid1.Columns(3).Width = 3000
DataGrid1.Columns(4).Width = 3000
DataGrid1.Columns(5).Width = 3000
DataGrid1.Columns(6).Width = 3000
DataGrid1.Columns(7).Width = 3000
DataGrid1.Columns(8).Width = 3000
DataGrid1.Columns(9).Width = 3000
End Sub

Private Sub Segarkan()
Call bukadb
Call Tampilkan_pelamar
Set DataGrid1.DataSource = Adodc1
With DataGrid1
End With
Call Atur_Grid
End Sub



Private Sub Form_Load()
Call bukadb
End Sub



