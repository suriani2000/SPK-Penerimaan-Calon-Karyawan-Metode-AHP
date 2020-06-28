VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form6"
   ScaleHeight     =   6075
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6000
      Top             =   5160
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Height          =   975
      Left            =   5520
      TabIndex        =   10
      Top             =   3720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1720
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
   Begin VB.CommandButton CMDKELUAR 
      Caption         =   "KELUAR"
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton CMDHAPUS 
      Caption         =   "HAPUS"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton CMDSIMPAN 
      Caption         =   "SIMPAN"
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "DITOLAK"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "DITERIMA"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "BATASAN"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
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
    MsgBox "Id Tidak Boleh Kosong", vbInformation, "Informasi"
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
koneksiDB.Execute "delete from batasan where id='" & Text1.Text & "'"
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
    MsgBox "ID tidak boleh kosong", vbInformation, "Informasi"
    Text1.SetFocus
Else
Set RsCari = New ADODB.Recordset
RsCari.Open "select * from batasan where id='" & Text1.Text & "'", koneksiDB
    If Not RsCari.EOF Then
        MsgBox "Data " & Text1 & " Sudah Ada", vbCritical, "Pesan"
        Text1.Text = ""
        Text1.SetFocus
    Else
        koneksiDB.Execute "insert into batasan (id,ditolak,diterima) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')"
        Call Hapus_Teks
        MsgBox "Data Tersimpan", vbInformation, "Pesan"
        Call Segarkan
    End If
End If
End Sub

Private Sub Simpan_Ulang()
If Trim(Text1.Text) = "" Then
    MsgBox "ID tidak boleh kosong", vbInformation, "Informasi"
    Text1.SetFocus
Else
koneksiDB.Execute " update batasan set diterima ='" & Text2.Text & "',ditolak ='" & Text3.Text & "',where id='" & Text1.Text & "'"
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
CMDSIMPAN.Caption = "Update"
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Tampilkan_batasan
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

Sub Tampilkan_batasan()
Call bukadb
Adodc1.ConnectionString = "DRIVER={MYSQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=db_calonkaryawan;UID=root;Option="
Adodc1.RecordSource = "batasan"
Adodc1.RecordSource = "select * from batasan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Atur_Grid()
DataGrid1.Columns(0).Caption = "ID BATASAN"
DataGrid1.Columns(1).Caption = "DITERIMA"
DataGrid1.Columns(2).Caption = "DITOLAK"
DataGrid1.Columns(0).Width = 2500
DataGrid1.Columns(1).Width = 5000
DataGrid1.Columns(2).Width = 3000
End Sub

Private Sub Segarkan()
Call bukadb
Call Tampilkan_batasan
Set DataGrid1.DataSource = Adodc1
With DataGrid1
End With
Call Atur_Grid
End Sub



Private Sub Form_Load()
Call bukadb
End Sub





