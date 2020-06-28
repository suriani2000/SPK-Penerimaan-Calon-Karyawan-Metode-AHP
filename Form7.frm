VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form7"
   ScaleHeight     =   4590
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   3600
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4048
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
   Begin VB.TextBox TXTCARI 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Cari"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCari As New ADODB.Recordset

Private Sub DataGrid1_DblClick()
With Adodc1.Recordset
    If Form3.Visible = True Then
        Form3.Text2.Text = !kd_pelamar
        Form3.Text3.Text = !nm_karyawan
    End If
End With
Unload Me
End Sub

Private Sub Form_Activate()
Tampilkan_pelamar
Atur_Grid
End Sub

Private Sub Form_Load()
Call bukadb
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



Private Sub Timer1_Timer()
Label2.Caption = "Jumlah Penyakit : " & Adodc1.Recordset.RecordCount
End Sub

Private Sub TXTCARI_Change()
Adodc1.RecordSource = "select * from tbl_pelamar where kd_pelamar like '%" & TXTCARI & "%' or nm_karyawan like '%" & TXTCARI & "%'"
Adodc1.Refresh
Atur_Grid
End Sub




