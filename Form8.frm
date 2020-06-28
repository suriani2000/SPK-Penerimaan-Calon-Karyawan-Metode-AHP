VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form8"
   ScaleHeight     =   6045
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7435
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
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   0
      Top             =   4920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   794
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
   Begin VB.Label Label2 
      Caption         =   "Cari"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCari As New ADODB.Recordset

Private Sub Form_Activate()
Tampilkan_Nilai
Atur_Grid
End Sub

Private Sub Form_Load()
Call bukadb
End Sub

Sub Tampilkan_Nilai()
Call bukadb
Adodc1.ConnectionString = "DRIVER={MYSQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=db_calonkaryawan;UID=root;Option="
Adodc1.RecordSource = "view_penilaian"
Adodc1.RecordSource = "select * from view_penilaian"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Atur_Grid()
DataGrid1.Columns(0).Caption = "KD Pelamar"
DataGrid1.Columns(1).Caption = "Nama Pelamar"
DataGrid1.Columns(2).Caption = "Tanggal Penilaian"
DataGrid1.Columns(3).Caption = "K1"
DataGrid1.Columns(4).Caption = "K2"
DataGrid1.Columns(5).Caption = "K3"
DataGrid1.Columns(6).Caption = "K4"

DataGrid1.Columns(0).Width = 3000
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 3000
DataGrid1.Columns(4).Width = 2000
DataGrid1.Columns(5).Width = 3000
DataGrid1.Columns(6).Width = 3000

End Sub

Private Sub TXTCARI_Change()
Adodc1.RecordSource = "select * from view_penilaian where kd_pelamar like '%" & TXTCARI & "%' or nm_karyawan like '%" & TXTCARI & "%'"
Adodc1.Refresh
Atur_Grid
End Sub


