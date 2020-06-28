VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form9"
   ScaleHeight     =   3945
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   3000
   End
   Begin VB.TextBox TXTCARI 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   3000
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA PENILAIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label LabelJumlah 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Cari"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCari As New ADODB.Recordset

Private Sub DataGrid1_DblClick()
With Adodc1.Recordset
    If Form2.Visible = True Then
        Form2.Text4.Text = !k1
        Form2.Text5.Text = !k2
        Form2.Text6.Text = !k3
        Form2.Text7.Text = !k4
    End If
End With
Unload Me
End Sub

Private Sub Form_Activate()
Tampilkan_penilaian
Atur_Grid
End Sub

Private Sub Form_Load()
Call bukadb
End Sub

Sub Tampilkan_penilaian()
Call bukadb
Adodc1.ConnectionString = "DRIVER={MYSQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=db_calonkaryawan;UID=root;Option="
Adodc1.RecordSource = "tb_penilaian"
Adodc1.RecordSource = "select * from tbL_penilaian"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Atur_Grid()
DataGrid1.Columns(0).Caption = "ID PENILAIAN"
DataGrid1.Columns(1).Caption = "KD PELAMAR"
DataGrid1.Columns(2).Caption = "TANGGAL PENILAIAN"
DataGrid1.Columns(3).Caption = "K1"
DataGrid1.Columns(4).Caption = "K2"
DataGrid1.Columns(5).Caption = "K3"
DataGrid1.Columns(6).Caption = "K4"
DataGrid1.Columns(0).Width = 2500
DataGrid1.Columns(1).Width = 5000
DataGrid1.Columns(2).Width = 3000
DataGrid1.Columns(3).Width = 3000
DataGrid1.Columns(4).Width = 3000
DataGrid1.Columns(5).Width = 3000
DataGrid1.Columns(6).Width = 3000
End Sub

Private Sub Segarkan()
Call bukadb
Call Tampilkan_penilaian
Set DataGrid1.DataSource = Adodc1
With DataGrid1
End With
Call Atur_Grid
End Sub

Private Sub Timer1_Timer()
LabelJumlah.Caption = "Jumlah Penilaian : " & Adodc1.Recordset.RecordCount
End Sub

Private Sub TXTCARI_Change()
Adodc1.RecordSource = "select * from tbl_penilaian where idpenilaian like '%" & TXTCARI & "%'"
Adodc1.Refresh
Atur_Grid
End Sub




