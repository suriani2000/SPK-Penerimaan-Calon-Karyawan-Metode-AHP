VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form12"
   ScaleHeight     =   5055
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "INPUT KRITERIA"
      Height          =   855
      Left            =   7680
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INPUT DATA PENILAIAN"
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HASIL PENILAIAN"
      Height          =   855
      Left            =   7680
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INPUT PELAMAR"
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "PENILAIAN CALON KARAYAWAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show

End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub Command4_Click()
Form3.Show

End Sub
