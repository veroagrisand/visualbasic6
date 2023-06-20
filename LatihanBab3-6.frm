VERSION 5.00
Begin VB.Form fBAB36 
   Caption         =   "Aplikasi sederhana dengan method dan eventdrivent"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "&Menu Utama"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   4920
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MULAI"
      Height          =   2295
      Left            =   3840
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8400
      Top             =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bersihkan"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   4200
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   2880
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   2
      Text            =   "Budi"
      Top             =   1560
      Width           =   8175
   End
   Begin VB.Label Label4 
      Caption         =   "Jam Sekarang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   8
      Top             =   5640
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Program Sederhana Method dan Eventdrivent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   9735
   End
End
Attribute VB_Name = "fBAB36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem (Combo1.Text)
End Sub

Private Sub Command2_Click()
List1.RemoveItem (List1.ListIndex)
List1.Refresh
End Sub

Private Sub Command3_Click()
List1.Clear
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Mulai" Then
    Combo1.Enabled = True
    List1.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Caption = "Selesai"
    Else
    Command4.Caption = "Mulai"
    Combo1.Enabled = False
    List1.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
End Sub

Private Sub Command5_Click()
fUTAMA.Enabled = True
fBAB36.Visible = False
fBAB36.Enabled = False
End Sub

Private Sub Form_Load()
Combo1.AddItem ("Budi")
Combo1.AddItem ("Ucup")
Combo1.AddItem ("Amaik")
Combo1.AddItem ("Daper")
Combo1.AddItem ("Laduak")
Combo1.AddItem ("Gigi")
End Sub

Private Sub List1_Click()
Label3.Caption = List1.ListIndex
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Format(Now(), "dd-MM-YYYY / HH:mm:ss")
End Sub
