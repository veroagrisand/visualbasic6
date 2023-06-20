VERSION 5.00
Begin VB.Form fBAB34 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Menu Utama"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Hitung"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Lagi"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Bilangan Rata-rata"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Bilangan1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "fBAB34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bilangan
Dim jumlah As Currency
Dim banyakbil, indeks As Integer

Sub inputbil()
    bilangan = InputBox("Bilangan ke-" & indeks, "Data")
    If bilangan = "" Then
        MsgBox "Anda belum memasukkan bilangan ke-" & indeks, , "Pesan"
        inputbil
    End If
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    banyakbil = Text1.Text
    jumlah = 0

    For indeks = 1 To banyakbil
        inputbil
        jumlah = jumlah + bilangan
    Next

    Text2.Text = jumlah / banyakbil
    Command1.SetFocus
End Sub

Private Sub Command3_Click()
fUTAMA.Enabled = True
fBAB34.Visible = False
fBAB34.Enabled = False
End Sub
