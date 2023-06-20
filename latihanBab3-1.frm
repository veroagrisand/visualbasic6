VERSION 5.00
Begin VB.Form fBAB31 
   Caption         =   "If...Then...Else"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Menu Utama"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hitung"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Hasil Perhitungan"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Bilangan Kedua"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Bilangan Pertama"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "fBAB31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bilangan, pangkat As Integer
Function hasilpangkat(bilangan, pangkat)
Dim i As Integer
Dim B As Double
B = 1
For i = 1 To Abs(pangkat)
 B = B * bilangan
Next
hasilpangkat = B
End Function
Private Sub Command1_Click()
If Text1.Text <> "" And Text2.Text <> "" Then
bilangan = Text1.Text
pangkek = Text2.Text
If pangkat = 0 Then
 Text3.Text = 1
 ElseIf pangkek = 1 Then:
 Text3.Text = bilangan
 ElseIf pangkek > 0 Then:
 Text3.Text = hasilpangkat(bilangan, pangkek)
 Else:
 Text3.Text = 1 / hasilpangkat(bilangan, pangkek)
End If
Else:
pesan = MsgBox("Maaf isian tidak boleh kosong", vbInformation)
End If
End Sub



Private Sub Command2_Click()
fUTAMA.Enabled = True
fBAB31.Visible = False
fBAB31.Enabled = False

End Sub
