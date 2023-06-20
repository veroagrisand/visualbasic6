VERSION 5.00
Begin VB.Form fBAB32 
   Caption         =   "Select...Case"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Menu Utama"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Hitung"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Lagi"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Hasil Konversi"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Suhu"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "fBAB32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pilihan As String
Sub inputpilihan()
 pilihan = InputBox("Konversi yang diinginkan" & Chr(10) & _
 "1. Celcius ke Fahreinheit" & Chr(10) & _
"2. Fahreinheit ke Reamur" & Chr(10) & _
"3. Reamur ke Celcius" & Chr(10) & _
"4. Celcius ke Kelvin" & Chr(10) & _
"Pilihan Anda :", "Pilihan Konversi")
If pilihan = "" Then
 MsgBox "Ulangi Pilihan Anda", , "Pesan"
 inputpilihan
End If
End Sub
Sub mengolahdata()
Select Case pilihan
 Case 1
 Text2.Text = ((9 / 5) * Text1.Text) - 32
 Case 2
 Text2.Text = (Text1.Text - 32) * (5 / 9)
 Case 3
 Text2.Text = ((5 / 4) * Text1.Text)
 Case 4
 Text2.Text = (Text1.Text + 273)
 Case Else
 Text2.Text = "Pilihan Tidak tersedia"
End Select
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub
Private Sub Command2_Click()
If Text1.Text = "" Then
 MsgBox "Anda belum mengisi kotak nilai", , "Pesan"
 Else
 inputpilihan
 mengolahdata
 Command2.SetFocus
End If
End Sub

Private Sub Command3_Click()
fUTAMA.Enabled = True
fBAB32.Visible = False
fBAB32.Enabled = False

End Sub
