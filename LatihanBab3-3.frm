VERSION 5.00
Begin VB.Form FBab33 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11130
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Menu Utama"
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hitung"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
End
Attribute VB_Name = "FBab33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim batas As String
Dim bil1, bil2, bil3 As Integer
Sub inputbatas()
 batas = InputBox("Batas barisan bilangan Finobuci", "Bilangan Finobuci")
 If batas = "" Then
 MsgBox "Anda belum mengisi batas bilangan", , "Pesan"
inputbatas
 End If
End Sub
Private Sub Command1_Click()
Cls
inputbatas
If CInt(batas) < 1 Then
 Print "Tidak ada bilangan finobuci kurang dari 1"
Else
bil1 = 1
bil2 = 0
bil3 = 1
Do Until bil1 >= CInt(batas)
 Print bil1
 bil1 = bil2 + bil3
 bil2 = bil3
 bil3 = bil1
Loop
End If
End Sub

Private Sub Command2_Click()
FUtama.Enabled = True
FBab33.Visible = False
FBab33.Enabled = False


End Sub
