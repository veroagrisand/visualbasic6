VERSION 5.00
Begin VB.Form fUTAMA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proyek Akhir Matakuliah PBO Semester Genap 2022/2023"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17355
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   17355
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Aplikasi Sederhana Kasir Minimarket"
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "Proyek Akhir.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Menu mBerkas 
      Caption         =   "Berkas"
      Begin VB.Menu smTugas 
         Caption         =   "Tugas-Tugas"
         Begin VB.Menu smT36 
            Caption         =   "Tugas Bab 3-6"
         End
         Begin VB.Menu smT 
            Caption         =   "Tugas-----"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu smLatihan 
         Caption         =   "Latihan - Latihan"
         Begin VB.Menu smL31 
            Caption         =   "Latihan Bab 3-1"
         End
         Begin VB.Menu smL32 
            Caption         =   "Latihan Bab 3-2"
         End
         Begin VB.Menu smL33 
            Caption         =   "Latihan Bab 3-3"
         End
         Begin VB.Menu smL34 
            Caption         =   "Latihan Bab 3-4"
         End
         Begin VB.Menu smL35 
            Caption         =   "Latihan Bab 3-5"
         End
         Begin VB.Menu smL36 
            Caption         =   "Latihan Bab 3-6"
         End
         Begin VB.Menu smMS 
            Caption         =   "Multimedia Sederhana"
         End
         Begin VB.Menu smBasisdata 
            Caption         =   "Basis Data"
         End
         Begin VB.Menu sm0 
            Caption         =   "-------------------"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu smKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mUjian 
      Caption         =   "Ujian"
      Begin VB.Menu smUTS 
         Caption         =   "Ujian Tengah Semester"
      End
      Begin VB.Menu smUAS 
         Caption         =   "Ujan Akhir Semester"
      End
   End
   Begin VB.Menu mPA 
      Caption         =   "Proyek Akhir"
      Begin VB.Menu smAKM 
         Caption         =   "Aplikasi Kasir Minimarket"
      End
      Begin VB.Menu smHelp 
         Caption         =   "Bantuan"
      End
   End
End
Attribute VB_Name = "fUTAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSFlexGrid1_Click()

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub smBasisdata_Click()
fBasisdata.Enabled = True
fBasisdata.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub smKeluar_Click()
tanya = MsgBox("Yakin Ingin Menutup Aplikasi [Y/N]?", vbYesNo)
If tanya = vbYes Then
    End
End If
End Sub

Private Sub smL31_Click()
fBAB31.Enabled = True
fBAB31.Visible = True
fUTAMA.Enabled = False

End Sub

Private Sub smL32_Click()
fBAB32.Enabled = True
fBAB32.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub smL33_Click()
FBab33.Enabled = True
FBab33.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub smL34_Click()
fBAB34.Enabled = True
fBAB34.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub smL35_Click()
FBab35.Enabled = True
FBab35.Visible = True
fUTAMA.Enabled = False

End Sub

Private Sub smL36_Click()
fBAB36.Enabled = True
fBAB36.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub smMS_Click()
fMULTIMEDIA.Enabled = True
fMULTIMEDIA.Visible = True
fUTAMA.Enabled = False
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Format(Now(), "HH:mm:ss")
End Sub
