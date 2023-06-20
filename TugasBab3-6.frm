VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   3000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pengolahan Nilai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   8775
      Begin VB.CommandButton Command3 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Ulang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Proses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text4 
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
         Left            =   2760
         TabIndex        =   9
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox Text3 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text2 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Yudisium"
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
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Nilai Huruf"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Nilai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data siswa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   13
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor BP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   5040
      TabIndex        =   14
      Top             =   3000
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
              If Text1.Text <> "" And Text2.Text <> "" Then
        If Text2.Text >= 80 And Text2.Text <= 100 Then
                      Text3.Text = "A"
                      Text4.Text = "SANGAT MEMUASKAN"
        ElseIf Text2.Text >= 65 And Text2.Text <= 79 Then
                      Text3.Text = "B"
                      Text4.Text = "MEMUASKAN"
        ElseIf Text2.Text >= 55 And Text2.Text <= 64 Then
                      Text3.Text = "C"
                      Text4.Text = "CUKUP"
        ElseIf Text2.Text >= 45 And Text2.Text <= 54 Then
                      Text3.Text = "D"
                      Text4.Text = "KURANG"
        ElseIf Text2.Text >= 0 And Text2.Text <= 44 Then
                      Text3.Text = "E"
                      Text4.Text = "GAGAL"
        ElseIf Text2.Text > 100 Then
            notif = MsgBox("Maaf Maksimum Nilai adalah 100", vbOKOnly)
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text2.SetFocus
        End If
    Else
        notif = MsgBox("Maaf Isian Masih Kosong", vbOKOnly)

    End If
End Sub
Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_Change()
    Label5.Caption = Len(Text1.Text)
    Label5.Caption = Label5.Caption + " Karakter"
        If Label5.Caption = "10 Karakter" Then
            Text2.Enabled = True
        Else
            Text2.Enabled = False
        End If

End Sub

Private Sub Timer1_Timer()
Label6.Caption = Format(Now(), "dd-MM-YYYY / HH:mm:ss")
End Sub
