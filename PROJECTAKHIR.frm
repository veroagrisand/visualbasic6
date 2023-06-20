VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMULTIMEDIA 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2990
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buka File"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   5160
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Menu Utama"
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   5040
      Width           =   2055
   End
End
Attribute VB_Name = "fMULTIMEDIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
fUTAMA.Enabled = True
fMULTIMEDIA.Visible = False
fMULTIMEDIA.Enabled = False

End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "File Video (*.avi) |*.avi| File Wav(*.wav)|*.wav| file Sequencer (*.mid)|*.mid| File MP3(*.mp3) |*.mp3|"
CommonDialog1.ShowOpen
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.FileName = CommonDialog1.FileName
MMControl1.Command = "Open"
End Sub
