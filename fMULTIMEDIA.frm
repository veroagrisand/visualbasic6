VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMULTIMEDIA 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9840
   DrawMode        =   0  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
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
