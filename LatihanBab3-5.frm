VERSION 5.00
Begin VB.Form fBAB35 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Menu Utama"
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   3735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   9600
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Width           =   4815
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&tampilkan Daftar Kontrol"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   5040
      Width           =   2535
   End
End
Attribute VB_Name = "FBab35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For Each Control In Form1.Controls
 List1.AddItem Control.Name
 Next Control
End Sub

Private Sub Command2_Click()
fUTAMA.Enabled = True
FBab35.Visible = False
FBab35.Enabled = False


End Sub
